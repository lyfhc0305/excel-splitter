from __future__ import annotations

import argparse
import json
import math
import os
import queue
import threading
import sys
import warnings
import shutil
import subprocess
import tempfile
from contextlib import contextmanager
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl
import tkinter as tk
from openpyxl.utils import get_column_letter
from tkinter import filedialog, font as tkfont, messagebox, ttk
from split_core import (
    DEFAULT_NAME_TEMPLATE, SplitCancelled, normalize_group_value, safe_file_name, collect_groups,
    collect_group_rows, collect_group_columns, validate_parameters, save_groups, build_target_sheet,
    build_target_sheet_by_columns, rebuild_formula, rebuild_formula_by_columns, plan_file_names,
    plan_sheet_titles, select_groups, validate_name_template, occupied_output_names,
    unique_name,
)


OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
CONVERTIBLE_EXTENSIONS = {".xls", ".et"}
KEY_CHOICE_LIMIT = 300


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="按关键列或关键行拆分 Excel，并保留原始表格样式和格式。"
    )
    parser.add_argument("--input", help="待拆分的 Excel 文件路径")
    parser.add_argument("--sheet", help="要拆分的工作表名称，不传则默认第一个工作表")
    parser.add_argument(
        "--mode",
        choices=["column", "row"],
        default="column",
        help="拆分方式：column=按关键列拆分数据行（默认），row=按关键行拆分数据列",
    )
    parser.add_argument("--header-rows", type=int, help="表头行数（column 模式）")
    parser.add_argument("--footer-rows", type=int, default=0, help="表尾行数（固定在每个文件末尾，column 模式），默认 0")
    parser.add_argument("--key-column", type=int, help="关键列序号，从 1 开始（column 模式）")
    parser.add_argument("--header-cols", type=int, help="左侧固定列数（row 模式）")
    parser.add_argument("--footer-cols", type=int, default=0, help="右侧固定列数（固定在每个文件右侧，row 模式），默认 0")
    parser.add_argument("--key-row", type=int, help="关键行序号，从 1 开始（row 模式）")
    parser.add_argument("--output-dir", help="输出目录，默认在源文件同级目录生成 split_output")
    parser.add_argument("--group", action="append", metavar="关键字", help="只输出该关键字的拆分对象，可重复使用；默认输出全部")
    parser.add_argument(
        "--name-template",
        help=f"输出文件命名规则，默认 {DEFAULT_NAME_TEMPLATE}；可用 {{文件名}}、{{关键字}}、{{序号}}、{{工作表}}",
    )
    parser.add_argument("--single-workbook", action="store_true", help="合并输出为一个工作簿，每个拆分对象一个工作表")
    parser.add_argument("--list-groups", action="store_true", help="只列出拆分对象及其行列数，不生成文件")
    args = parser.parse_args(argv)
    supplied = sys.argv[1:] if argv is None else argv
    if supplied:
        if not args.input:
            parser.error("命令行运行必须提供 --input。")
        required = ("header_cols", "key_row") if args.mode == "row" else ("header_rows", "key_column")
        for name in required:
            if getattr(args, name) is None:
                parser.error(f"缺少参数 --{name.replace('_', '-')}。")
        if args.name_template is not None:
            if args.single_workbook:
                parser.error("--name-template 只用于每个拆分对象一个文件的输出，不能与 --single-workbook 同时使用。")
            try:
                validate_name_template(args.name_template)
            except ValueError as exc:
                parser.error(str(exc))
    return args


def load_workbook_compatible(input_path: Path) -> Tuple[openpyxl.Workbook, Optional[tempfile.TemporaryDirectory]]:
    input_path = Path(input_path)
    if not input_path.is_file():
        raise ValueError(f"找不到输入文件：{input_path}")
    suffix = input_path.suffix.lower()
    if suffix in OPENPYXL_EXTENSIONS:
        return read_workbook(input_path, keep_vba=suffix in {".xlsm", ".xltm"}), None

    if suffix == ".et":
        from zipfile import is_zipfile
        if is_zipfile(input_path):
            with input_path.open("rb") as file_handle:
                return read_workbook(file_handle, keep_vba=True), None

    if suffix in CONVERTIBLE_EXTENSIONS:
        converted_path, temp_dir = convert_to_xlsx(input_path)
        try:
            return read_workbook(converted_path), temp_dir
        except Exception:
            temp_dir.cleanup()
            raise

    raise ValueError(f"暂不支持该文件格式：{suffix}")


# openpyxl warns that worksheet extensions "will be removed". Those extensions
# (unknown extLst, x14 data validation, conditional formatting, slicers) are not
# the cells being split; refusing the file blocks ordinary workbooks. Images,
# charts and shapes are different: openpyxl drops them, and validate_features
# treats the same content as fatal when it can still see it.
_FATAL_LOAD_MARKERS = (
    "shapes and drawings will be lost",
    "will be removed because it cannot be read",
    "image format is not supported",
    "unable to read chart",
)
_VBA_STREAM = "_VBA_PROJECT_CUR".encode("utf-16le")


def classify_load_warnings(messages):
    """Split load warnings into a user-visible notice or a hard failure."""
    notices, fatal = [], []
    for message in messages:
        text = str(message)
        lowered = text.casefold()
        if "extension is not supported and will be removed" in lowered:
            notices.append(f"文件含无法保留的扩展（{text}），已忽略并继续拆分。")
            continue
        if any(marker in lowered for marker in _FATAL_LOAD_MARKERS):
            fatal.append(text)
    return notices, fatal


def legacy_workbook_has_vba(path) -> bool:
    """OLE .xls / legacy .et store macros in the ``_VBA_PROJECT_CUR`` stream."""
    try:
        data = Path(path).read_bytes()
    except OSError:
        return False
    return _VBA_STREAM in data


def read_workbook(path, **options):
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always", UserWarning)
        workbook = openpyxl.load_workbook(path, data_only=False, read_only=False, rich_text=True, **options)
    notices, fatal = classify_load_warnings(str(item.message) for item in captured)
    if fatal:
        close_workbook_compatible(workbook, None)
        if any("shapes and drawings will be lost" in item.casefold() for item in fatal):
            raise ValueError("此工作表含有尚不能完整保留的内容：形状或图形。请先在副本中删除这些内容后拆分。")
        raise ValueError("文件包含无法完整保留的功能：" + "；".join(fatal))
    workbook._splitter_notices = notices
    return workbook


def select_sheet(workbook, name=None):
    if not workbook.worksheets:
        raise ValueError("文件中没有可拆分的数据工作表。")
    if name is None:
        return workbook.worksheets[0]
    for sheet in workbook.worksheets:
        if sheet.title == name:
            return sheet
    raise ValueError(f"找不到工作表：{name}")


def close_workbook_compatible(workbook: openpyxl.Workbook, temp_dir: Optional[tempfile.TemporaryDirectory]) -> None:
    try:
        workbook.close()
        if workbook.vba_archive is not None:
            workbook.vba_archive.close()
    finally:
        if temp_dir:
            temp_dir.cleanup()


def convert_to_xlsx(input_path: Path) -> Tuple[Path, tempfile.TemporaryDirectory]:
    temp_dir = tempfile.TemporaryDirectory()
    temp_path = Path(temp_dir.name)

    # LibreOffice is tried first and returns on success, so the HasVBProject
    # check inside the Windows COM script never runs when soffice is installed.
    # Detect the macro stream here and fail instead of silently dropping VBA.
    from zipfile import is_zipfile
    if input_path.suffix.lower() in {".xls", ".et"} and not is_zipfile(input_path) and legacy_workbook_has_vba(input_path):
        temp_dir.cleanup()
        raise ValueError("文件包含 VBA 宏，转换为 xlsx 会丢失宏。请先使用不含宏的副本。")

    try:
        try:
            converted_path = convert_with_libreoffice(input_path, temp_path)
        except (OSError, subprocess.SubprocessError):
            converted_path = None
        if converted_path:
            return converted_path, temp_dir

        converted_path = convert_with_windows_com(input_path, temp_path)
        if converted_path:
            return converted_path, temp_dir
    except Exception:
        temp_dir.cleanup()
        raise

    temp_dir.cleanup()
    raise ValueError(
        "无法自动转换该文件。请安装 LibreOffice，或在 Windows 上安装 Excel/WPS 后重试。"
    )


def convert_with_libreoffice(input_path: Path, output_dir: Path) -> Optional[Path]:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if not executable:
        return None

    result = subprocess.run(
        [
            executable,
            f"-env:UserInstallation={(output_dir / 'lo-profile').resolve().as_uri()}",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(input_path.resolve()),
        ],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )
    if result.returncode != 0:
        return None

    expected_path = output_dir / f"{input_path.stem}.xlsx"
    if expected_path.exists():
        return expected_path

    converted_files = list(output_dir.glob("*.xlsx"))
    return converted_files[0] if converted_files else None


def convert_with_windows_com(input_path: Path, output_dir: Path) -> Optional[Path]:
    if not shutil.which("powershell"):
        return None

    output_path = output_dir / f"{input_path.stem}.xlsx"
    script_path = output_dir / "convert_spreadsheet.ps1"
    script_path.write_text(
        r"""
param(
    [string]$InputPath,
    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$progIds = @("Excel.Application", "Ket.Application", "ET.Application")
$app = $null

foreach ($progId in $progIds) {
    try {
        $app = New-Object -ComObject $progId
        break
    } catch {
        $app = $null
    }
}

if ($null -eq $app) {
    throw "未找到可用的 Excel/WPS COM 转换器"
}

$app.Visible = $false
$app.DisplayAlerts = $false
$workbook = $null

try {
    $app.AutomationSecurity = 3
    $workbook = $app.Workbooks.Open($InputPath, 0, $true)
    if ($workbook.HasVBProject) {
        throw "文件包含 VBA 宏，转换为 xlsx 会丢失宏。请先使用不含宏的副本。"
    }
    $workbook.SaveAs($OutputPath, 51)
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    $app.Quit()
}
""",
        encoding="utf-8-sig",
    )
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
            str(input_path.resolve()),
            str(output_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )
    if result.returncode != 0 or not output_path.exists():
        return None
    return output_path



class WorkbookCache:
    """Keeps the last workbook read by the GUI worker, so previews do not re-read the file.

    Only the worker thread uses it. A changed file size or modification time reloads the
    file, so edits saved in Excel/WPS are picked up by the next preview or split.
    """

    def __init__(self) -> None:
        self._key = None
        self._workbook = None
        self._temp_dir = None

    def get(self, input_path: Path) -> openpyxl.Workbook:
        input_path = Path(input_path)
        if not input_path.is_file():
            raise ValueError(f"找不到输入文件：{input_path}")
        stat = input_path.stat()
        key = (str(input_path.resolve()), stat.st_mtime_ns, stat.st_size)
        if key != self._key:
            self.clear()
            self._workbook, self._temp_dir = load_workbook_compatible(input_path)
            self._key = key
        return self._workbook

    def clear(self) -> None:
        workbook, temp_dir = self._workbook, self._temp_dir
        self._key = self._workbook = self._temp_dir = None
        if workbook is not None:
            close_workbook_compatible(workbook, temp_dir)


@contextmanager
def open_workbook(input_path, cache: Optional[WorkbookCache] = None):
    if cache is not None:
        yield cache.get(input_path)
        return
    workbook, temp_dir = load_workbook_compatible(Path(input_path))
    try:
        yield workbook
    finally:
        close_workbook_compatible(workbook, temp_dir)


def _split_workbook(input_path, sheet_name, leading, key, output_dir, trailing=0, by_columns=False, *,
                    groups=None, name_template=None, single_workbook=False, progress=None, cancel=None, cache=None):
    input_path = Path(input_path)
    with open_workbook(input_path, cache) as source_wb:
        source_ws = select_sheet(source_wb, sheet_name)
        selected = select_groups(collect_groups(source_ws, leading, key, trailing, by_columns), groups)
        return save_groups(source_wb, source_ws, selected, input_path, output_dir, leading, trailing, by_columns,
                           name_template, single_workbook, progress, cancel)


def split_workbook(input_path, sheet_name, header_rows, key_column, output_dir=None, footer_rows=0, **options):
    return _split_workbook(input_path, sheet_name, header_rows, key_column, output_dir, footer_rows, **options)


def split_workbook_by_row(input_path, sheet_name, header_cols, key_row, output_dir=None, footer_cols=0, **options):
    return _split_workbook(input_path, sheet_name, header_cols, key_row, output_dir, footer_cols, True, **options)


def list_groups(input_path, sheet_name, leading, key, trailing=0, by_columns=False, cache=None):
    """Return the sheet title and each group's row (or column) count without writing files."""
    with open_workbook(input_path, cache) as workbook:
        ws = select_sheet(workbook, sheet_name)
        groups = collect_groups(ws, leading, key, trailing, by_columns)
        return ws.title, {name: len(indices) for name, indices in groups.items()}


def preview_indices(total, *focus):
    indices = set(range(1, min(total, 20) + 1))
    indices.update(range(max(1, total - 2), total + 1))
    for value in focus:
        if value is not None:
            indices.update(range(max(1, value - 3), min(total, value + 4) + 1))
    return sorted(indices)


def sample_indices(start, end, head, tail=0):
    """``start..end`` keeping at most ``head`` leading and ``tail`` trailing indices."""
    if end < start:
        return []
    if end - start + 1 <= head + tail:
        return list(range(start, end + 1))
    return list(range(start, start + head)) + list(range(end - tail + 1, end + 1))


def cell_text(ws, row, column) -> str:
    return (normalize_group_value(ws.cell(row=row, column=column).value) or "").replace("\n", " ")


def column_label(column, text) -> str:
    letter = get_column_letter(column)
    return f"{letter} · {text}" if text else letter


def add_preview_rows(rows, ws, indices, kind, columns, section_end=None) -> None:
    """Append previewed rows, with a "⋯" row wherever part of the sheet is skipped."""
    def gap(count):
        return {"row_no": "⋯", "kind": "gap", "values": [f"省略 {count} 行"] + [""] * (len(columns) - 1)}

    previous = None
    for row in indices:
        if previous is not None and row > previous + 1:
            rows.append(gap(row - previous - 1))
        rows.append({"row_no": row, "kind": kind, "values": [cell_text(ws, row, column) for column in columns]})
        previous = row
    if section_end is not None and previous is not None and previous < section_end:
        rows.append(gap(section_end - previous))


def load_workbook_info(input_path: Path, cache: Optional[WorkbookCache] = None) -> Dict[str, object]:
    with open_workbook(input_path, cache) as workbook:
        sheet_names = [sheet.title for sheet in workbook.worksheets]
        first_sheet = select_sheet(workbook)
        column_preview = []
        preview_row = min(first_sheet.max_row, 6)
        for column in preview_indices(first_sheet.max_column):
            value = first_sheet.cell(preview_row, column).value
            display = normalize_group_value(value) or "(空)"
            column_preview.append(f"{column}. {get_column_letter(column)} - {display}")
        return {
            "sheet_names": sheet_names,
            "max_row": first_sheet.max_row,
            "max_column": first_sheet.max_column,
            "column_preview": column_preview,
        }


def build_sheet_preview(
    input_path: Path,
    sheet_name: str,
    header_rows: int,
    key_column: int,
    footer_rows: int = 0,
    cache: Optional[WorkbookCache] = None,
) -> Dict[str, object]:
    with open_workbook(input_path, cache) as workbook:
        ws = select_sheet(workbook, sheet_name)
        validate_parameters(ws, header_rows, key_column, footer_rows)
        max_row, max_column = ws.max_row, ws.max_column
        data_end = max_row - footer_rows
        label_row = max(1, min(max_row, header_rows))
        preview_columns = preview_indices(max_column, key_column)
        column_headers = [
            ("★ " if column == key_column else "") + column_label(column, cell_text(ws, label_row, column))
            for column in preview_columns
        ]

        preview_rows: List[Dict[str, object]] = []
        add_preview_rows(preview_rows, ws, sample_indices(1, header_rows, 6, 3), "header", preview_columns)
        add_preview_rows(preview_rows, ws, sample_indices(header_rows + 1, data_end, 20), "data", preview_columns, data_end)
        add_preview_rows(preview_rows, ws, sample_indices(data_end + 1, max_row, 3, 3), "footer", preview_columns)

        choices = sorted(set(range(1, min(max_column, KEY_CHOICE_LIMIT) + 1)) | {key_column})
        key_choices = [(column, column_label(column, cell_text(ws, label_row, column))) for column in choices]

        groups = collect_group_rows(ws, header_rows, key_column, footer_rows)
        split_objects = [{"name": name, "count": len(rows)} for name, rows in groups.items()]

        return {
            "sheet_title": ws.title,
            "max_row": max_row,
            "max_column": max_column,
            "preview_columns": preview_columns,
            "column_headers": column_headers,
            "key_position": preview_columns.index(key_column),
            "preview_rows": preview_rows,
            "key_choices": key_choices,
            "split_objects": split_objects,
            "group_count": len(split_objects),
            "blank_group": groups.blank,
            "data_count": data_end - header_rows,
        }


def build_sheet_preview_by_row(
    input_path: Path,
    sheet_name: str,
    header_cols: int,
    key_row: int,
    footer_cols: int = 0,
    cache: Optional[WorkbookCache] = None,
) -> Dict[str, object]:
    with open_workbook(input_path, cache) as workbook:
        ws = select_sheet(workbook, sheet_name)
        validate_parameters(ws, header_cols, key_row, footer_cols, True)
        max_row, max_column = ws.max_row, ws.max_column
        footer_start = max_column - footer_cols + 1 if footer_cols else None
        preview_columns = preview_indices(max_column, header_cols, footer_start)

        column_headers: List[str] = []
        for column in preview_columns:
            fixed = column <= header_cols or (footer_start is not None and column >= footer_start)
            column_headers.append(column_label(column, cell_text(ws, key_row, column)) + ("（固定）" if fixed else ""))

        preview_rows: List[Dict[str, object]] = []
        rows = set(sample_indices(1, max_row, 25, 3)) | set(range(max(1, key_row - 3), min(max_row, key_row + 3) + 1))
        add_preview_rows(preview_rows, ws, sorted(rows), "data", preview_columns)
        for row in preview_rows:
            if row["row_no"] == key_row:
                row["kind"] = "keyrow"

        key_choices = []
        for row in sorted(set(range(1, min(max_row, 50) + 1)) | {key_row}):
            samples = [text for text in (cell_text(ws, row, column) for column in preview_columns) if text][:3]
            key_choices.append((row, f"第 {row} 行 · {' | '.join(samples)}" if samples else f"第 {row} 行（空行）"))

        groups = collect_group_columns(ws, header_cols, key_row, footer_cols)
        split_objects = [{"name": name, "count": len(cols)} for name, cols in groups.items()]

        return {
            "sheet_title": ws.title,
            "max_row": max_row,
            "max_column": max_column,
            "preview_columns": preview_columns,
            "column_headers": column_headers,
            "key_position": None,
            "preview_rows": preview_rows,
            "key_choices": key_choices,
            "split_objects": split_objects,
            "group_count": len(split_objects),
            "blank_group": groups.blank,
            "data_count": max_column - header_cols - footer_cols,
        }


APP_TITLE = "Excel 拆表工具"

PALETTE = {
    "bg": "#F1F5F9",
    "card": "#FFFFFF",
    "field": "#FFFFFF",
    "border": "#E2E8F0",
    "border_strong": "#CBD5E1",
    "heading": "#F8FAFC",
    "hover": "#F1F5F9",
    "pressed": "#E2E8F0",
    "text": "#0F172A",
    "label": "#334155",
    "muted": "#64748B",
    "disabled": "#94A3B8",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_press": "#1E40AF",
    "accent_soft": "#EFF6FF",
    "accent_disabled": "#93C5FD",
    "success": "#16A34A",
    "warning": "#D97706",
    "danger": "#DC2626",
    "row_header": "#FEF3C7",
    "row_header_edge": "#F59E0B",
    "row_footer": "#DCFCE7",
    "row_footer_edge": "#22C55E",
    "row_key": "#FEE2E2",
    "row_key_edge": "#EF4444",
    "row_split": "#E0E7FF",
    "row_stripe": "#F8FAFC",
    "tooltip": "#1E293B",
}

FONT_CANDIDATES = {
    "win32": ("Microsoft YaHei UI", "Microsoft YaHei", "SimHei"),
    "darwin": ("PingFang SC", "Hiragino Sans GB", "Heiti SC", "STHeiti"),
    "linux": ("Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei", "WenQuanYi Zen Hei", "Droid Sans Fallback"),
}

SETTINGS_KEYS = (
    "mode", "header_rows", "footer_rows", "key_column", "header_cols", "footer_cols", "key_row",
    "output_mode", "name_template", "last_dir",
)

TIPS = {
    "header_rows": "每个输出文件顶部都会保留的固定行，如标题、列名。可为 0。",
    "footer_rows": "每个输出文件底部都会保留的固定行，如合计、签字栏。可为 0。",
    "key_column": "按这一列的内容分组：内容相同的数据行输出到同一个文件。也可以单击预览表格的列标题来选择。",
    "header_cols": "每个输出文件左侧都会保留的固定列，如项目名称。可为 0。",
    "footer_cols": "每个输出文件右侧都会保留的固定列，如合计列。可为 0。",
    "key_row": "按这一行的内容分组：内容相同的数据列输出到同一个文件。也可以单击预览表格中的行来选择。",
    "template": "可用占位符：{文件名} 源文件名，{关键字} 分组关键字，{序号} 按输出顺序编号，{工作表} 工作表名称。",
    "reload": "重新读取文件（F5）。在 Excel/WPS 中修改并保存源文件后使用。",
    "open": "在文件管理器中打开输出目录",
    "output": "选择源文件后默认输出到它旁边的 split_output 文件夹；手动修改后，切换源文件时不再自动更改。",
    "column_mode": "单击列标题设为关键列；右键单击行可设为表头或表尾的分界。",
    "row_mode": "单击某一行设为关键行；右键单击列标题可设为左右固定列的分界。",
    "notes": "✓ 只读取源文件，不会修改原表\n✓ 空白关键字单独成组，不会被遗漏\n✓ 可随时取消，失败不会留下半批文件",
}

HELP_TEXT = """使用步骤
1. 选择 Excel 文件和工作表。
2. 选择拆分方式，设置表头、表尾（或左右固定列）以及关键列（或关键行）。单击预览的列标题可设为关键列；右键单击预览中的行或列可快速设置分界。
3. 在“拆分对象”中勾选需要输出的分组，确认输出目录、输出方式和文件命名。
4. 点击“开始拆分”。拆分过程中可以取消，未完成的拆分不会留下任何输出文件。

文件命名占位符
{文件名} 源文件名　{关键字} 分组关键字
{序号} 按输出顺序编号　{工作表} 工作表名称

快捷键
Ctrl+O 打开文件　F5 重新读取
Ctrl+Enter 开始拆分　Esc 取消拆分

预览只显示部分行列，拆分时处理整张工作表。公式由 Excel/WPS 打开输出文件时重算。"""


def enable_dpi_awareness() -> None:
    """Render crisply on scaled Windows displays instead of being bitmap-stretched."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def pick_font_family(root: tk.Misc) -> str:
    available = set(tkfont.families(root))
    for family in FONT_CANDIDATES.get(sys.platform, FONT_CANDIDATES["linux"]):
        if family in available:
            return family
    return tkfont.nametofont("TkDefaultFont", root).actual("family")


def default_settings_path() -> Path:
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home())
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "ExcelSplitter" / "settings.json"


def load_settings(path: Path) -> Dict[str, str]:
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {key: value for key, value in data.items() if key in SETTINGS_KEYS and isinstance(value, str)}


def save_settings(path: Path, data: Dict[str, str]) -> None:
    # Remembering settings is a convenience; failing to write them never blocks the user.
    try:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(".tmp")
        temp_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        temp_path.replace(path)
    except OSError:
        pass


def open_path(path: Path) -> None:
    if sys.platform == "win32":
        os.startfile(str(path))
    else:
        opener = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([opener, str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _fill_rounded(image: tk.PhotoImage, size: int, radius: float, color: str, inset: int = 0) -> None:
    """Fill a rounded square, ``inset`` pixels inside a ``size`` x ``size`` image."""
    span = size - 2 * inset
    for y in range(span):
        dy = max(radius - y - 0.5, y + 0.5 - (span - radius), 0)
        cut = radius - math.sqrt(max(radius * radius - dy * dy, 0)) if dy else 0
        x0 = inset + int(round(cut))
        if size - x0 > x0:
            image.put(color, to=(x0, inset + y, size - x0, inset + y + 1))


def draw_checkbox(root: tk.Misc, size: int, state: str) -> tk.PhotoImage:
    """Checkbox glyph for list rows: ``on``, ``off`` or ``mixed``."""
    image = tk.PhotoImage(master=root, width=size, height=size)
    radius = max(2.0, size / 5)
    if state == "off":
        border = max(1, round(size / 14))
        _fill_rounded(image, size, radius, PALETTE["border_strong"])
        _fill_rounded(image, size, radius - border, PALETTE["field"], border)
        return image
    _fill_rounded(image, size, radius, PALETTE["accent"])
    stroke = max(2, round(size / 7))
    if state == "mixed":
        top = size // 2 - stroke // 2
        image.put("#FFFFFF", to=(round(size * 0.26), top, round(size * 0.74), top + stroke))
        return image
    points = [(0.25, 0.52), (0.43, 0.70), (0.76, 0.32)]
    for (x0, y0), (x1, y1) in zip(points, points[1:]):
        steps = max(4, size * 2)
        for step in range(steps + 1):
            t = step / steps
            left = int((x0 + (x1 - x0) * t) * size - stroke / 2)
            top = int((y0 + (y1 - y0) * t) * size - stroke / 2)
            image.put("#FFFFFF", to=(left, top, left + stroke, top + stroke))
    return image


def draw_app_icon(root: tk.Misc, size: int) -> tk.PhotoImage:
    """A small spreadsheet glyph used for the window icon and the header."""
    image = tk.PhotoImage(master=root, width=size, height=size)
    _fill_rounded(image, size, size * 0.22, PALETTE["accent"])

    def box(x0, y0, x1, y1, color):
        image.put(color, to=(round(x0 * size), round(y0 * size), max(round(x1 * size), round(x0 * size) + 1),
                             max(round(y1 * size), round(y0 * size) + 1)))

    line = 1 / size * max(1, round(size / 30))
    box(0.2, 0.24, 0.8, 0.76, "#FFFFFF")
    box(0.2, 0.24, 0.8, 0.39, "#BFDBFE")
    for y in (0.39, 0.52, 0.64):
        box(0.2, y, 0.8, y + line, PALETTE["accent"])
    box(0.45, 0.24, 0.45 + line, 0.76, PALETTE["accent"])
    return image


class Tooltip:
    """Hover hint shown after a short delay; hidden again on leave or click."""

    def __init__(self, widget: tk.Widget, text: str, app: "SplitterApp") -> None:
        self.widget, self.text, self.app = widget, text, app
        self.window = None
        self.after_id = None
        widget.bind("<Enter>", self.schedule, add="+")
        widget.bind("<Leave>", self.hide, add="+")
        widget.bind("<ButtonPress>", self.hide, add="+")

    def schedule(self, _event=None) -> None:
        self.cancel()
        self.after_id = self.widget.after(550, self.show)

    def cancel(self) -> None:
        if self.after_id is not None:
            self.widget.after_cancel(self.after_id)
            self.after_id = None

    def show(self) -> None:
        self.after_id = None
        if self.window is not None or self.app._closed:
            return
        px = self.app.px
        self.window = tk.Toplevel(self.widget)
        self.window.wm_overrideredirect(True)
        self.window.wm_geometry(
            f"+{self.widget.winfo_rootx() + px(4)}+{self.widget.winfo_rooty() + self.widget.winfo_height() + px(6)}"
        )
        tk.Label(
            self.window, text=self.text, justify="left", wraplength=px(300), font=self.app.fonts["small"],
            bg=PALETTE["tooltip"], fg="#F8FAFC", padx=px(9), pady=px(6),
        ).pack()

    def hide(self, _event=None) -> None:
        self.cancel()
        if self.window is not None:
            self.window.destroy()
            self.window = None


def run_cli(args: argparse.Namespace) -> None:
    by_columns = args.mode == "row"
    output_dir = Path(args.output_dir) if args.output_dir else None
    if by_columns:
        leading, key, trailing = args.header_cols, args.key_row, getattr(args, "footer_cols", 0) or 0
    else:
        leading, key, trailing = args.header_rows, args.key_column, getattr(args, "footer_rows", 0) or 0
    if getattr(args, "list_groups", False):
        title, counts = list_groups(Path(args.input), args.sheet, leading, key, trailing, by_columns)
        print(f"工作表 {title}：共 {len(counts)} 个拆分对象")
        print(f"关键字\t{'列' if by_columns else '行'}数")
        for name, count in counts.items():
            print(f"{name}\t{count}")
        return
    options = dict(
        groups=getattr(args, "group", None),
        name_template=getattr(args, "name_template", None),
        single_workbook=getattr(args, "single_workbook", False),
    )
    if by_columns:
        files = split_workbook_by_row(
            input_path=Path(args.input),
            sheet_name=args.sheet,
            header_cols=leading,
            key_row=key,
            output_dir=output_dir,
            footer_cols=trailing,
            **options,
        )
    else:
        files = split_workbook(
            input_path=Path(args.input),
            sheet_name=args.sheet,
            header_rows=leading,
            key_column=key,
            output_dir=output_dir,
            footer_rows=trailing,
            **options,
        )
    summary = "\n".join(str(path) for path in files)
    if options["single_workbook"]:
        print(f"拆分完成，共生成 1 个工作簿（{len(files.sheet_titles)} 个工作表）：\n{summary}")
    else:
        print(f"拆分完成，共生成 {len(files)} 个文件：\n{summary}")
    for warning in getattr(files, "warnings", []):
        print(f"公式提示：{warning}", file=sys.stderr)


def main() -> None:
    args = parse_args()
    if args.input:
        try:
            run_cli(args)
        except Exception as exc:
            print(f"拆分失败：{exc}", file=sys.stderr)
            raise SystemExit(1)
        return
    from splitter_app import SplitterApp
    SplitterApp().run()


if __name__ == "__main__":
    main()

from splitter_app import SplitterApp  # noqa: E402

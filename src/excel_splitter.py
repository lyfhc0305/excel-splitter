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
    plan_sheet_titles, select_groups, validate_name_template,
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


def read_workbook(path, **options):
    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always", UserWarning)
        workbook = openpyxl.load_workbook(path, data_only=False, read_only=False, rich_text=True, **options)
    lost_features = [str(item.message) for item in captured if "removed" in str(item.message) or "not supported" in str(item.message)]
    if lost_features:
        close_workbook_compatible(workbook, None)
        raise ValueError("文件包含无法完整保留的功能：" + "；".join(lost_features))
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


class SplitterApp:
    def __init__(self, settings_path: Optional[Path] = None) -> None:
        enable_dpi_awareness()
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.scale = min(3.0, max(1.0, self.root.winfo_fpixels("1i") / 96))
        self.settings_path = Path(settings_path) if settings_path else default_settings_path()
        settings = load_settings(self.settings_path)

        def number(key: str, default: str) -> str:
            value = settings.get(key, "")
            return value if value.isdigit() else default

        def choice(key: str, options: Tuple[str, ...]) -> str:
            return settings.get(key) if settings.get(key) in options else options[0]

        self.input_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.mode_var = tk.StringVar(value=choice("mode", ("column", "row")))
        self.header_rows_var = tk.StringVar(value=number("header_rows", "1"))
        self.footer_rows_var = tk.StringVar(value=number("footer_rows", "0"))
        self.key_column_var = tk.StringVar(value=number("key_column", "1"))
        self.header_cols_var = tk.StringVar(value=number("header_cols", "1"))
        self.footer_cols_var = tk.StringVar(value=number("footer_cols", "0"))
        self.key_row_var = tk.StringVar(value=number("key_row", "1"))
        self.output_var = tk.StringVar()
        self.output_mode_var = tk.StringVar(value=choice("output_mode", ("files", "workbook")))
        self.name_template_var = tk.StringVar(value=settings.get("name_template") or DEFAULT_NAME_TEMPLATE)
        self.status_var = tk.StringVar()
        self.last_dir = settings.get("last_dir", "")

        self.preview_column_map: Dict[str, int] = {}
        self.preview_row_map: Dict[str, int] = {}
        self.group_items: Dict[str, str] = {}
        self.group_counts: Dict[str, int] = {}
        self.groups: Optional[List[Dict[str, object]]] = None
        self.excluded = set()
        self._key_choices = {"column": [], "row": []}
        self._preview_info = None
        self._preview_failed = False
        self._naming_error = None
        self._output_auto = True
        self._setting_output = False
        self._loaded_path = ""
        self._disabled_widgets = []

        self._jobs = queue.Queue()
        self._results = queue.Queue()
        self._progress = queue.Queue()
        self._versions = {}
        self._cache = WorkbookCache()
        self._cancel = threading.Event()
        self._preview_after = None
        self._splitting = False
        self._closed = False
        self._close_requested = False
        self._place_window()
        self._build_ui()
        self.on_mode_changed(refresh=False)
        self._set_status("先选择 Excel 文件，界面会自动读取工作表并生成预览。")
        threading.Thread(target=self._worker, daemon=True).start()
        self._poll_after = self.root.after(75, self._poll_jobs)
        self.root.protocol("WM_DELETE_WINDOW", self._close)
        for variable in (self.header_rows_var, self.footer_rows_var, self.key_column_var,
                         self.header_cols_var, self.footer_cols_var, self.key_row_var):
            variable.trace_add("write", self.on_option_changed)
        self.name_template_var.trace_add("write", lambda *_args: self.update_group_names())
        self.output_var.trace_add("write", self._on_output_edited)
        self.input_var.trace_add("write", lambda *_args: self._update_start_state())
        self._bind_shortcuts()
        self.update_group_names()

    def px(self, value: float) -> int:
        return int(round(value * self.scale))

    def _place_window(self) -> None:
        px = self.px
        screen_w, screen_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        width, height = min(px(1180), screen_w - px(40)), min(px(760), screen_h - px(80))
        x, y = max(0, (screen_w - width) // 2), max(0, (screen_h - height) // 3)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(px(1000), width), min(px(640), height))

    # ---- 外观 ----

    def _setup_style(self) -> None:
        c, px = PALETTE, self.px
        family = pick_font_family(self.root)
        size = 13 if sys.platform == "darwin" else 10
        self.fonts = {
            "base": (family, size),
            "small": (family, size - 1),
            "bold": (family, size, "bold"),
            "card": (family, size + 1, "bold"),
            "title": (family, size + 5, "bold"),
            "button": (family, size + 1, "bold"),
        }
        self.table_font = tkfont.Font(root=self.root, family=family, size=size)
        self.heading_font = tkfont.Font(root=self.root, family=family, size=size, weight="bold")
        if sys.platform != "darwin":
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
                tkfont.nametofont(name, self.root).configure(family=family, size=size)
        self.root.option_add("*TCombobox*Listbox.font", self.fonts["base"])
        self.root.option_add("*TCombobox*Listbox.background", c["field"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", c["accent_soft"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", c["text"])
        self.menu_options = dict(
            tearoff=0, font=self.fonts["base"], bg=c["card"], fg=c["text"], activebackground=c["accent_soft"],
            activeforeground=c["accent"], relief="solid", borderwidth=1,
        )

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(
            ".", font=self.fonts["base"], background=c["bg"], foreground=c["text"], bordercolor=c["border"],
            lightcolor=c["border"], darkcolor=c["border"], troughcolor=c["bg"], focuscolor=c["accent"],
            selectbackground=c["accent_soft"], selectforeground=c["text"], insertcolor=c["text"],
        )
        style.configure("TFrame", background=c["bg"])
        style.configure("Card.TFrame", background=c["card"], bordercolor=c["border"], relief="solid", borderwidth=1)
        style.configure("CardBody.TFrame", background=c["card"])
        style.configure("Bar.TFrame", background=c["card"])
        style.configure("TLabel", background=c["card"], foreground=c["text"])
        style.configure("Title.TLabel", font=self.fonts["title"])
        style.configure("Subtitle.TLabel", foreground=c["muted"])
        style.configure("CardTitle.TLabel", font=self.fonts["card"])
        style.configure("Field.TLabel", foreground=c["label"])
        style.configure("Muted.TLabel", foreground=c["muted"], font=self.fonts["small"])
        style.configure("Error.TLabel", foreground=c["danger"], font=self.fonts["small"])
        style.configure("Note.TLabel", foreground=c["success"], font=self.fonts["small"])
        style.configure("Empty.TLabel", foreground=c["muted"], font=self.fonts["base"])
        style.configure("EmptyError.TLabel", foreground=c["danger"], font=self.fonts["base"])
        style.configure("Status.TLabel", foreground=c["label"])
        for kind, color in (("Idle", c["disabled"]), ("Busy", c["accent"]), ("Success", c["success"]),
                            ("Warning", c["warning"]), ("Error", c["danger"])):
            style.configure(f"{kind}.Dot.TLabel", foreground=color, font=self.fonts["small"])

        flat = dict(lightcolor=c["card"], darkcolor=c["card"])
        style.configure(
            "TButton", padding=(px(12), px(5)), width=0, background=c["card"], foreground=c["label"],
            bordercolor=c["border_strong"], focusthickness=1, **flat,
        )
        style.map(
            "TButton",
            background=[("disabled", c["heading"]), ("pressed", c["pressed"]), ("active", c["hover"])],
            foreground=[("disabled", c["disabled"])],
            bordercolor=[("disabled", c["border"]), ("focus", c["accent"]), ("active", c["accent"])],
            lightcolor=[("pressed", c["pressed"]), ("active", c["hover"])],
            darkcolor=[("pressed", c["pressed"]), ("active", c["hover"])],
        )
        style.configure(
            "Accent.TButton", padding=(px(26), px(8)), font=self.fonts["button"], background=c["accent"],
            foreground="#FFFFFF", bordercolor=c["accent"], lightcolor=c["accent"], darkcolor=c["accent"],
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", c["accent_disabled"]), ("pressed", c["accent_press"]), ("active", c["accent_hover"])],
            foreground=[("disabled", "#FFFFFF")],
            bordercolor=[("disabled", c["accent_disabled"]), ("pressed", c["accent_press"]), ("active", c["accent_hover"])],
            lightcolor=[("disabled", c["accent_disabled"]), ("pressed", c["accent_press"]), ("active", c["accent_hover"])],
            darkcolor=[("disabled", c["accent_disabled"]), ("pressed", c["accent_press"]), ("active", c["accent_hover"])],
        )
        style.configure(
            "Link.TButton", padding=(px(6), px(2)), background=c["card"], foreground=c["accent"],
            bordercolor=c["card"], focusthickness=0, relief="flat", **flat,
        )
        style.map(
            "Link.TButton",
            background=[("pressed", c["accent_soft"]), ("active", c["accent_soft"])],
            foreground=[("disabled", c["disabled"]), ("active", c["accent_hover"])],
            bordercolor=[("active", c["accent_soft"])],
            lightcolor=[("active", c["accent_soft"])],
            darkcolor=[("active", c["accent_soft"])],
        )

        style.layout("Segment.TRadiobutton", style.layout("Toolbutton"))
        style.configure(
            "Segment.TRadiobutton", padding=(px(8), px(6)), anchor="center", background=c["heading"],
            foreground=c["muted"], bordercolor=c["border_strong"], lightcolor=c["heading"], darkcolor=c["heading"],
        )
        style.map(
            "Segment.TRadiobutton",
            background=[("selected", c["accent_soft"]), ("active", c["hover"])],
            foreground=[("disabled", c["disabled"]), ("selected", c["accent"]), ("active", c["label"])],
            bordercolor=[("selected", c["accent"])],
            lightcolor=[("selected", c["accent_soft"])],
            darkcolor=[("selected", c["accent_soft"])],
            font=[("selected", self.fonts["bold"])],
        )
        style.configure(
            "TRadiobutton", background=c["card"], foreground=c["label"], indicatorbackground=c["field"],
            indicatorforeground=c["accent"], upperbordercolor=c["border_strong"], lowerbordercolor=c["border_strong"],
            indicatormargin=(0, 0, px(6), 0),
        )
        style.map(
            "TRadiobutton",
            background=[("active", c["card"])],
            foreground=[("disabled", c["disabled"])],
            indicatorbackground=[("disabled", c["heading"]), ("pressed", c["accent_soft"])],
            upperbordercolor=[("selected", c["accent"]), ("active", c["accent"])],
            lowerbordercolor=[("selected", c["accent"]), ("active", c["accent"])],
        )

        field = dict(
            padding=(px(7), px(4)), fieldbackground=c["field"], bordercolor=c["border_strong"],
            lightcolor=c["field"], darkcolor=c["field"], foreground=c["text"],
        )
        field_map = dict(
            bordercolor=[("focus", c["accent"]), ("hover", c["disabled"])],
            lightcolor=[("focus", c["accent"])],
            fieldbackground=[("disabled", c["heading"])],
            foreground=[("disabled", c["disabled"])],
        )
        style.configure("TEntry", **field)
        style.map("TEntry", **field_map)
        style.configure("TSpinbox", arrowsize=px(12), arrowcolor=c["muted"], background=c["heading"], **field)
        style.map("TSpinbox", arrowcolor=[("disabled", c["disabled"])], **field_map)
        style.configure("TCombobox", arrowsize=px(13), arrowcolor=c["muted"], background=c["heading"], **field)
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", c["field"]), ("disabled", c["heading"])],
            selectbackground=[("readonly", c["field"])],
            selectforeground=[("readonly", c["text"])],
            bordercolor=field_map["bordercolor"],
            lightcolor=field_map["lightcolor"],
            foreground=field_map["foreground"],
            arrowcolor=[("disabled", c["disabled"])],
        )

        style.configure(
            "Treeview", background=c["card"], fieldbackground=c["card"], foreground=c["text"],
            rowheight=self.table_font.metrics("linespace") + px(9), bordercolor=c["border"], borderwidth=0,
            font=self.table_font,
        )
        style.map("Treeview", background=[("selected", c["accent_soft"])], foreground=[("selected", c["text"])])
        style.configure(
            "Treeview.Heading", font=self.heading_font, background=c["heading"], foreground=c["label"],
            bordercolor=c["border"], lightcolor=c["heading"], darkcolor=c["heading"], relief="flat",
            padding=(px(8), px(6)),
        )
        style.map("Treeview.Heading", background=[("active", c["hover"])], foreground=[("active", c["accent"])])
        # The group list shows a checkbox image in the tree column; no expand indicator is needed.
        style.layout("Groups.Treeview.Item", [("Treeitem.padding", {"sticky": "nswe", "children": [
            ("Treeitem.image", {"side": "left", "sticky": ""}),
            ("Treeitem.focus", {"side": "left", "sticky": "", "children": [
                ("Treeitem.text", {"side": "left", "sticky": ""})]})]})])
        style.configure("Groups.Treeview.Item", padding=(px(10), 0, 0, 0))
        style.configure(
            "TScrollbar", background=c["border_strong"], troughcolor=c["heading"], bordercolor=c["heading"],
            lightcolor=c["border_strong"], darkcolor=c["border_strong"], arrowcolor=c["muted"],
            arrowsize=px(12), gripcount=0,
        )
        style.map("TScrollbar", background=[("active", c["disabled"])])
        style.configure(
            "Accent.Horizontal.TProgressbar", background=c["accent"], troughcolor=c["border"],
            bordercolor=c["border"], lightcolor=c["accent"], darkcolor=c["accent"], thickness=px(8),
        )

    def _card(self, parent: tk.Widget, title: str):
        px = self.px
        card = ttk.Frame(parent, style="Card.TFrame", padding=(px(16), px(12), px(16), px(14)))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        head = ttk.Frame(card, style="CardBody.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, px(10)))
        head.columnconfigure(1, weight=1)
        ttk.Label(head, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        body = ttk.Frame(card, style="CardBody.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        return card, head, body

    def _build_ui(self) -> None:
        self._setup_style()
        px = self.px
        self.root.configure(bg=PALETTE["bg"])
        self.app_icon = draw_app_icon(self.root, 64)
        self.header_icon = draw_app_icon(self.root, px(30))
        self.check_images = {state: draw_checkbox(self.root, max(12, px(15)), state) for state in ("on", "off", "mixed")}
        try:
            self.root.iconphoto(True, self.app_icon)
        except tk.TclError:
            pass
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header = ttk.Frame(self.root, style="Bar.TFrame", padding=(px(20), px(10)))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(2, weight=1)
        ttk.Label(header, image=self.header_icon).grid(row=0, column=0, padx=(0, px(10)))
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(
            header, style="Subtitle.TLabel",
            text="按关键列拆分数据行，或按关键行拆分数据列；保留样式、公式、合并单元格和打印设置",
        ).grid(row=0, column=2, sticky="w", padx=(px(14), 0))
        self.help_button = ttk.Button(header, text="使用说明", style="Link.TButton", command=self.show_help)
        self.help_button.grid(row=0, column=3, sticky="e")
        tk.Frame(self.root, height=1, bg=PALETTE["border"]).grid(row=1, column=0, sticky="ew")

        body = ttk.Frame(self.root, padding=(px(16), px(14)))
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, minsize=px(340))
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        sidebar = ttk.Frame(body)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, px(14)))
        sidebar.columnconfigure(0, weight=1)
        workspace = ttk.Frame(body)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(0, weight=1, uniform="workspace")
        workspace.rowconfigure(1, weight=1, uniform="workspace")

        self._build_file_card(sidebar).grid(row=0, column=0, sticky="ew")
        self._build_options_card(sidebar).grid(row=1, column=0, sticky="ew", pady=(px(12), 0))
        self._build_tips_card(sidebar).grid(row=2, column=0, sticky="ew", pady=(px(12), 0))
        self._build_preview_card(workspace).grid(row=0, column=0, sticky="nsew")
        self._build_groups_card(workspace).grid(row=1, column=0, sticky="nsew", pady=(px(12), 0))
        self._build_status_bar()

    def _build_file_card(self, parent: tk.Widget) -> ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "① 选择文件")
        open_button = ttk.Button(head, text="打开输出目录", style="Link.TButton", command=self.open_output_dir)
        open_button.grid(row=0, column=2, sticky="e")
        Tooltip(open_button, TIPS["open"], self)
        body.columnconfigure(1, weight=1)
        ttk.Label(body, text="源文件", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, px(8)))
        self.input_entry = ttk.Entry(body, textvariable=self.input_var, width=12)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=(0, px(8)))
        self.input_entry.bind("<Return>", lambda _event: self.reload_workbook_info(force=True))
        self.input_entry.bind("<FocusOut>", self._on_input_focus_out)
        ttk.Button(body, text="选择…", command=self.choose_input).grid(row=0, column=2, sticky="ew")
        ttk.Label(body, text="工作表", style="Field.TLabel").grid(row=1, column=0, sticky="w", pady=(px(8), 0))
        self.sheet_combo = ttk.Combobox(body, textvariable=self.sheet_var, state="readonly", width=12)
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=(0, px(8)), pady=(px(8), 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)
        reload_button = ttk.Button(body, text="重新读取", command=lambda: self.reload_workbook_info(force=True))
        reload_button.grid(row=1, column=2, sticky="ew", pady=(px(8), 0))
        Tooltip(reload_button, TIPS["reload"], self)
        self.sheet_info = ttk.Label(body, text="尚未选择文件", style="Muted.TLabel")
        self.sheet_info.grid(row=2, column=1, columnspan=2, sticky="w", pady=(px(5), 0))
        tk.Frame(body, height=1, bg=PALETTE["border"]).grid(row=3, column=0, columnspan=3, sticky="ew", pady=px(12))
        ttk.Label(body, text="输出到", style="Field.TLabel").grid(row=4, column=0, sticky="w")
        self.output_entry = ttk.Entry(body, textvariable=self.output_var, width=12)
        self.output_entry.grid(row=4, column=1, sticky="ew", padx=(0, px(8)))
        ttk.Button(body, text="选择…", command=self.choose_output_dir).grid(row=4, column=2, sticky="ew")
        Tooltip(self.output_entry, TIPS["output"], self)
        return card

    def _parameter_frame(self, parent, fields, key_label, key_var, key_tip, on_key_selected):
        """Two fixed-size spinboxes on one line, then the key index with a picker by content."""
        px = self.px
        frame = ttk.Frame(parent, style="CardBody.TFrame")
        frame.columnconfigure(6, weight=1)
        column = 0
        for index, (label, variable, unit, tip) in enumerate(fields):
            text = ttk.Label(frame, text=label, style="Field.TLabel")
            text.grid(row=0, column=column, sticky="w", padx=(0 if index == 0 else px(16), px(6)))
            spin = ttk.Spinbox(frame, from_=0, to=1048576, textvariable=variable, width=5)
            spin.grid(row=0, column=column + 1, sticky="w")
            ttk.Label(frame, text=unit, style="Field.TLabel").grid(row=0, column=column + 2, sticky="w", padx=(px(5), 0))
            for widget in (text, spin):
                Tooltip(widget, tip, self)
            column += 3
        text = ttk.Label(frame, text=key_label, style="Field.TLabel")
        text.grid(row=1, column=0, sticky="w", pady=(px(10), 0), padx=(0, px(6)))
        spin = ttk.Spinbox(frame, from_=1, to=1048576, textvariable=key_var, width=5)
        spin.grid(row=1, column=1, sticky="w", pady=(px(10), 0))
        combo = ttk.Combobox(frame, state="readonly", width=8)
        combo.grid(row=1, column=2, columnspan=5, sticky="ew", padx=(px(8), 0), pady=(px(10), 0))
        combo.bind("<<ComboboxSelected>>", on_key_selected)
        for widget in (text, spin, combo):
            Tooltip(widget, key_tip, self)
        return frame, combo

    def _segmented(self, parent: tk.Widget, variable: tk.StringVar, options, command) -> ttk.Frame:
        frame = ttk.Frame(parent, style="CardBody.TFrame")
        frame.columnconfigure(tuple(range(len(options))), weight=1, uniform="segment")
        for column, (value, text) in enumerate(options):
            ttk.Radiobutton(
                frame, text=text, value=value, variable=variable, command=command, style="Segment.TRadiobutton",
            ).grid(row=0, column=column, sticky="ew")
        return frame

    def _build_options_card(self, parent: tk.Widget) -> ttk.Frame:
        px = self.px
        card, _head, body = self._card(parent, "② 拆分设置")
        self._segmented(
            body, self.mode_var, (("column", "按关键列拆分行"), ("row", "按关键行拆分列")), self.on_mode_changed,
        ).grid(row=0, column=0, sticky="ew")
        self.mode_hint = ttk.Label(body, style="Muted.TLabel", wraplength=px(300), justify="left")
        self.mode_hint.grid(row=1, column=0, sticky="ew", pady=(px(8), px(12)))
        self.column_options, self.key_column_combo = self._parameter_frame(
            body,
            [("表头", self.header_rows_var, "行", TIPS["header_rows"]), ("表尾", self.footer_rows_var, "行", TIPS["footer_rows"])],
            "关键列", self.key_column_var, TIPS["key_column"], self.on_key_column_selected,
        )
        self.column_options.grid(row=2, column=0, sticky="ew")
        self.row_options, self.key_row_combo = self._parameter_frame(
            body,
            [("左侧固定", self.header_cols_var, "列", TIPS["header_cols"]), ("右侧固定", self.footer_cols_var, "列", TIPS["footer_cols"])],
            "关键行", self.key_row_var, TIPS["key_row"], self.on_key_row_selected,
        )
        self.row_options.grid(row=2, column=0, sticky="ew")
        self.row_options.grid_remove()
        return card

    def _build_tips_card(self, parent: tk.Widget) -> ttk.Frame:
        card, _head, body = self._card(parent, "说明")
        ttk.Label(body, text=TIPS["notes"], style="Note.TLabel", justify="left").grid(row=0, column=0, sticky="w")
        return card

    def _scrolled_tree(self, parent: tk.Widget, horizontal: bool, **options) -> ttk.Treeview:
        frame = ttk.Frame(parent, style="Card.TFrame", padding=1)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = ttk.Treeview(frame, **options)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll_y.set)
        if horizontal:
            scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(xscrollcommand=scroll_x.set)
            tree.bind("<Shift-MouseWheel>", lambda event: tree.xview_scroll(-1 if event.delta > 0 else 1, "units"))
            tree.bind("<Shift-Button-4>", lambda _event: tree.xview_scroll(-1, "units"))
            tree.bind("<Shift-Button-5>", lambda _event: tree.xview_scroll(1, "units"))
        return tree

    def _build_preview_card(self, parent: tk.Widget) -> ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "表格预览")
        self.legend = ttk.Frame(head, style="CardBody.TFrame")
        self.legend.grid(row=0, column=2, sticky="e")
        body.rowconfigure(0, weight=1)
        self.preview_table = self._scrolled_tree(body, True, show="headings", selectmode="none", height=6)
        self.preview_table.bind("<ButtonRelease-1>", self.on_preview_table_click)
        self.preview_table.bind("<Button-3>", self.on_preview_context_menu)
        if sys.platform == "darwin":
            self.preview_table.bind("<Button-2>", self.on_preview_context_menu)
            self.preview_table.bind("<Control-Button-1>", self.on_preview_context_menu)
        self.preview_table.tag_configure("header", background=PALETTE["row_header"])
        self.preview_table.tag_configure("footer", background=PALETTE["row_footer"])
        self.preview_table.tag_configure("keyrow", background=PALETTE["row_key"])
        self.preview_table.tag_configure("split", background=PALETTE["row_split"])
        self.preview_table.tag_configure("gap", foreground=PALETTE["disabled"])
        self.preview_table.tag_configure("odd", background=PALETTE["row_stripe"])
        self.preview_message = ttk.Label(self.preview_table.master, style="Empty.TLabel", justify="center",
                                         anchor="center", wraplength=px(460))
        self._show_preview_message("尚未选择文件\n\n点击左侧“选择…”或按 Ctrl+O 打开 Excel 文件")
        self.preview_hint = ttk.Label(body, style="Muted.TLabel")
        self.preview_hint.grid(row=1, column=0, sticky="w", pady=(px(8), 0))
        return card

    def _build_groups_card(self, parent: tk.Widget) -> ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "拆分对象")
        self.group_summary = ttk.Label(head, style="Muted.TLabel")
        self.group_summary.grid(row=0, column=1, sticky="w", padx=(px(12), 0))
        tools = ttk.Frame(head, style="CardBody.TFrame")
        tools.grid(row=0, column=2, sticky="e")
        for text, command in (("全选", self.select_all_groups), ("全不选", self.select_no_groups), ("反选", self.invert_groups)):
            ttk.Button(tools, text=text, style="Link.TButton", command=command).pack(side="left", padx=(px(2), 0))
        body.rowconfigure(0, weight=1)
        tree = self._scrolled_tree(body, False, columns=("name", "count", "output"), show="tree headings",
                                   style="Groups.Treeview", height=3)
        self.split_preview_table = tree
        tree.heading("#0", image=self.check_images["on"], command=self.toggle_all_groups)
        tree.column("#0", width=px(44), minwidth=px(44), stretch=False)
        tree.heading("name", text="拆分对象", anchor="w")
        tree.heading("count", text="行数", anchor="center")
        tree.heading("output", text="输出文件", anchor="w")
        tree.column("name", width=px(220), minwidth=px(80), anchor="w", stretch=False)
        tree.column("count", width=px(70), minwidth=px(50), anchor="center", stretch=False)
        tree.column("output", width=px(320), minwidth=px(120), anchor="w", stretch=True)
        tree.tag_configure("excluded", foreground=PALETTE["disabled"])
        tree.tag_configure("blank", foreground=PALETTE["warning"])
        tree.bind("<ButtonRelease-1>", self.on_group_click)
        tree.bind("<space>", self.on_group_space)
        self.groups_message = ttk.Label(tree.master, style="Empty.TLabel", justify="center", anchor="center")
        self._show_groups_message("生成预览后，在这里勾选需要输出的拆分对象")

        options = ttk.Frame(body, style="CardBody.TFrame")
        options.grid(row=1, column=0, sticky="ew", pady=(px(10), 0))
        options.columnconfigure(3, weight=1)
        ttk.Label(options, text="输出方式", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, px(8)))
        self._segmented(
            options, self.output_mode_var, (("files", "每个对象一个文件"), ("workbook", "合并为一个工作簿")),
            self.update_group_names,
        ).grid(row=0, column=1, sticky="w")
        self.template_label = ttk.Label(options, text="文件命名", style="Field.TLabel")
        self.template_label.grid(row=0, column=2, sticky="w", padx=(px(20), px(8)))
        self.template_entry = ttk.Entry(options, textvariable=self.name_template_var, width=12)
        self.template_entry.grid(row=0, column=3, sticky="ew")
        Tooltip(self.template_entry, TIPS["template"], self)
        self.workbook_hint = ttk.Label(options, style="Field.TLabel")
        self.workbook_hint.grid(row=0, column=3, sticky="w")
        self.workbook_hint.grid_remove()
        self.template_hint = ttk.Label(options, style="Error.TLabel", justify="left")
        self.template_hint.grid(row=1, column=0, columnspan=4, sticky="e", pady=(px(4), 0))
        self.template_hint.grid_remove()
        options.bind("<Configure>", lambda event: self.template_hint.configure(wraplength=max(event.width, 1)))
        return card

    def _build_status_bar(self) -> None:
        px = self.px
        tk.Frame(self.root, height=1, bg=PALETTE["border"]).grid(row=3, column=0, sticky="ew")
        bar = ttk.Frame(self.root, style="Bar.TFrame", padding=(px(20), px(10)))
        bar.grid(row=4, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)
        self.status_dot = ttk.Label(bar, text="●", style="Idle.Dot.TLabel")
        self.status_dot.grid(row=0, column=0, padx=(0, px(8)))
        self.status_label = ttk.Label(bar, textvariable=self.status_var, style="Status.TLabel", width=1)
        self.status_label.grid(row=0, column=1, sticky="ew")
        self.status_label.bind("<Configure>", lambda event: self.status_label.configure(wraplength=max(event.width, 1)))
        self.progress_bar = ttk.Progressbar(bar, style="Accent.Horizontal.TProgressbar", length=px(220))
        self.progress_bar.grid(row=0, column=2, padx=(px(12), px(10)))
        self.progress_bar.grid_remove()
        self.cancel_button = ttk.Button(bar, text="取消", command=self.cancel_split)
        self.cancel_button.grid(row=0, column=3, padx=(0, px(10)))
        self.cancel_button.grid_remove()
        self.plan_label = ttk.Label(bar, style="Muted.TLabel")
        self.plan_label.grid(row=0, column=4, padx=(px(12), px(14)))
        self.start_button = ttk.Button(bar, text="开始拆分", style="Accent.TButton", command=self.run_split)
        self.start_button.grid(row=0, column=5)
        self._update_start_state()

    def _bind_shortcuts(self) -> None:
        modifier = "Command" if sys.platform == "darwin" else "Control"
        for key in ("o", "O"):
            self.root.bind(f"<{modifier}-{key}>", lambda _event: self._shortcut(self.choose_input))
        self.root.bind("<F5>", lambda _event: self._shortcut(lambda: self.reload_workbook_info(force=True)))
        self.root.bind(f"<{modifier}-Return>", lambda _event: self._shortcut(self.run_split))
        self.root.bind("<Escape>", lambda _event: self.cancel_split())

    def _shortcut(self, action) -> str:
        if not self._splitting:
            action()
        return "break"

    def _set_status(self, text: str, kind: str = "idle") -> None:
        self.status_var.set(text)
        self.status_dot.configure(style=f"{kind.capitalize()}.Dot.TLabel")

    def _show_preview_message(self, text: str, error: bool = False) -> None:
        self.preview_message.configure(text=text, style="EmptyError.TLabel" if error else "Empty.TLabel")
        self.preview_message.place(relx=0.5, rely=0.5, anchor="center")

    def _show_groups_message(self, text: str) -> None:
        self.groups_message.configure(text=text)
        self.groups_message.place(relx=0.5, rely=0.55, anchor="center")

    def _set_legend(self, items) -> None:
        px = self.px
        for child in self.legend.winfo_children():
            child.destroy()
        for marker, text in items:
            if marker in PALETTE:
                tk.Frame(self.legend, width=px(12), height=px(12), bg=PALETTE[marker], highlightthickness=1,
                         highlightbackground=PALETTE[marker + "_edge"]).pack(side="left", padx=(px(14), px(5)))
            else:
                ttk.Label(self.legend, text=marker, style="Field.TLabel").pack(side="left", padx=(px(14), px(3)))
            ttk.Label(self.legend, text=text, style="Muted.TLabel").pack(side="left")

    def show_help(self) -> None:
        messagebox.showinfo("使用说明", HELP_TEXT, parent=self.root)

    # ---- 文件与目录 ----

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择待拆分的 Excel 文件",
            initialdir=self.last_dir or None,
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm *.xls *.et"), ("所有文件", "*.*")],
        )
        if path:
            self.load_input_path(path)

    def load_input_path(self, path) -> None:
        path = str(path)
        self.input_var.set(path)
        self.input_entry.icursor("end")
        self.input_entry.xview_moveto(1.0)
        self.last_dir = str(Path(path).parent)
        if self._output_auto or not self.output_var.get().strip():
            self._set_output(str(Path(path).parent / "split_output"), auto=True)
        self.excluded.clear()
        self.reload_workbook_info()

    def _on_input_focus_out(self, _event) -> None:
        # Only an existing file is loaded automatically; partial paths wait for Enter or a button.
        text = self.input_var.get().strip()
        if text and text != self._loaded_path and not self._splitting and Path(text).is_file():
            self.load_input_path(text)

    def _set_output(self, value: str, auto: bool) -> None:
        self._setting_output = True
        try:
            self.output_var.set(value)
        finally:
            self._setting_output = False
        self._output_auto = auto
        self.output_entry.xview_moveto(1.0)

    def _on_output_edited(self, *_args) -> None:
        if not self._setting_output:
            self._output_auto = False

    def choose_output_dir(self) -> None:
        initial_dir = self.output_var.get().strip() or str(Path(self.input_var.get().strip()).parent)
        path = filedialog.askdirectory(title="选择输出目录", initialdir=initial_dir, mustexist=False)
        if path:
            self._set_output(path, auto=False)

    def _output_dir(self) -> Optional[Path]:
        output_text = self.output_var.get().strip()
        if output_text:
            return Path(output_text)
        input_text = self.input_var.get().strip()
        return Path(input_text).parent / "split_output" if input_text else None

    def open_output_dir(self, path: Optional[Path] = None) -> None:
        target = Path(path) if path else self._output_dir()
        if target is None:
            messagebox.showinfo("输出目录", "请先选择 Excel 文件或输出目录。", parent=self.root)
            return
        if not target.is_dir():
            messagebox.showinfo("输出目录", f"目录尚未创建：{target}\n开始拆分后会自动创建。", parent=self.root)
            return
        try:
            open_path(target)
        except OSError as exc:
            messagebox.showerror("无法打开目录", str(exc), parent=self.root)

    # ---- 后台任务 ----

    def _worker(self):
        while True:
            job = self._jobs.get()
            if job is None or self._closed:
                break
            kind, version, work, done = job
            if self._versions.get(kind) != version:
                continue
            try:
                value, error = work(), None
            except Exception as exc:
                value, error = None, exc
            self._results.put((kind, version, done, value, error))
        self._cache.clear()

    def _start_job(self, kind, work, done):
        version = self._versions.get(kind, 0) + 1
        self._versions[kind] = version
        self._jobs.put((kind, version, work, done))

    def _poll_jobs(self):
        try:
            while True:
                kind, version, done, value, error = self._results.get_nowait()
                if self._versions.get(kind) == version:
                    done(value, error)
        except queue.Empty:
            pass
        finally:
            latest = None
            try:
                while True:
                    latest = self._progress.get_nowait()
            except queue.Empty:
                pass
            if latest is not None and self._splitting:
                self._show_progress(*latest)
            if not self._closed:
                self._poll_after = self.root.after(75, self._poll_jobs)

    def _close(self):
        if self._splitting:
            if messagebox.askyesno(
                "正在拆分", "拆分尚未完成，是否停止并关闭窗口？\n未完成的拆分不会留下输出文件。",
                icon="warning", parent=self.root,
            ):
                self._close_requested = True
                self.cancel_split()
            return
        self._save_settings()
        self._closed = True
        self.root.after_cancel(self._poll_after)
        if self._preview_after is not None:
            self.root.after_cancel(self._preview_after)
        self._jobs.put(None)
        self.root.destroy()

    def _save_settings(self) -> None:
        save_settings(self.settings_path, {
            "mode": self.mode_var.get(),
            "header_rows": self.header_rows_var.get().strip(),
            "footer_rows": self.footer_rows_var.get().strip(),
            "key_column": self.key_column_var.get().strip(),
            "header_cols": self.header_cols_var.get().strip(),
            "footer_cols": self.footer_cols_var.get().strip(),
            "key_row": self.key_row_var.get().strip(),
            "output_mode": self.output_mode_var.get(),
            "name_template": self.name_template_var.get(),
            "last_dir": self.last_dir,
        })

    # ---- 读取与预览 ----

    def reload_workbook_info(self, force: bool = False) -> None:
        input_path = self.input_var.get().strip()
        if not input_path or self._splitting:
            return
        self._loaded_path = input_path
        self._preview_failed = False
        self._versions["preview"] = self._versions.get("preview", 0) + 1
        self._set_status("正在读取文件……", "busy")
        self.sheet_info.configure(text="正在读取……")
        self.clear_preview_table()
        self.clear_split_preview_table()
        self._show_preview_message("正在读取文件……")

        def done(info, error):
            if self.input_var.get().strip() != input_path:
                return
            if error:
                self._preview_failed = True
                self._update_start_state()
                self.sheet_var.set("")
                self.sheet_combo["values"] = ()
                self.sheet_info.configure(text="读取失败")
                self._set_status(f"读取失败：{error}", "error")
                self._show_preview_message(f"无法读取文件\n\n{error}", error=True)
                messagebox.showerror("读取失败", str(error), parent=self.root)
                return
            sheet_names = info["sheet_names"]
            self.sheet_combo["values"] = sheet_names
            if self.sheet_var.get().strip() not in sheet_names:
                self.sheet_var.set(sheet_names[0])
            self.refresh_sheet_preview()

        def work():
            if force:
                self._cache.clear()
            return load_workbook_info(Path(input_path), self._cache)

        self._start_job("load", work, done)

    def on_sheet_changed(self, _event: object) -> None:
        self.excluded.clear()
        self.refresh_sheet_preview()

    def on_option_changed(self, *_args: object) -> None:
        self._versions["preview"] = self._versions.get("preview", 0) + 1
        if self._preview_after is not None:
            self.root.after_cancel(self._preview_after)
        self._preview_after = self.root.after(300, self.refresh_sheet_preview)
        self._sync_key_combo()

    def _read_parameters(self, mode: str) -> Tuple[int, int, int]:
        if mode == "row":
            values = (self.header_cols_var.get(), self.footer_cols_var.get(), self.key_row_var.get())
        else:
            values = (self.header_rows_var.get(), self.footer_rows_var.get(), self.key_column_var.get())
        leading, trailing, key = (int(value.strip()) for value in values)
        return leading, trailing, key

    def _set_key_choices(self, mode: str, choices) -> None:
        self._key_choices[mode] = [number for number, _label in choices]
        combo = self.key_row_combo if mode == "row" else self.key_column_combo
        combo["values"] = [label for _number, label in choices]
        self._sync_key_combo()

    def _sync_key_combo(self) -> None:
        mode = self.mode_var.get()
        combo = self.key_row_combo if mode == "row" else self.key_column_combo
        variable = self.key_row_var if mode == "row" else self.key_column_var
        numbers = self._key_choices[mode]
        try:
            combo.current(numbers.index(int(variable.get().strip())))
        except (ValueError, tk.TclError):
            combo.set("")

    def _on_key_selected(self, mode: str, combo: ttk.Combobox, variable: tk.StringVar) -> None:
        index = combo.current()
        if 0 <= index < len(self._key_choices[mode]):
            number = str(self._key_choices[mode][index])
            if variable.get().strip() != number:
                variable.set(number)

    def on_key_column_selected(self, _event: object) -> None:
        self._on_key_selected("column", self.key_column_combo, self.key_column_var)

    def on_key_row_selected(self, _event: object) -> None:
        self._on_key_selected("row", self.key_row_combo, self.key_row_var)

    def on_mode_changed(self, refresh: bool = True) -> None:
        if self.mode_var.get() == "row":
            self.column_options.grid_remove()
            self.row_options.grid()
            self.mode_hint.configure(text="关键行相同的数据列放入同一个文件，左右固定列复制到每个文件。")
            self._set_legend([("row_key", "关键行"), ("（固定）", "固定列"), ("⋯", "省略")])
            self.preview_hint.configure(text=TIPS["row_mode"])
            self.split_preview_table.heading("count", text="列数")
        else:
            self.row_options.grid_remove()
            self.column_options.grid()
            self.mode_hint.configure(text="关键列相同的数据行放入同一个文件，表头表尾复制到每个文件。")
            self._set_legend([("row_header", "表头"), ("row_footer", "表尾"), ("★", "关键列"), ("⋯", "省略")])
            self.preview_hint.configure(text=TIPS["column_mode"])
            self.split_preview_table.heading("count", text="行数")
        if refresh:
            self.excluded.clear()
            self._sync_key_combo()
            self.refresh_sheet_preview()

    def refresh_sheet_preview(self) -> None:
        if self._preview_after is not None:
            self.root.after_cancel(self._preview_after)
            self._preview_after = None
        if self._splitting or self._closed:
            return
        input_path = self.input_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        if not input_path or not sheet_name:
            return
        self._versions["preview"] = self._versions.get("preview", 0) + 1
        mode = self.mode_var.get()
        try:
            leading, trailing, key = self._read_parameters(mode)
        except ValueError:
            self._preview_error(ValueError("拆分参数必须填写整数。"))
            return
        preview_function = build_sheet_preview_by_row if mode == "row" else build_sheet_preview
        self._set_status("正在生成预览……", "busy")

        def done(info, error):
            if self.input_var.get().strip() != input_path:
                return
            if error:
                self._preview_error(error)
                return
            self._preview_info = info
            self._preview_failed = False
            self.sheet_info.configure(text=f"{info['sheet_title']} · {info['max_row']} 行 × {info['max_column']} 列")
            self._set_key_choices(mode, info["key_choices"])
            self.render_preview_table(info["column_headers"], info["preview_rows"], info["preview_columns"], info["key_position"])
            self.render_split_preview_table(info["split_objects"])
            unit = "列" if mode == "row" else "行"
            blank = "（含空白关键字分组）" if info["blank_group"] else ""
            self._set_status(
                f"共 {info['data_count']} {unit}数据，分为 {info['group_count']} 组{blank}。预览只显示部分行列，拆分时处理整张工作表。"
            )

        self._start_job(
            "preview", lambda: preview_function(Path(input_path), sheet_name, leading, key, trailing, cache=self._cache), done
        )

    def _preview_error(self, error) -> None:
        self._preview_info = None
        self._preview_failed = True
        self._set_status(f"无法预览：{error}", "error")
        self.clear_preview_table()
        self.clear_split_preview_table()
        self._show_preview_message(f"无法预览\n\n{error}", error=True)

    def clear_preview_table(self) -> None:
        self.preview_table.delete(*self.preview_table.get_children())
        self.preview_table["columns"] = ()
        self.preview_column_map = {}
        self.preview_row_map = {}

    def render_preview_table(self, column_headers, preview_rows, preview_columns=None, key_position=None) -> None:
        table, px = self.preview_table, self.px
        measure, measure_heading = self.table_font.measure, self.heading_font.measure
        columns = ["row_no"] + [f"col_{index}" for index in range(len(column_headers))]
        table.delete(*table.get_children())
        table["columns"] = columns
        self.preview_column_map = {}
        self.preview_row_map = {}
        labels = [f"★ {row['row_no']}" if row["kind"] == "keyrow" else str(row["row_no"]) for row in preview_rows]
        table.heading("row_no", text="行", anchor="center")
        width = max([measure_heading("行")] + [measure(label) for label in labels]) + px(24)
        table.column("row_no", width=max(width, px(52)), minwidth=px(40), anchor="center", stretch=False)
        for index, header in enumerate(column_headers):
            column_id = f"col_{index}"
            widths = [measure_heading(header)]
            for row in preview_rows:
                if index < len(row["values"]):
                    # A long title in a header row is usually a merged cell; let it clip.
                    width = measure(str(row["values"][index]))
                    widths.append(min(width, px(120)) if row["kind"] == "header" else width)
            table.heading(column_id, text=header, anchor="w")
            width = min(max(max(widths) + px(22), px(64)), px(220))
            table.column(column_id, width=width, minwidth=px(40), anchor="w", stretch=False)
            if preview_columns is not None:
                self.preview_column_map[f"#{index + 2}"] = preview_columns[index]
        stripe = False
        for label, row in zip(labels, preview_rows):
            kind = row["kind"]
            if kind == "data":
                tags = ("odd",) if stripe else ()
                stripe = not stripe
            else:
                tags = (kind,)
            item = table.insert("", "end", values=[label] + [str(value) for value in row["values"]], tags=tags)
            if isinstance(row["row_no"], int):
                self.preview_row_map[item] = row["row_no"]
        self.preview_message.place_forget()
        table.xview_moveto(0)
        table.yview_moveto(0)
        if key_position is not None:
            self.root.after_idle(self._scroll_preview_to, f"col_{key_position}")

    def _scroll_preview_to(self, column_id: str) -> None:
        if self._closed:
            return
        table = self.preview_table
        columns = list(table["columns"])
        if column_id not in columns:
            return
        widths = [int(table.column(column, "width")) for column in columns]
        position = columns.index(column_id)
        right = sum(widths[:position + 1])
        overflow = right + self.px(24) - table.winfo_width()
        if overflow > 0:  # Scroll just far enough that the key column is fully visible.
            table.xview_moveto(overflow / sum(widths))

    def on_preview_table_click(self, event: tk.Event) -> None:
        if self._splitting:
            return
        region = self.preview_table.identify("region", event.x, event.y)
        if self.mode_var.get() == "row":
            row = self.preview_row_map.get(self.preview_table.identify_row(event.y)) if region == "cell" else None
            if row is not None and self.key_row_var.get().strip() != str(row):
                self.key_row_var.set(str(row))
            return
        if region != "heading":
            return
        column = self.preview_column_map.get(self.preview_table.identify_column(event.x))
        if column is not None and self.key_column_var.get().strip() != str(column):
            self.key_column_var.set(str(column))

    def on_preview_context_menu(self, event: tk.Event) -> None:
        info = self._preview_info
        table = self.preview_table
        region = table.identify("region", event.x, event.y)
        if self._splitting or not info or region not in ("heading", "cell"):
            return
        column = self.preview_column_map.get(table.identify_column(event.x))
        row = self.preview_row_map.get(table.identify_row(event.y)) if region == "cell" else None
        menu = tk.Menu(self.root, **self.menu_options)
        if self.mode_var.get() == "row":
            if row is not None:
                menu.add_command(label=f"设为关键行（第 {row} 行）", command=lambda: self.key_row_var.set(str(row)))
            if column is not None:
                letter, right = get_column_letter(column), info["max_column"] - column + 1
                menu.add_command(label=f"左侧固定到 {letter} 列（共 {column} 列）",
                                 command=lambda: self.header_cols_var.set(str(column)))
                menu.add_command(label=f"从 {letter} 列起固定到最右侧（共 {right} 列）",
                                 command=lambda: self.footer_cols_var.set(str(right)))
        else:
            if column is not None:
                menu.add_command(label=f"设为关键列（{get_column_letter(column)} 列）",
                                 command=lambda: self.key_column_var.set(str(column)))
            if row is not None:
                bottom = info["max_row"] - row + 1
                menu.add_command(label=f"表头到第 {row} 行为止（共 {row} 行）",
                                 command=lambda: self.header_rows_var.set(str(row)))
                menu.add_command(label=f"表尾从第 {row} 行开始（共 {bottom} 行）",
                                 command=lambda: self.footer_rows_var.set(str(bottom)))
        if menu.index("end") is not None:
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

    # ---- 拆分对象 ----

    def clear_split_preview_table(self) -> None:
        self.split_preview_table.delete(*self.split_preview_table.get_children())
        self.group_items = {}
        self.groups = None
        self._show_groups_message("生成预览后，在这里勾选需要输出的拆分对象")
        self.update_group_names()

    def render_split_preview_table(self, split_objects: List[Dict[str, object]]) -> None:
        tree = self.split_preview_table
        tree.delete(*tree.get_children())
        self.groups = list(split_objects)
        self.group_items = {}
        self.group_counts = {item["name"]: item["count"] for item in split_objects}
        self.excluded &= set(self.group_counts)
        blank = (self._preview_info or {}).get("blank_group")
        for item in split_objects:
            tags = ("blank",) if item["name"] == blank else ()
            iid = tree.insert("", "end", image=self.check_images["on"], values=(item["name"], item["count"], ""), tags=tags)
            self.group_items[iid] = item["name"]
        if split_objects:
            self.groups_message.place_forget()
        else:
            self._show_groups_message("当前条件下没有识别到拆分对象")
        tree.yview_moveto(0)
        self.update_group_names()

    def _selected_names(self) -> List[str]:
        return [item["name"] for item in self.groups or () if item["name"] not in self.excluded]

    def update_group_names(self) -> None:
        """Refresh checkboxes, planned output names, counts and the start button."""
        tree = self.split_preview_table
        selected = self._selected_names()
        single = self.output_mode_var.get() == "workbook"
        input_text = self.input_var.get().strip()
        stem = safe_file_name(Path(input_text).stem) if input_text else "源文件名"
        planned, error = {}, None
        if single:
            planned = dict(zip(selected, plan_sheet_titles(selected)))
        else:
            try:
                sheet = (self._preview_info or {}).get("sheet_title", self.sheet_var.get())
                reserved = {Path(input_text).name} if input_text else set()
                planned = dict(zip(selected, plan_file_names(self.name_template_var.get(), Path(input_text).stem or "源文件名",
                                                             sheet, selected, reserved)))
            except ValueError as exc:
                error = str(exc)
        self._naming_error = error
        for iid, name in self.group_items.items():
            included = name not in self.excluded
            tags = tuple(tag for tag in tree.item(iid, "tags") if tag != "excluded") + (() if included else ("excluded",))
            output = planned.get(name, "命名规则有误" if error else "") if included else "不输出"
            tree.item(iid, image=self.check_images["on" if included else "off"], tags=tags,
                      values=(name, self.group_counts.get(name, ""), output))
        groups = self.groups or []
        state = "on" if not self.excluded else ("off" if not selected else "mixed")
        tree.heading("#0", image=self.check_images[state])
        tree.heading("output", text="输出工作表" if single else "输出文件")
        unit = "列" if self.mode_var.get() == "row" else "行"
        total = sum(self.group_counts.get(name, 0) for name in selected)
        self.group_summary.configure(
            text=f"共 {len(groups)} 组 · 已选 {len(selected)} 组 · {total} {unit}" if groups else ""
        )
        if single:
            self.template_label.configure(text="输出文件")
            self.template_entry.grid_remove()
            self.workbook_hint.configure(text=f"{stem}_拆分.xlsx（每个对象一个工作表）")
            self.workbook_hint.grid()
            self.template_hint.grid_remove()
        else:
            self.template_label.configure(text="文件命名")
            self.workbook_hint.grid_remove()
            self.template_entry.grid()
            self.template_hint.configure(text=error or "")
            if error:
                self.template_hint.grid()
            else:
                self.template_hint.grid_remove()
        if not groups:
            plan = ""
        elif not selected:
            plan = "未选择拆分对象"
        else:
            plan = f"将生成 1 个工作簿（{len(selected)} 个工作表）" if single else f"将生成 {len(selected)} 个文件"
        self.plan_label.configure(text=plan)
        self._update_start_state()

    def _update_start_state(self) -> None:
        ready = (bool(self.input_var.get().strip()) and not self._preview_failed and not self._naming_error
                 and not (self.groups and not self._selected_names()))
        if not self._splitting:
            self.start_button.configure(state="normal" if ready else "disabled")

    def on_group_click(self, event: tk.Event) -> None:
        tree = self.split_preview_table
        if self._splitting or tree.identify("region", event.x, event.y) not in ("tree", "cell"):
            return
        item = tree.identify_row(event.y)
        if item in self.group_items:
            self._toggle_groups([item])

    def on_group_space(self, _event: tk.Event) -> str:
        if not self._splitting:
            self._toggle_groups(self.split_preview_table.selection())
        return "break"

    def _toggle_groups(self, items) -> None:
        names = [self.group_items[item] for item in items if item in self.group_items]
        if not names:
            return
        include = any(name in self.excluded for name in names)
        for name in names:
            if include:
                self.excluded.discard(name)
            else:
                self.excluded.add(name)
        self.update_group_names()

    def select_all_groups(self) -> None:
        if not self._splitting:
            self.excluded.clear()
            self.update_group_names()

    def select_no_groups(self) -> None:
        if not self._splitting:
            self.excluded = {item["name"] for item in self.groups or ()}
            self.update_group_names()

    def invert_groups(self) -> None:
        if not self._splitting:
            self.excluded = {item["name"] for item in self.groups or ()} - self.excluded
            self.update_group_names()

    def toggle_all_groups(self) -> None:
        if self.excluded:
            self.select_all_groups()
        else:
            self.select_no_groups()

    # ---- 拆分 ----

    def _set_split_busy(self, busy):
        self._splitting = busy
        if busy:
            self._disabled_widgets = []

            def visit(widget):
                for child in widget.winfo_children():
                    if isinstance(child, (ttk.Entry, ttk.Button, ttk.Radiobutton)) and child not in (self.cancel_button, self.help_button):
                        self._disabled_widgets.append((child, child.cget("state")))
                        child.configure(state="disabled")
                    visit(child)

            visit(self.root)
            self.plan_label.grid_remove()
            self.progress_bar.configure(mode="indeterminate", value=0)
            self.progress_bar.grid()
            self.progress_bar.start(15)
            self.cancel_button.configure(state="normal")
            self.cancel_button.grid()
        else:
            for widget, state in self._disabled_widgets:
                widget.configure(state=state)
            self._disabled_widgets = []
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            self.cancel_button.grid_remove()
            self.plan_label.grid()
            self._update_start_state()

    def _show_progress(self, done: int, total: int, name: Optional[str]) -> None:
        if str(self.progress_bar.cget("mode")) != "determinate":
            self.progress_bar.stop()
            self.progress_bar.configure(mode="determinate")
        self.progress_bar.configure(maximum=max(total, 1), value=done)
        if self._cancel.is_set():
            return
        if name is None:
            self._set_status("正在写入输出文件……", "busy")
        else:
            self._set_status(f"正在生成 {done + 1}/{total}：{name}", "busy")

    def cancel_split(self) -> None:
        if not self._splitting or self._cancel.is_set():
            return
        self._cancel.set()
        self.cancel_button.configure(state="disabled")
        self._set_status("正在取消，当前步骤完成后停止……", "warning")

    def run_split(self) -> None:
        if self._splitting:
            return
        input_text = self.input_var.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件", "请先选择待拆分的 Excel 文件。", parent=self.root)
            return
        mode = self.mode_var.get()
        try:
            leading, trailing, key = self._read_parameters(mode)
        except ValueError:
            messagebox.showwarning("参数错误", "固定行列数和关键行列序号必须是整数。", parent=self.root)
            return
        single = self.output_mode_var.get() == "workbook"
        template = self.name_template_var.get()
        if not single:
            try:
                validate_name_template(template)
            except ValueError as exc:
                messagebox.showwarning("文件命名规则有误", str(exc), parent=self.root)
                return
        selected = self._selected_names() if self.groups is not None and self.excluded else None
        if selected == []:
            messagebox.showwarning("未选择拆分对象", "请至少勾选一个拆分对象。", parent=self.root)
            return
        split_function = split_workbook_by_row if mode == "row" else split_workbook
        input_path = Path(input_text)
        sheet = self.sheet_var.get().strip() or None
        output_text = self.output_var.get().strip()
        output_dir = Path(output_text) if output_text else None
        self._versions["preview"] = self._versions.get("preview", 0) + 1
        self._versions["load"] = self._versions.get("load", 0) + 1
        self._cancel.clear()
        while not self._progress.empty():
            self._progress.get_nowait()
        self._set_split_busy(True)
        self._set_status("正在准备拆分……", "busy")

        def progress(done_count, total, name):
            self._progress.put((done_count, total, name))

        def work():
            return split_function(
                input_path, sheet, leading, key, output_dir, trailing, groups=selected, name_template=template,
                single_workbook=single, progress=progress, cancel=self._cancel, cache=self._cache,
            )

        def done(files, error):
            self._set_split_busy(False)
            if self._close_requested:
                self._close()
                return
            if isinstance(error, SplitCancelled):
                self._set_status("已取消拆分，未生成任何文件。", "warning")
                return
            if error:
                self._set_status(f"拆分失败：{error}", "error")
                messagebox.showerror("拆分失败", str(error), parent=self.root)
                return
            self._save_settings()
            output_parent = files[0].parent
            if single:
                summary = f"已生成工作簿 {files[0].name}（{len(files.sheet_titles)} 个工作表）"
            else:
                summary = f"已生成 {len(files)} 个文件"
                names = selected if selected is not None else [item["name"] for item in self.groups or ()]
                if len(names) == len(files):
                    for item, name in self.group_items.items():
                        if name in names:
                            self.split_preview_table.set(item, "output", files[names.index(name)].name)
            message = f"{summary}。\n输出目录：{output_parent}\n公式将在 Excel/WPS 打开时重算。"
            if files.warnings:
                self._set_status(f"{summary}，其中 {len(files.warnings)} 处含引用错误，请核对。", "warning")
                message += "\n\n需要核对的公式：\n" + "\n".join(files.warnings[:10])
                title, icon = "拆分完成，需检查公式", "warning"
            else:
                self._set_status(f"{summary}，输出目录：{output_parent}", "success")
                title, icon = "拆分完成", "info"
            if messagebox.askyesno(title, message + "\n\n是否打开输出目录？", icon=icon, parent=self.root):
                self.open_output_dir(output_parent)

        self._start_job("split", work, done)

    def run(self) -> None:
        self.root.mainloop()


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
    SplitterApp().run()


if __name__ == "__main__":
    main()

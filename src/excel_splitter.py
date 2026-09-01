from __future__ import annotations

import argparse
import copy
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
import tkinter as tk
from openpyxl.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from tkinter import filedialog, messagebox, ttk


CELL_OR_RANGE_PATTERN = re.compile(
    r"(?P<prefix>(?:'[^']+'|[A-Za-z0-9_]+)!)?"
    r"(?P<start_col>\$?[A-Z]{1,3})(?P<start_row>\$?\d+)"
    r"(?::(?P<end_col>\$?[A-Z]{1,3})(?P<end_row>\$?\d+))?"
)
OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
CONVERTIBLE_EXTENSIONS = {".xls", ".et"}


def parse_args() -> argparse.Namespace:
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
    return parser.parse_args()


def normalize_group_value(value: object) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped or None
    return str(value)


def safe_file_name(name: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", name.strip())
    return cleaned[:80] or "未命名"


def load_workbook_compatible(input_path: Path) -> Tuple[openpyxl.Workbook, Optional[tempfile.TemporaryDirectory]]:
    suffix = input_path.suffix.lower()
    if suffix in OPENPYXL_EXTENSIONS:
        return openpyxl.load_workbook(input_path, data_only=False, read_only=False), None

    if suffix == ".et":
        try:
            file_handle = input_path.open("rb")
            workbook = openpyxl.load_workbook(file_handle, data_only=False, read_only=False)
            workbook._external_file_handle = file_handle
            return workbook, None
        except Exception:
            pass

    if suffix in CONVERTIBLE_EXTENSIONS:
        converted_path, temp_dir = convert_to_xlsx(input_path)
        return openpyxl.load_workbook(converted_path, data_only=False, read_only=False), temp_dir

    raise ValueError(f"暂不支持该文件格式：{suffix}")


def close_workbook_compatible(workbook: openpyxl.Workbook, temp_dir: Optional[tempfile.TemporaryDirectory]) -> None:
    external_file_handle = getattr(workbook, "_external_file_handle", None)
    workbook.close()
    if external_file_handle:
        external_file_handle.close()
    if temp_dir:
        temp_dir.cleanup()


def convert_to_xlsx(input_path: Path) -> Tuple[Path, tempfile.TemporaryDirectory]:
    temp_dir = tempfile.TemporaryDirectory()
    temp_path = Path(temp_dir.name)

    try:
        converted_path = convert_with_libreoffice(input_path, temp_path)
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
    $workbook = $app.Workbooks.Open($InputPath)
    $workbook.SaveAs($OutputPath, 51)
} finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    $app.Quit()
}
""",
        encoding="utf-8",
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
    )
    if result.returncode != 0 or not output_path.exists():
        return None
    return output_path


def copy_workbook_metadata(source_wb: openpyxl.Workbook, target_wb: openpyxl.Workbook) -> None:
    target_wb.properties = copy.copy(source_wb.properties)
    target_wb.security = copy.copy(source_wb.security)
    target_wb.calculation = copy.copy(source_wb.calculation)
    target_wb.views = copy.copy(source_wb.views)
    target_wb.code_name = source_wb.code_name
    target_wb.defined_names = copy.copy(source_wb.defined_names)
    if source_wb.vba_archive is not None:
        target_wb.vba_archive = copy.copy(source_wb.vba_archive)
    target_wb.template = source_wb.template
    target_wb.iso_dates = source_wb.iso_dates


def copy_sheet_settings(source_ws: Worksheet, target_ws: Worksheet) -> None:
    target_ws.sheet_format = copy.copy(source_ws.sheet_format)
    target_ws.sheet_properties = copy.copy(source_ws.sheet_properties)
    target_ws.page_margins = copy.copy(source_ws.page_margins)
    target_ws.page_setup = copy.copy(source_ws.page_setup)
    target_ws.print_options = copy.copy(source_ws.print_options)
    target_ws.views = copy.copy(source_ws.views)
    target_ws.freeze_panes = source_ws.freeze_panes
    target_ws.auto_filter = copy.copy(source_ws.auto_filter)
    target_ws.print_title_cols = source_ws.print_title_cols
    target_ws.print_title_rows = source_ws.print_title_rows
    target_ws.oddHeader = copy.copy(source_ws.oddHeader)
    target_ws.oddFooter = copy.copy(source_ws.oddFooter)
    target_ws.evenHeader = copy.copy(source_ws.evenHeader)
    target_ws.evenFooter = copy.copy(source_ws.evenFooter)
    target_ws.firstHeader = copy.copy(source_ws.firstHeader)
    target_ws.firstFooter = copy.copy(source_ws.firstFooter)


def copy_row_and_column_dimensions(
    source_ws: Worksheet,
    target_ws: Worksheet,
    included_rows: Sequence[int],
) -> None:
    for column_key, dimension in source_ws.column_dimensions.items():
        target_dimension = target_ws.column_dimensions[column_key]
        target_dimension.width = dimension.width
        target_dimension.hidden = dimension.hidden
        target_dimension.bestFit = dimension.bestFit
        target_dimension.outlineLevel = dimension.outlineLevel
        target_dimension.outline_level = dimension.outline_level
        target_dimension.collapsed = dimension.collapsed
        target_dimension.min = dimension.min
        target_dimension.max = dimension.max

    for target_row, source_row in enumerate(included_rows, start=1):
        if source_row in source_ws.row_dimensions:
            source_dimension = source_ws.row_dimensions[source_row]
            target_dimension = target_ws.row_dimensions[target_row]
            target_dimension.height = source_dimension.height
            target_dimension.hidden = source_dimension.hidden
            target_dimension.outlineLevel = source_dimension.outlineLevel
            target_dimension.outline_level = source_dimension.outline_level
            target_dimension.collapsed = source_dimension.collapsed
            target_dimension.ht = source_dimension.ht


def rebuild_formula(
    formula: str,
    row_mapping: Dict[int, int],
    header_rows: int,
    footer_start: Optional[int] = None,
) -> str:
    def map_single_row(row_number: int) -> int:
        return row_mapping.get(row_number, row_number)

    def replace(match: re.Match[str]) -> str:
        prefix = match.group("prefix") or ""
        start_col = match.group("start_col")
        start_row_token = match.group("start_row")
        end_col = match.group("end_col")
        end_row_token = match.group("end_row")

        start_row = int(start_row_token.replace("$", ""))
        start_row_is_abs = start_row_token.startswith("$")

        if end_col and end_row_token:
            end_row = int(end_row_token.replace("$", ""))
            end_row_is_abs = end_row_token.startswith("$")

            data_rows = [
                mapped_row
                for original_row, mapped_row in row_mapping.items()
                if original_row > header_rows
                and (footer_start is None or original_row < footer_start)
                and start_row <= original_row <= end_row
            ]
            if data_rows:
                new_start_row = min(data_rows)
                new_end_row = max(data_rows)
            else:
                new_start_row = map_single_row(start_row)
                new_end_row = map_single_row(end_row)

            if start_row <= header_rows and end_row <= header_rows:
                new_start_row = map_single_row(start_row)
                new_end_row = map_single_row(end_row)

            return (
                f"{prefix}{start_col}{'$' if start_row_is_abs else ''}{new_start_row}:"
                f"{end_col}{'$' if end_row_is_abs else ''}{new_end_row}"
            )

        new_row = map_single_row(start_row)
        return f"{prefix}{start_col}{'$' if start_row_is_abs else ''}{new_row}"

    return CELL_OR_RANGE_PATTERN.sub(replace, formula)


def copy_cell_value_and_style(
    source_cell: Cell,
    target_cell: Cell,
    row_mapping: Dict[int, int],
    header_rows: int,
    footer_start: Optional[int] = None,
) -> None:
    if source_cell.data_type == "f":
        formula = str(source_cell.value)
        if not formula.startswith("="):
            formula = f"={formula}"
        target_cell.value = rebuild_formula(
            formula=formula,
            row_mapping=row_mapping,
            header_rows=header_rows,
            footer_start=footer_start,
        )
    else:
        target_cell.value = source_cell.value

    if source_cell.has_style:
        target_cell.font = copy.copy(source_cell.font)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.alignment = copy.copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy.copy(source_cell.protection)

    if source_cell.hyperlink:
        target_cell._hyperlink = copy.copy(source_cell.hyperlink)
    if source_cell.comment:
        target_cell.comment = copy.copy(source_cell.comment)


def rebuild_merged_ranges(
    source_ws: Worksheet,
    target_ws: Worksheet,
    row_mapping: Dict[int, int],
    included_rows_set: set[int],
) -> None:
    for merged_range in source_ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        source_rows = set(range(min_row, max_row + 1))
        if not source_rows.issubset(included_rows_set):
            continue

        mapped_rows = [row_mapping[row] for row in range(min_row, max_row + 1)]
        target_ws.merge_cells(
            f"{get_column_letter(min_col)}{min(mapped_rows)}:{get_column_letter(max_col)}{max(mapped_rows)}"
        )


def build_merged_value_map(ws: Worksheet, key_column: int) -> Dict[int, object]:
    merged_value_map: Dict[int, object] = {}
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        if not (min_col <= key_column <= max_col):
            continue
        top_left_value = ws.cell(min_row, min_col).value
        for row in range(min_row, max_row + 1):
            merged_value_map[row] = top_left_value
    return merged_value_map


def collect_group_rows(
    ws: Worksheet,
    header_rows: int,
    key_column: int,
    footer_rows: int = 0,
) -> Dict[str, List[int]]:
    groups: Dict[str, List[int]] = {}
    merged_value_map = build_merged_value_map(ws, key_column)
    last_data_row = ws.max_row - max(0, footer_rows)
    for row in range(header_rows + 1, last_data_row + 1):
        raw_value = ws.cell(row=row, column=key_column).value
        if raw_value is None and row in merged_value_map:
            raw_value = merged_value_map[row]
        group_value = normalize_group_value(raw_value)
        if group_value:
            groups.setdefault(group_value, []).append(row)
    return groups


def build_target_sheet(
    source_ws: Worksheet,
    group_rows: Sequence[int],
    header_rows: int,
    source_wb: openpyxl.Workbook,
    footer_rows: int = 0,
) -> openpyxl.Workbook:
    target_wb = openpyxl.Workbook()
    target_ws = target_wb.active
    target_ws.title = source_ws.title[:31]

    copy_workbook_metadata(source_wb, target_wb)
    copy_sheet_settings(source_ws, target_ws)

    footer_count = max(0, footer_rows)
    footer_row_list = (
        list(range(source_ws.max_row - footer_count + 1, source_ws.max_row + 1))
        if footer_count
        else []
    )
    footer_start = footer_row_list[0] if footer_row_list else None

    included_rows = list(range(1, header_rows + 1)) + list(group_rows) + footer_row_list
    row_mapping = {source_row: target_row for target_row, source_row in enumerate(included_rows, start=1)}
    included_rows_set = set(included_rows)

    copy_row_and_column_dimensions(source_ws, target_ws, included_rows)

    for source_row in included_rows:
        target_row = row_mapping[source_row]
        for column in range(1, source_ws.max_column + 1):
            source_cell = source_ws.cell(row=source_row, column=column)
            target_cell = target_ws.cell(row=target_row, column=column)
            copy_cell_value_and_style(
                source_cell,
                target_cell,
                row_mapping=row_mapping,
                header_rows=header_rows,
                footer_start=footer_start,
            )

    rebuild_merged_ranges(source_ws, target_ws, row_mapping, included_rows_set)
    target_ws.sheet_state = source_ws.sheet_state
    target_ws.row_breaks = copy.copy(source_ws.row_breaks)
    target_ws.col_breaks = copy.copy(source_ws.col_breaks)

    return target_wb


# ---------------------------------------------------------------------------
# 按关键行拆分（把数据列按关键行中的关键字分组，每组列生成一个文件）
# ---------------------------------------------------------------------------


def rebuild_formula_by_columns(
    formula: str,
    col_mapping: Dict[int, int],
    header_cols: int,
    footer_start: Optional[int] = None,
) -> str:
    def map_single_col(col_number: int) -> int:
        return col_mapping.get(col_number, col_number)

    def replace(match: re.Match[str]) -> str:
        prefix = match.group("prefix") or ""
        start_col_token = match.group("start_col")
        start_row_token = match.group("start_row")
        end_col_token = match.group("end_col")
        end_row_token = match.group("end_row")

        start_col = column_index_from_string(start_col_token.replace("$", ""))
        start_col_is_abs = start_col_token.startswith("$")

        if end_col_token and end_row_token:
            end_col = column_index_from_string(end_col_token.replace("$", ""))
            end_col_is_abs = end_col_token.startswith("$")

            data_cols = [
                mapped_col
                for original_col, mapped_col in col_mapping.items()
                if original_col > header_cols
                and (footer_start is None or original_col < footer_start)
                and start_col <= original_col <= end_col
            ]
            if data_cols:
                new_start_col = min(data_cols)
                new_end_col = max(data_cols)
            else:
                new_start_col = map_single_col(start_col)
                new_end_col = map_single_col(end_col)

            if start_col <= header_cols and end_col <= header_cols:
                new_start_col = map_single_col(start_col)
                new_end_col = map_single_col(end_col)

            return (
                f"{prefix}{'$' if start_col_is_abs else ''}{get_column_letter(new_start_col)}{start_row_token}:"
                f"{'$' if end_col_is_abs else ''}{get_column_letter(new_end_col)}{end_row_token}"
            )

        new_col = map_single_col(start_col)
        return f"{prefix}{'$' if start_col_is_abs else ''}{get_column_letter(new_col)}{start_row_token}"

    return CELL_OR_RANGE_PATTERN.sub(replace, formula)


def copy_cell_value_and_style_by_columns(
    source_cell: Cell,
    target_cell: Cell,
    col_mapping: Dict[int, int],
    header_cols: int,
    footer_start: Optional[int] = None,
) -> None:
    if source_cell.data_type == "f":
        formula = str(source_cell.value)
        if not formula.startswith("="):
            formula = f"={formula}"
        target_cell.value = rebuild_formula_by_columns(
            formula=formula,
            col_mapping=col_mapping,
            header_cols=header_cols,
            footer_start=footer_start,
        )
    else:
        target_cell.value = source_cell.value

    if source_cell.has_style:
        target_cell.font = copy.copy(source_cell.font)
        target_cell.fill = copy.copy(source_cell.fill)
        target_cell.border = copy.copy(source_cell.border)
        target_cell.alignment = copy.copy(source_cell.alignment)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy.copy(source_cell.protection)

    if source_cell.hyperlink:
        target_cell._hyperlink = copy.copy(source_cell.hyperlink)
    if source_cell.comment:
        target_cell.comment = copy.copy(source_cell.comment)


def copy_dimensions_by_columns(
    source_ws: Worksheet,
    target_ws: Worksheet,
    included_cols: Sequence[int],
) -> None:
    for row_key, source_dimension in source_ws.row_dimensions.items():
        target_dimension = target_ws.row_dimensions[row_key]
        target_dimension.height = source_dimension.height
        target_dimension.hidden = source_dimension.hidden
        target_dimension.outlineLevel = source_dimension.outlineLevel
        target_dimension.outline_level = source_dimension.outline_level
        target_dimension.collapsed = source_dimension.collapsed
        target_dimension.ht = source_dimension.ht

    source_col_dims: Dict[int, object] = {}
    for column_key, dimension in source_ws.column_dimensions.items():
        try:
            fallback_index = column_index_from_string(column_key)
        except ValueError:
            fallback_index = None
        start = dimension.min or fallback_index
        if start is None:
            continue
        end = dimension.max or start
        for column in range(start, end + 1):
            source_col_dims[column] = dimension

    for target_col, source_col in enumerate(included_cols, start=1):
        dimension = source_col_dims.get(source_col)
        if dimension is None:
            continue
        target_letter = get_column_letter(target_col)
        target_dimension = target_ws.column_dimensions[target_letter]
        target_dimension.width = dimension.width
        target_dimension.hidden = dimension.hidden
        target_dimension.bestFit = dimension.bestFit
        target_dimension.outlineLevel = dimension.outlineLevel
        target_dimension.outline_level = dimension.outline_level
        target_dimension.collapsed = dimension.collapsed
        target_dimension.min = target_col
        target_dimension.max = target_col


def rebuild_merged_ranges_by_columns(
    source_ws: Worksheet,
    target_ws: Worksheet,
    col_mapping: Dict[int, int],
    included_cols_set: set[int],
) -> None:
    for merged_range in source_ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        # 合并区域可能跨越被剔除的数据列（如覆盖整行的标题），
        # 此时收缩到保留下来的列继续合并。
        mapped_cols = [
            col_mapping[col]
            for col in range(min_col, max_col + 1)
            if col in included_cols_set
        ]
        if not mapped_cols:
            continue
        new_min_col, new_max_col = min(mapped_cols), max(mapped_cols)
        if min_col not in included_cols_set:
            # 原合并区域左上角所在列被剔除时，把其内容补到新的左上角，
            # 避免合并后的标题丢失文字。
            source_top_left = source_ws.cell(row=min_row, column=min_col)
            target_top_left = target_ws.cell(row=min_row, column=new_min_col)
            if target_top_left.value is None and source_top_left.value is not None:
                target_top_left.value = source_top_left.value
                if source_top_left.has_style:
                    target_top_left.font = copy.copy(source_top_left.font)
                    target_top_left.fill = copy.copy(source_top_left.fill)
                    target_top_left.border = copy.copy(source_top_left.border)
                    target_top_left.alignment = copy.copy(source_top_left.alignment)
                    target_top_left.number_format = source_top_left.number_format
        if new_min_col == new_max_col and min_row == max_row:
            continue
        target_ws.merge_cells(
            f"{get_column_letter(new_min_col)}{min_row}:{get_column_letter(new_max_col)}{max_row}"
        )


def build_merged_value_map_by_row(ws: Worksheet, key_row: int) -> Dict[int, object]:
    merged_value_map: Dict[int, object] = {}
    for merged_range in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        if not (min_row <= key_row <= max_row):
            continue
        top_left_value = ws.cell(min_row, min_col).value
        for column in range(min_col, max_col + 1):
            merged_value_map[column] = top_left_value
    return merged_value_map


def collect_group_columns(
    ws: Worksheet,
    header_cols: int,
    key_row: int,
    footer_cols: int = 0,
) -> Dict[str, List[int]]:
    groups: Dict[str, List[int]] = {}
    merged_value_map = build_merged_value_map_by_row(ws, key_row)
    last_data_col = ws.max_column - max(0, footer_cols)
    for column in range(header_cols + 1, last_data_col + 1):
        raw_value = ws.cell(row=key_row, column=column).value
        if raw_value is None and column in merged_value_map:
            raw_value = merged_value_map[column]
        group_value = normalize_group_value(raw_value)
        if group_value:
            groups.setdefault(group_value, []).append(column)
    return groups


def build_target_sheet_by_columns(
    source_ws: Worksheet,
    group_cols: Sequence[int],
    header_cols: int,
    source_wb: openpyxl.Workbook,
    footer_cols: int = 0,
) -> openpyxl.Workbook:
    target_wb = openpyxl.Workbook()
    target_ws = target_wb.active
    target_ws.title = source_ws.title[:31]

    copy_workbook_metadata(source_wb, target_wb)
    copy_sheet_settings(source_ws, target_ws)

    footer_count = max(0, footer_cols)
    footer_col_list = (
        list(range(source_ws.max_column - footer_count + 1, source_ws.max_column + 1))
        if footer_count
        else []
    )
    footer_start = footer_col_list[0] if footer_col_list else None

    included_cols = list(range(1, header_cols + 1)) + list(group_cols) + footer_col_list
    col_mapping = {source_col: target_col for target_col, source_col in enumerate(included_cols, start=1)}
    included_cols_set = set(included_cols)

    copy_dimensions_by_columns(source_ws, target_ws, included_cols)

    for source_col in included_cols:
        target_col = col_mapping[source_col]
        for row in range(1, source_ws.max_row + 1):
            source_cell = source_ws.cell(row=row, column=source_col)
            target_cell = target_ws.cell(row=row, column=target_col)
            copy_cell_value_and_style_by_columns(
                source_cell,
                target_cell,
                col_mapping=col_mapping,
                header_cols=header_cols,
                footer_start=footer_start,
            )

    rebuild_merged_ranges_by_columns(source_ws, target_ws, col_mapping, included_cols_set)
    target_ws.sheet_state = source_ws.sheet_state
    target_ws.row_breaks = copy.copy(source_ws.row_breaks)
    target_ws.col_breaks = copy.copy(source_ws.col_breaks)

    return target_wb


def split_workbook_by_row(
    input_path: Path,
    sheet_name: Optional[str],
    header_cols: int,
    key_row: int,
    output_dir: Optional[Path],
    footer_cols: int = 0,
) -> List[Path]:
    source_wb, temp_dir = load_workbook_compatible(input_path)
    source_ws = source_wb[sheet_name] if sheet_name else source_wb.worksheets[0]

    try:
        footer_cols = max(0, footer_cols)
        header_cols = max(0, header_cols)
        if header_cols + footer_cols >= source_ws.max_column:
            raise ValueError("左侧固定列数与右侧固定列数之和不能大于或等于工作表总列数。")
        if key_row > source_ws.max_row or key_row < 1:
            raise ValueError("关键行超出工作表行数范围。")

        groups = collect_group_columns(source_ws, header_cols, key_row, footer_cols)
        if not groups:
            raise ValueError("未在关键行中找到可拆分的数据。")

        destination = output_dir or input_path.parent / "split_output"
        destination.mkdir(parents=True, exist_ok=True)

        output_files: List[Path] = []
        base_name = input_path.stem
        for group_name, group_cols in groups.items():
            target_wb = build_target_sheet_by_columns(
                source_ws=source_ws,
                group_cols=group_cols,
                header_cols=header_cols,
                source_wb=source_wb,
                footer_cols=footer_cols,
            )
            output_path = destination / f"{base_name}_{safe_file_name(group_name)}.xlsx"
            target_wb.save(output_path)
            target_wb.close()
            output_files.append(output_path)

        return output_files
    finally:
        close_workbook_compatible(source_wb, temp_dir)


def split_workbook(
    input_path: Path,
    sheet_name: Optional[str],
    header_rows: int,
    key_column: int,
    output_dir: Optional[Path],
    footer_rows: int = 0,
) -> List[Path]:
    source_wb, temp_dir = load_workbook_compatible(input_path)
    source_ws = source_wb[sheet_name] if sheet_name else source_wb.worksheets[0]

    try:
        footer_rows = max(0, footer_rows)
        if header_rows + footer_rows >= source_ws.max_row:
            raise ValueError("表头行数与表尾行数之和不能大于或等于工作表总行数。")
        if key_column > source_ws.max_column:
            raise ValueError("关键列超出工作表最大列数。")

        groups = collect_group_rows(source_ws, header_rows, key_column, footer_rows)
        if not groups:
            raise ValueError("未在关键列中找到可拆分的数据。")

        destination = output_dir or input_path.parent / "split_output"
        destination.mkdir(parents=True, exist_ok=True)

        output_files: List[Path] = []
        base_name = input_path.stem
        for group_name, group_rows in groups.items():
            target_wb = build_target_sheet(
                source_ws=source_ws,
                group_rows=group_rows,
                header_rows=header_rows,
                source_wb=source_wb,
                footer_rows=footer_rows,
            )
            output_path = destination / f"{base_name}_{safe_file_name(group_name)}.xlsx"
            target_wb.save(output_path)
            target_wb.close()
            output_files.append(output_path)

        return output_files
    finally:
        close_workbook_compatible(source_wb, temp_dir)


def load_workbook_info(input_path: Path) -> Dict[str, object]:
    workbook, temp_dir = load_workbook_compatible(input_path)
    try:
        sheet_names = workbook.sheetnames
        first_sheet = workbook[sheet_names[0]]
        column_preview = []
        preview_row = min(first_sheet.max_row, 6)
        for column in range(1, first_sheet.max_column + 1):
            value = first_sheet.cell(preview_row, column).value
            display = normalize_group_value(value) or "(空)"
            column_preview.append(f"{column}. {get_column_letter(column)} - {display}")
        return {
            "sheet_names": sheet_names,
            "max_row": first_sheet.max_row,
            "max_column": first_sheet.max_column,
            "column_preview": column_preview,
        }
    finally:
        close_workbook_compatible(workbook, temp_dir)


def build_sheet_preview(
    input_path: Path,
    sheet_name: str,
    header_rows: int,
    key_column: int,
    footer_rows: int = 0,
) -> Dict[str, object]:
    workbook, temp_dir = load_workbook_compatible(input_path)
    ws = workbook[sheet_name]
    try:
        safe_header_rows = max(1, min(header_rows, ws.max_row))
        safe_key_column = max(1, min(key_column, ws.max_column))
        safe_footer_rows = max(0, min(footer_rows, ws.max_row - safe_header_rows - 1))
        footer_start = ws.max_row - safe_footer_rows + 1 if safe_footer_rows else None

        preview_columns = list(range(1, ws.max_column + 1))

        column_headers: List[str] = []
        header_reference_row = min(ws.max_row, safe_header_rows)
        for column in preview_columns:
            header_text = normalize_group_value(ws.cell(row=header_reference_row, column=column).value) or "(空)"
            prefix = "[关键] " if column == safe_key_column else ""
            column_headers.append(f"{prefix}{get_column_letter(column)}列/{column}: {header_text}")

        def row_values(target_row: int) -> List[str]:
            collected = []
            for column in preview_columns:
                value = normalize_group_value(ws.cell(row=target_row, column=column).value) or ""
                collected.append(value.replace("\n", " "))
            return collected

        preview_rows: List[Dict[str, object]] = []
        preview_data_limit = min(ws.max_row, safe_header_rows + 8)
        if footer_start is not None:
            preview_data_limit = min(preview_data_limit, footer_start - 1)
        for row in range(1, preview_data_limit + 1):
            row_kind = "header" if row <= safe_header_rows else "data"
            preview_rows.append({"row_no": row, "kind": row_kind, "values": row_values(row)})
            if row == safe_header_rows and row < ws.max_row:
                preview_rows.append(
                    {
                        "row_no": "↓",
                        "kind": "split",
                        "values": ["以下数据行将参与拆分"] + [""] * (len(preview_columns) - 1),
                    }
                )
        if footer_start is not None:
            preview_rows.append(
                {
                    "row_no": "↓",
                    "kind": "split",
                    "values": ["以下为固定表尾，会附加到每个文件末尾"] + [""] * (len(preview_columns) - 1),
                }
            )
            for row in range(footer_start, ws.max_row + 1):
                preview_rows.append({"row_no": row, "kind": "footer", "values": row_values(row)})

        key_column_preview: List[str] = []
        for column in range(1, ws.max_column + 1):
            value = normalize_group_value(ws.cell(header_reference_row, column).value) or "(空)"
            marker = " <关键列>" if column == safe_key_column else ""
            key_column_preview.append(f"{column}. {get_column_letter(column)} - {value}{marker}")

        groups = collect_group_rows(ws, safe_header_rows, safe_key_column, safe_footer_rows)
        split_objects = [{"name": name, "count": len(rows)} for name, rows in groups.items()]

        return {
            "sheet_title": ws.title,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "preview_columns": preview_columns,
            "column_headers": column_headers,
            "preview_rows": preview_rows,
            "key_column_preview": key_column_preview,
            "split_objects": split_objects,
            "group_count": len(split_objects),
        }
    finally:
        close_workbook_compatible(workbook, temp_dir)


def build_sheet_preview_by_row(
    input_path: Path,
    sheet_name: str,
    header_cols: int,
    key_row: int,
    footer_cols: int = 0,
) -> Dict[str, object]:
    workbook, temp_dir = load_workbook_compatible(input_path)
    ws = workbook[sheet_name]
    try:
        safe_header_cols = max(0, min(header_cols, ws.max_column - 1))
        safe_key_row = max(1, min(key_row, ws.max_row))
        safe_footer_cols = max(0, min(footer_cols, ws.max_column - safe_header_cols - 1))
        footer_start = ws.max_column - safe_footer_cols + 1 if safe_footer_cols else None

        preview_columns = list(range(1, ws.max_column + 1))

        column_headers: List[str] = []
        for column in preview_columns:
            key_text = normalize_group_value(ws.cell(row=safe_key_row, column=column).value) or "(空)"
            if column <= safe_header_cols:
                prefix = "[左固定] "
            elif footer_start is not None and column >= footer_start:
                prefix = "[右固定] "
            else:
                prefix = ""
            column_headers.append(f"{prefix}{get_column_letter(column)}列/{column}: {key_text}")

        def row_values(target_row: int) -> List[str]:
            collected = []
            for column in preview_columns:
                value = normalize_group_value(ws.cell(row=target_row, column=column).value) or ""
                collected.append(value.replace("\n", " "))
            return collected

        preview_rows: List[Dict[str, object]] = []
        preview_limit = min(ws.max_row, max(safe_key_row, 1) + 8)
        for row in range(1, preview_limit + 1):
            row_kind = "keyrow" if row == safe_key_row else "data"
            row_label = f"{row} <关键行>" if row == safe_key_row else row
            preview_rows.append({"row_no": row_label, "kind": row_kind, "values": row_values(row)})
        if preview_limit < ws.max_row:
            preview_rows.append(
                {
                    "row_no": "…",
                    "kind": "split",
                    "values": [f"（其余 {ws.max_row - preview_limit} 行未显示，所有行都会完整保留）"]
                    + [""] * (len(preview_columns) - 1),
                }
            )

        key_row_preview: List[str] = []
        preview_row_limit = min(ws.max_row, 30)
        for row in range(1, preview_row_limit + 1):
            samples: List[str] = []
            for column in range(1, ws.max_column + 1):
                value = normalize_group_value(ws.cell(row=row, column=column).value)
                if value:
                    samples.append(value.replace("\n", " "))
                if len(samples) >= 3:
                    break
            display = " | ".join(samples) if samples else "(空行)"
            marker = " <关键行>" if row == safe_key_row else ""
            key_row_preview.append(f"{row}. {display}{marker}")

        groups = collect_group_columns(ws, safe_header_cols, safe_key_row, safe_footer_cols)
        split_objects = [{"name": name, "count": len(cols)} for name, cols in groups.items()]

        return {
            "sheet_title": ws.title,
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "preview_columns": preview_columns,
            "column_headers": column_headers,
            "preview_rows": preview_rows,
            "key_row_preview": key_row_preview,
            "split_objects": split_objects,
            "group_count": len(split_objects),
        }
    finally:
        close_workbook_compatible(workbook, temp_dir)


class SplitterApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Excel 拆表工具")
        self.root.geometry("980x760")
        self.root.minsize(900, 660)
        self.root.configure(bg="#F4F6FA")

        self.input_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.mode_var = tk.StringVar(value="column")
        self.header_rows_var = tk.StringVar(value="6")
        self.footer_rows_var = tk.StringVar(value="0")
        self.key_column_var = tk.StringVar(value="3")
        self.header_cols_var = tk.StringVar(value="1")
        self.footer_cols_var = tk.StringVar(value="0")
        self.key_row_var = tk.StringVar(value="1")
        self.output_var = tk.StringVar()
        self.summary_var = tk.StringVar(value="先选择 Excel 文件，界面会自动加载工作表和列信息。")
        self.preview_column_map: Dict[str, int] = {}

        self._build_ui()
        self.header_rows_var.trace_add("write", self.on_option_changed)
        self.footer_rows_var.trace_add("write", self.on_option_changed)
        self.key_column_var.trace_add("write", self.on_option_changed)
        self.header_cols_var.trace_add("write", self.on_option_changed)
        self.footer_cols_var.trace_add("write", self.on_option_changed)
        self.key_row_var.trace_add("write", self.on_option_changed)

    def _build_ui(self) -> None:
        # ---- 配色与字体 ----
        BG = "#F4F6FA"            # 窗口底色
        ACCENT = "#2F6FED"        # 主色（蓝）
        ACCENT_HOVER = "#255FCC"
        ACCENT_PRESS = "#1E4FB0"
        TEXT = "#1F2937"          # 正文
        MUTED = "#6B7280"         # 辅助灰
        BORDER = "#D9DEE7"        # 边框
        FONT = "Microsoft YaHei UI"

        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", font=(FONT, 10), background=BG, foreground=TEXT, bordercolor=BORDER)
        style.configure("TFrame", background=BG)
        style.configure("TLabel", background=BG, foreground=TEXT)
        style.configure("Title.TLabel", font=(FONT, 17, "bold"), background=BG, foreground=TEXT)
        style.configure("Muted.TLabel", background=BG, foreground=MUTED)
        style.configure("Summary.TLabel", background=BG, foreground=ACCENT, font=(FONT, 10, "bold"))
        style.configure(
            "TLabelframe", background=BG, bordercolor=BORDER, relief="solid", borderwidth=1
        )
        style.configure(
            "TLabelframe.Label", background=BG, foreground=TEXT, font=(FONT, 10, "bold")
        )
        style.configure("TRadiobutton", background=BG, foreground=TEXT)
        style.map("TRadiobutton", background=[("active", BG)])
        style.configure(
            "TButton",
            padding=(14, 6),
            background="#FFFFFF",
            foreground=TEXT,
            bordercolor=BORDER,
            focuscolor=ACCENT,
        )
        style.map(
            "TButton",
            background=[("pressed", "#E5EAF3"), ("active", "#EFF3FB")],
            bordercolor=[("active", ACCENT)],
        )
        style.configure(
            "Accent.TButton",
            padding=(22, 9),
            font=(FONT, 11, "bold"),
            background=ACCENT,
            foreground="#FFFFFF",
            bordercolor=ACCENT,
        )
        style.map(
            "Accent.TButton",
            background=[("pressed", ACCENT_PRESS), ("active", ACCENT_HOVER)],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "TEntry", padding=(6, 4), fieldbackground="#FFFFFF",
            bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER,
        )
        style.configure(
            "TCombobox", padding=(6, 4), fieldbackground="#FFFFFF", bordercolor=BORDER
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF")],
            bordercolor=[("active", ACCENT)],
        )
        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=TEXT,
            rowheight=24,
            bordercolor=BORDER,
            font=(FONT, 9),
        )
        style.configure(
            "Treeview.Heading",
            font=(FONT, 9, "bold"),
            background="#EEF2F8",
            foreground=TEXT,
            padding=(6, 4),
        )
        style.map(
            "Treeview",
            background=[("selected", "#D6E4FF")],
            foreground=[("selected", TEXT)],
        )
        style.configure("TScrollbar", background="#E5EAF3", troughcolor=BG, bordercolor=BG, arrowcolor=MUTED)
        style.configure("TSeparator", background=BORDER)

        container = ttk.Frame(self.root, padding=18)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(6, weight=1)

        title = ttk.Label(container, text="Excel 拆表工具", style="Title.TLabel")
        title.grid(row=0, column=0, columnspan=3, sticky="w")

        desc = ttk.Label(
            container,
            text="支持按关键列（拆数据行）或按关键行（拆数据列）两种方式拆分。会保留原表样式、格式和公式。",
            style="Muted.TLabel",
        )
        desc.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 6))

        ttk.Separator(container, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="sew", pady=(0, 10)
        )

        ttk.Label(container, text="Excel 文件").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=self.input_var).grid(row=2, column=1, sticky="ew", pady=6, padx=(0, 8))
        ttk.Button(container, text="浏览", command=self.choose_input).grid(row=2, column=2, sticky="ew", pady=6)

        ttk.Label(container, text="工作表").grid(row=3, column=0, sticky="w", pady=6)
        self.sheet_combo = ttk.Combobox(container, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=3, column=1, sticky="ew", pady=6, padx=(0, 8))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)
        ttk.Button(container, text="刷新", command=self.reload_workbook_info).grid(row=3, column=2, sticky="ew", pady=6)

        options = ttk.Frame(container)
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(8, 12))
        options.columnconfigure(0, weight=1)

        mode_bar = ttk.Frame(options)
        mode_bar.grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Label(mode_bar, text="拆分方式").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(
            mode_bar,
            text="按关键列拆分（每组数据行一个文件）",
            variable=self.mode_var,
            value="column",
            command=self.on_mode_changed,
        ).grid(row=0, column=1, sticky="w", padx=(12, 0))
        ttk.Radiobutton(
            mode_bar,
            text="按关键行拆分（每组数据列一个文件）",
            variable=self.mode_var,
            value="row",
            command=self.on_mode_changed,
        ).grid(row=0, column=2, sticky="w", padx=(16, 0))

        self.column_options = ttk.Frame(options)
        self.column_options.grid(row=1, column=0, sticky="ew")
        self.column_options.columnconfigure(1, weight=1)
        self.column_options.columnconfigure(3, weight=1)

        ttk.Label(self.column_options, text="表头行数").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(self.column_options, textvariable=self.header_rows_var, width=10).grid(
            row=0, column=1, sticky="w", padx=(8, 24), pady=(0, 6)
        )
        ttk.Label(self.column_options, text="表尾行数").grid(row=0, column=2, sticky="w", pady=(0, 6))
        ttk.Entry(self.column_options, textvariable=self.footer_rows_var, width=10).grid(
            row=0, column=3, sticky="w", padx=(8, 0), pady=(0, 6)
        )

        ttk.Label(self.column_options, text="关键列序号").grid(row=1, column=0, sticky="w")
        ttk.Entry(self.column_options, textvariable=self.key_column_var, width=10).grid(
            row=1, column=1, sticky="w", padx=(8, 24)
        )
        ttk.Label(self.column_options, text="或点选列").grid(row=1, column=2, sticky="w")
        self.key_column_combo = ttk.Combobox(self.column_options, state="readonly")
        self.key_column_combo.grid(row=1, column=3, sticky="ew", padx=(8, 0))
        self.key_column_combo.bind("<<ComboboxSelected>>", self.on_key_column_selected)

        self.row_options = ttk.Frame(options)
        self.row_options.grid(row=1, column=0, sticky="ew")
        self.row_options.columnconfigure(1, weight=1)
        self.row_options.columnconfigure(3, weight=1)

        ttk.Label(self.row_options, text="左侧固定列数").grid(row=0, column=0, sticky="w", pady=(0, 6))
        ttk.Entry(self.row_options, textvariable=self.header_cols_var, width=10).grid(
            row=0, column=1, sticky="w", padx=(8, 24), pady=(0, 6)
        )
        ttk.Label(self.row_options, text="右侧固定列数").grid(row=0, column=2, sticky="w", pady=(0, 6))
        ttk.Entry(self.row_options, textvariable=self.footer_cols_var, width=10).grid(
            row=0, column=3, sticky="w", padx=(8, 0), pady=(0, 6)
        )

        ttk.Label(self.row_options, text="关键行序号").grid(row=1, column=0, sticky="w")
        ttk.Entry(self.row_options, textvariable=self.key_row_var, width=10).grid(
            row=1, column=1, sticky="w", padx=(8, 24)
        )
        ttk.Label(self.row_options, text="或点选行").grid(row=1, column=2, sticky="w")
        self.key_row_combo = ttk.Combobox(self.row_options, state="readonly")
        self.key_row_combo.grid(row=1, column=3, sticky="ew", padx=(8, 0))
        self.key_row_combo.bind("<<ComboboxSelected>>", self.on_key_row_selected)

        self.row_options.grid_remove()

        summary = ttk.Label(container, textvariable=self.summary_var, style="Summary.TLabel")
        summary.grid(row=5, column=0, columnspan=3, sticky="new", pady=(0, 12))

        preview_area = ttk.Frame(container)
        preview_area.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(0, 12))
        preview_area.columnconfigure(0, weight=5)
        preview_area.columnconfigure(1, weight=2)
        preview_area.rowconfigure(0, weight=1)

        preview_box = ttk.LabelFrame(preview_area, text="表格预览", height=520)
        preview_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        preview_box.columnconfigure(0, weight=1)
        preview_box.rowconfigure(0, weight=1)
        preview_box.rowconfigure(1, weight=1)
        preview_box.grid_propagate(False)
        self.preview_hint_var = tk.StringVar(value="浅黄色为表头，浅绿色为固定表尾，蓝色为分界提示；关键列会在列标题中标记，可用底部横向滚动条左右查看。")
        ttk.Label(preview_box, textvariable=self.preview_hint_var, style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )
        preview_table_frame = ttk.Frame(preview_box)
        preview_table_frame.grid(row=1, column=0, sticky="nsew")
        preview_table_frame.columnconfigure(0, weight=1)
        preview_table_frame.rowconfigure(0, weight=1)
        self.preview_table = ttk.Treeview(preview_table_frame, show="headings")
        self.preview_table.grid(row=0, column=0, sticky="nsew")
        self.preview_table.bind("<ButtonRelease-1>", self.on_preview_table_click)
        preview_scroll_y = ttk.Scrollbar(preview_table_frame, orient="vertical", command=self.preview_table.yview)
        preview_scroll_y.grid(row=0, column=1, sticky="ns")
        preview_scroll_x = ttk.Scrollbar(preview_table_frame, orient="horizontal", command=self.preview_table.xview)
        preview_scroll_x.grid(row=1, column=0, sticky="ew")
        self.preview_table.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        self.preview_table.tag_configure("header", background="#fff4cc")
        self.preview_table.tag_configure("split", background="#dceeff")
        self.preview_table.tag_configure("footer", background="#e2efda")
        self.preview_table.tag_configure("keyrow", background="#ffe0e0")

        side_panel = ttk.Frame(preview_area)
        side_panel.grid(row=0, column=1, sticky="nsew")
        side_panel.columnconfigure(0, weight=1)
        side_panel.rowconfigure(0, weight=1)
        side_panel.rowconfigure(1, weight=1)

        self.reference_box = ttk.LabelFrame(side_panel, text="关键列参考", height=220)
        column_box = self.reference_box
        column_box.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        column_box.columnconfigure(0, weight=1)
        column_box.rowconfigure(0, weight=1)
        column_box.grid_propagate(False)
        self.preview_text = tk.Text(
            column_box,
            height=8,
            wrap="word",
            font=("Microsoft YaHei UI", 9),
            bg="#FFFFFF",
            fg="#1F2937",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#D9DEE7",
            highlightcolor="#2F6FED",
            padx=8,
            pady=6,
        )
        self.preview_text.grid(row=0, column=0, sticky="nsew")
        self.preview_text.configure(state="disabled")

        self.split_box = ttk.LabelFrame(side_panel, text="拆分对象预览", height=292)
        split_box = self.split_box
        split_box.grid(row=1, column=0, sticky="nsew")
        split_box.columnconfigure(0, weight=1)
        split_box.rowconfigure(0, weight=1)
        split_box.grid_propagate(False)
        self.split_preview_table = ttk.Treeview(split_box, columns=("name", "count"), show="headings", height=10)
        self.split_preview_table.heading("name", text="拆分对象")
        self.split_preview_table.heading("count", text="行数")
        self.split_preview_table.column("name", width=180, anchor="w")
        self.split_preview_table.column("count", width=70, anchor="center")
        self.split_preview_table.grid(row=0, column=0, sticky="nsew")
        split_scroll_y = ttk.Scrollbar(split_box, orient="vertical", command=self.split_preview_table.yview)
        split_scroll_y.grid(row=0, column=1, sticky="ns")
        self.split_preview_table.configure(yscrollcommand=split_scroll_y.set)

        ttk.Label(container, text="输出目录").grid(row=7, column=0, sticky="nw", pady=6)
        output_frame = ttk.Frame(container)
        output_frame.grid(row=7, column=1, columnspan=2, sticky="ew", pady=6)
        output_frame.columnconfigure(0, weight=1)
        ttk.Entry(output_frame, textvariable=self.output_var).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(output_frame, text="选择目录", command=self.choose_output_dir).grid(row=0, column=1, sticky="ew")

        action_bar = ttk.Frame(container)
        action_bar.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(16, 0))
        action_bar.columnconfigure(0, weight=1)
        ttk.Label(
            action_bar,
            text="输出为空时，默认生成到源文件同级目录的 split_output。",
            style="Muted.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            action_bar, text="开始拆分", command=self.run_split, style="Accent.TButton"
        ).grid(row=0, column=1, sticky="e")

    def choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择待拆分的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm *.xls *.et")],
        )
        if not path:
            return
        self.input_var.set(path)
        if not self.output_var.get().strip():
            self.output_var.set(str(Path(path).parent / "split_output"))
        self.reload_workbook_info()

    def choose_output_dir(self) -> None:
        initial_dir = self.output_var.get().strip() or self.input_var.get().strip() or "."
        path = filedialog.askdirectory(title="选择输出目录", initialdir=initial_dir, mustexist=False)
        if path:
            self.output_var.set(path)

    def reload_workbook_info(self) -> None:
        input_path = self.input_var.get().strip()
        if not input_path:
            return

        try:
            info = load_workbook_info(Path(input_path))
        except Exception as exc:
            messagebox.showerror("读取失败", f"无法读取 Excel 文件：\n{exc}")
            return

        sheet_names = info["sheet_names"]
        self.sheet_combo["values"] = sheet_names
        current_sheet = self.sheet_var.get().strip()
        if current_sheet not in sheet_names:
            self.sheet_var.set(sheet_names[0])

        self.summary_var.set(
            f"共 {len(sheet_names)} 个工作表，默认工作表 {self.sheet_var.get()}，"
            f"首个工作表共有 {info['max_row']} 行、{info['max_column']} 列。"
        )
        self.set_text_content(self.preview_text, info["column_preview"])
        self.key_column_combo["values"] = info["column_preview"]
        self.sync_key_column_combo()
        self.clear_preview_table()
        self.clear_split_preview_table()
        self.refresh_sheet_preview()

    def on_sheet_changed(self, _event: object) -> None:
        self.refresh_sheet_preview()

    def on_option_changed(self, *_args: object) -> None:
        self.root.after_idle(self.refresh_sheet_preview)

    def on_key_column_selected(self, _event: object) -> None:
        selected = self.key_column_combo.get().strip()
        if not selected:
            return
        try:
            column_no = int(selected.split(".", 1)[0])
        except ValueError:
            return
        if self.key_column_var.get().strip() != str(column_no):
            self.key_column_var.set(str(column_no))

    def on_key_row_selected(self, _event: object) -> None:
        selected = self.key_row_combo.get().strip()
        if not selected:
            return
        try:
            row_no = int(selected.split(".", 1)[0])
        except ValueError:
            return
        if self.key_row_var.get().strip() != str(row_no):
            self.key_row_var.set(str(row_no))

    def on_mode_changed(self) -> None:
        mode = self.mode_var.get()
        if mode == "row":
            self.column_options.grid_remove()
            self.row_options.grid()
            self.reference_box.configure(text="关键行参考")
            self.split_box.configure(text="拆分对象预览")
            self.split_preview_table.heading("count", text="列数")
            self.preview_hint_var.set(
                "浅红色为关键行，列标题中会标记左右固定列；点击表格中的任意一行即可将其设为关键行。"
            )
        else:
            self.row_options.grid_remove()
            self.column_options.grid()
            self.reference_box.configure(text="关键列参考")
            self.split_box.configure(text="拆分对象预览")
            self.split_preview_table.heading("count", text="行数")
            self.preview_hint_var.set(
                "浅黄色为表头，浅绿色为固定表尾，蓝色为分界提示；关键列会在列标题中标记，可用底部横向滚动条左右查看。"
            )
        self.refresh_sheet_preview()

    def on_preview_table_click(self, event: tk.Event) -> None:
        region = self.preview_table.identify("region", event.x, event.y)
        if self.mode_var.get() == "row":
            if region != "cell":
                return
            item_id = self.preview_table.identify_row(event.y)
            if not item_id:
                return
            values = self.preview_table.item(item_id, "values")
            if not values:
                return
            row_text = str(values[0]).split(" ", 1)[0].strip()
            if not row_text.isdigit():
                return
            if self.key_row_var.get().strip() != row_text:
                self.key_row_var.set(row_text)
            return

        if region != "heading":
            return
        column_id = self.preview_table.identify_column(event.x)
        if not column_id or column_id == "#1":
            return
        mapped = self.preview_column_map.get(column_id)
        if mapped is not None and self.key_column_var.get().strip() != str(mapped):
            self.key_column_var.set(str(mapped))

    def refresh_sheet_preview(self) -> None:
        input_path = self.input_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        if not input_path or not sheet_name:
            return

        try:
            if self.mode_var.get() == "row":
                header_cols = self.parse_int(self.header_cols_var.get(), fallback=1)
                footer_cols = self.parse_int(self.footer_cols_var.get(), fallback=0)
                key_row = self.parse_int(self.key_row_var.get(), fallback=1)
                info = build_sheet_preview_by_row(
                    input_path=Path(input_path),
                    sheet_name=sheet_name,
                    header_cols=header_cols,
                    key_row=key_row,
                    footer_cols=footer_cols,
                )
                self.summary_var.set(
                    f"当前工作表 {info['sheet_title']}，共有 {info['max_row']} 行、{info['max_column']} 列。"
                    f" 已按左固定 {header_cols} 列、右固定 {footer_cols} 列、关键行 {key_row} 生成预览，预计拆分为 {info['group_count']} 个对象。"
                )
                self.set_text_content(self.preview_text, info["key_row_preview"])
                self.key_row_combo["values"] = info["key_row_preview"]
                self.sync_key_row_combo()
                self.render_preview_table(info["column_headers"], info["preview_rows"])
                self.render_split_preview_table(info["split_objects"])
            else:
                header_rows = self.parse_int(self.header_rows_var.get(), fallback=6)
                footer_rows = self.parse_int(self.footer_rows_var.get(), fallback=0)
                key_column = self.parse_int(self.key_column_var.get(), fallback=1)
                info = build_sheet_preview(
                    input_path=Path(input_path),
                    sheet_name=sheet_name,
                    header_rows=header_rows,
                    key_column=key_column,
                    footer_rows=footer_rows,
                )
                self.summary_var.set(
                    f"当前工作表 {info['sheet_title']}，共有 {info['max_row']} 行、{info['max_column']} 列。"
                    f" 已按表头 {header_rows} 行、表尾 {footer_rows} 行、关键列 {key_column} 生成预览，预计拆分为 {info['group_count']} 个对象。"
                )
                self.set_text_content(self.preview_text, info["key_column_preview"])
                self.key_column_combo["values"] = info["key_column_preview"]
                self.sync_key_column_combo()
                self.render_preview_table(info["column_headers"], info["preview_rows"])
                self.render_split_preview_table(info["split_objects"])
        except Exception:
            self.summary_var.set("参数变更后暂时无法生成预览，请检查文件和拆分参数。")
            self.set_text_content(self.preview_text, ["预览失败"])
            self.clear_preview_table()
            self.clear_split_preview_table()

    def set_text_content(self, widget: tk.Text, lines: List[str]) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", "\n".join(lines))
        widget.configure(state="disabled")

    def clear_preview_table(self) -> None:
        self.preview_table.delete(*self.preview_table.get_children())
        self.preview_table["columns"] = ()

    def render_preview_table(self, column_headers: List[str], preview_rows: List[Dict[str, object]]) -> None:
        columns = ["row_no"] + [f"col_{index}" for index in range(len(column_headers))]
        self.preview_table.delete(*self.preview_table.get_children())
        self.preview_table["columns"] = columns
        self.preview_column_map = {}
        self.preview_table.heading("row_no", text="行号")
        self.preview_table.column("row_no", width=60, anchor="center", stretch=False)

        for index, header in enumerate(column_headers):
            column_id = f"col_{index}"
            self.preview_table.heading(column_id, text=header)
            self.preview_table.column(column_id, width=150, anchor="w", stretch=False)
            header_text = header.split(": ", 1)[0]
            if "列/" in header_text:
                try:
                    column_no = int(header_text.rsplit("/", 1)[1])
                    self.preview_column_map[f"#{index + 2}"] = column_no
                except ValueError:
                    pass

        for row in preview_rows:
            row_no = row["row_no"]
            values = [str(row_no)] + [str(item) for item in row["values"]]
            tags = ()
            if row["kind"] == "header":
                tags = ("header",)
            elif row["kind"] == "footer":
                tags = ("footer",)
            elif row["kind"] == "split":
                tags = ("split",)
            self.preview_table.insert("", "end", values=values, tags=tags)

    def clear_split_preview_table(self) -> None:
        self.split_preview_table.delete(*self.split_preview_table.get_children())

    def sync_key_column_combo(self) -> None:
        current = self.key_column_var.get().strip()
        values = list(self.key_column_combo.cget("values"))
        if not current or not values:
            return
        for value in values:
            if value.startswith(f"{current}. "):
                self.key_column_combo.set(value)
                return

    def sync_key_row_combo(self) -> None:
        current = self.key_row_var.get().strip()
        values = list(self.key_row_combo.cget("values"))
        if not current or not values:
            return
        for value in values:
            if value.startswith(f"{current}. "):
                self.key_row_combo.set(value)
                return

    def render_split_preview_table(self, split_objects: List[Dict[str, object]]) -> None:
        self.clear_split_preview_table()
        if not split_objects:
            self.split_preview_table.insert("", "end", values=("当前条件下没有识别到对象", "-"))
            return
        for item in split_objects:
            self.split_preview_table.insert("", "end", values=(item["name"], item["count"]))

    def parse_int(self, value: str, fallback: Optional[int] = None) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            if fallback is not None:
                return fallback
            raise

    def run_split(self) -> None:
        input_text = self.input_var.get().strip()
        if not input_text:
            messagebox.showwarning("缺少文件", "请先选择待拆分的 Excel 文件。")
            return

        mode = self.mode_var.get()
        try:
            input_path = Path(input_text)
            if mode == "row":
                header_cols = self.parse_int(self.header_cols_var.get())
                footer_cols = self.parse_int(self.footer_cols_var.get(), fallback=0)
                key_row = self.parse_int(self.key_row_var.get())
            else:
                header_rows = self.parse_int(self.header_rows_var.get())
                footer_rows = self.parse_int(self.footer_rows_var.get(), fallback=0)
                key_column = self.parse_int(self.key_column_var.get())
        except ValueError:
            if mode == "row":
                messagebox.showwarning("参数错误", "固定列数和关键行序号必须是整数。")
            else:
                messagebox.showwarning("参数错误", "表头行数、表尾行数和关键列序号必须是整数。")
            return

        output_text = self.output_var.get().strip()
        output_dir = Path(output_text) if output_text else None

        try:
            if mode == "row":
                files = split_workbook_by_row(
                    input_path=input_path,
                    sheet_name=self.sheet_var.get().strip() or None,
                    header_cols=header_cols,
                    key_row=key_row,
                    output_dir=output_dir,
                    footer_cols=footer_cols,
                )
            else:
                files = split_workbook(
                    input_path=input_path,
                    sheet_name=self.sheet_var.get().strip() or None,
                    header_rows=header_rows,
                    key_column=key_column,
                    output_dir=output_dir,
                    footer_rows=footer_rows,
                )
        except Exception as exc:
            messagebox.showerror("拆分失败", str(exc))
            return

        output_parent = files[0].parent if files else (output_dir or input_path.parent / "split_output")
        self.summary_var.set(f"拆分完成，共生成 {len(files)} 个文件，输出目录：{output_parent}")
        self.clear_split_preview_table()
        for path in files:
            self.split_preview_table.insert("", "end", values=(Path(path).name, "已生成"))
        messagebox.showinfo("拆分完成", f"已生成 {len(files)} 个文件。\n输出目录：{output_parent}")

    def run(self) -> None:
        self.root.mainloop()


def run_cli(args: argparse.Namespace) -> None:
    if args.mode == "row":
        files = split_workbook_by_row(
            input_path=Path(args.input),
            sheet_name=args.sheet,
            header_cols=args.header_cols,
            key_row=args.key_row,
            output_dir=Path(args.output_dir) if args.output_dir else None,
            footer_cols=getattr(args, "footer_cols", 0) or 0,
        )
    else:
        files = split_workbook(
            input_path=Path(args.input),
            sheet_name=args.sheet,
            header_rows=args.header_rows,
            key_column=args.key_column,
            output_dir=Path(args.output_dir) if args.output_dir else None,
            footer_rows=getattr(args, "footer_rows", 0) or 0,
        )
    summary = "\n".join(str(path) for path in files)
    print(f"拆分完成，共生成 {len(files)} 个文件：\n{summary}")


def main() -> None:
    args = parse_args()
    if args.input and args.mode == "row" and args.header_cols is not None and args.key_row:
        run_cli(args)
        return
    if args.input and args.mode == "column" and args.header_rows and args.key_column:
        run_cli(args)
        return
    SplitterApp().run()


if __name__ == "__main__":
    main()

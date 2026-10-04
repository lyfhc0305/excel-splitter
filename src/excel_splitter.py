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

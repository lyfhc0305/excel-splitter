import contextlib
import datetime
import io
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils.datetime import CALENDAR_MAC_1904

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))
import excel_splitter as app
import split_core as core


class FormulaTests(unittest.TestCase):
    def test_formula_tokens_and_absolute_references(self):
        result = core.rebuild_formula('=IF(A4="A4",LOG10($B$4),SUM(B1:B6))', {1: 1, 4: 2, 6: 3}, 1, 6)
        self.assertEqual(result, '=IF(A2="A4",LOG10($B$2),SUM(B1:B3))')

    def test_deleted_refs_are_errors_not_other_groups(self):
        self.assertEqual(core.rebuild_formula('=B3+SUM(B2:B3)', {1: 1, 4: 2}), '=#REF!+SUM(#REF!)')

    def test_whole_axis_ranges(self):
        self.assertEqual(core.rebuild_formula('=SUM($2:$6)+SUM(B:B)', {1: 1, 3: 2, 6: 3}), '=SUM($2:$3)+SUM(B:B)')
        self.assertEqual(core.rebuild_formula_by_columns('=SUM($B:$F)+SUM(2:6)', {1: 1, 3: 2, 6: 3}), '=SUM($B:$C)+SUM(2:6)')

    def test_escaped_unicode_sheet_and_other_sheet(self):
        self.assertEqual(core.rebuild_formula("='财务''部'!B4", {4: 1}, sheet_name="财务'部"), "='财务''部'!B1")
        with self.assertRaisesRegex(ValueError, '其他工作表'):
            core.rebuild_formula('=Other!B4', {4: 1}, sheet_name='Sheet')

    def test_dynamic_reference_is_not_silently_corrupted(self):
        with self.assertRaisesRegex(ValueError, '动态引用'):
            core.rebuild_formula('=INDIRECT("B4")', {4: 1})


class NamingAndGroupingTests(unittest.TestCase):
    def test_group_values_follow_excel_display(self):
        values = [3.0, 0.1 + 0.2, True, datetime.datetime(2026, 9, 25), datetime.datetime(2026, 9, 25, 8, 30), ' x ', '']
        self.assertEqual([core.normalize_group_value(value) for value in values],
                         ['3', '0.3', 'TRUE', '2026-09-25', '2026-09-25 08:30:00', 'x', None])

    def test_sheet_titles_are_valid_and_unique(self):
        used = set()
        titles = [core.safe_sheet_title(name, used) for name in ['a/b', 'A_B', "'引号'", 'History', 'x' * 40, 'x' * 40, ' ']]
        self.assertEqual(titles[:4], ['a_b', 'A_B_2', '引号', 'History_'])
        self.assertEqual([len(title) for title in titles[4:6]], [31, 31])
        self.assertNotEqual(titles[4], titles[5])
        self.assertEqual(titles[6], '未命名')

    def test_name_template_rules(self):
        self.assertEqual(core.plan_file_names(core.DEFAULT_NAME_TEMPLATE, '表', 'S', ['甲', 'A/B'], {'表_甲.xlsx'}),
                         ['表_甲_2.xlsx', '表_A_B.xlsx'])
        names = core.plan_file_names('{序号}-{关键字}（{工作表}）', '表', '分配', [str(i) for i in range(10)])
        self.assertEqual(names[0], '01-0（分配）.xlsx')
        self.assertEqual(names[9], '10-9（分配）.xlsx')
        self.assertEqual(core.plan_file_names('{关键字}', '表', 'S', ['CON', 'con']), ['_CON.xlsx', '_con_2.xlsx'])
        for template, message in [('', '不能为空'), ('{文件名}', '必须包含'), ('{关键词}', '未知占位符'), ('{关键字', '花括号')]:
            with self.subTest(template=template), self.assertRaisesRegex(ValueError, message):
                core.validate_name_template(template)

    def test_select_groups_keeps_order_and_rejects_unknown(self):
        groups = {'甲': [2], '乙': [3], '丙': [4]}
        self.assertEqual(list(core.select_groups(groups, ['丙', ' 甲 '])), ['甲', '丙'])
        self.assertIs(core.select_groups(groups, None), groups)
        with self.assertRaisesRegex(ValueError, '找不到拆分对象：丁'):
            core.select_groups(groups, ['丁'])
        with self.assertRaisesRegex(ValueError, '没有选择'):
            core.select_groups(groups, [])


class SplitTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / '源表.xlsx'
        self.out = self.root / 'out'

    def tearDown(self):
        self.temp.cleanup()

    def save(self, rows, title='Sheet'):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = title
        for row in rows:
            ws.append(row)
        wb.save(self.source)
        return wb, ws

    def split(self, **kwargs):
        return app.split_workbook(self.source, None, kwargs.pop('header_rows', 1), 1, self.out, **kwargs)

    def test_totals_fixed_rows_styles_and_dates_roundtrip(self):
        wb, ws = self.save([['部门', '值'], ['甲', 10], ['乙', 20], ['甲', 30], ['合计', '=SUM(B2:B4)']])
        wb.epoch = CALENDAR_MAC_1904
        wb.calculation.calcMode = 'manual'
        ws['B4'].font = Font(name='Arial', bold=True, color='FF112233')
        ws['B4'].number_format = '0.00'
        ws.row_dimensions[4].height = 35
        ws.row_dimensions[4].font = Font(italic=True)
        ws.column_dimensions['B'].width = 26
        ws.column_dimensions['B'].font = Font(bold=True)
        ws.print_area = 'A1:B5'
        ws.print_title_rows = '1:1'
        ws.freeze_panes = 'B2'
        ws.auto_filter.ref = 'A1:B4'
        ws.protection.sheet = True
        ws.row_breaks.append(Break(id=4))
        wb.save(self.source)
        files = self.split(footer_rows=1)
        self.assertEqual(len(files), 2)
        result = openpyxl.load_workbook(files[0])
        try:
            target = result.active
            self.assertEqual(target['B4'].value, '=SUM(B2:B3)')
            self.assertEqual(target['B3'].value, 30)
            self.assertTrue(target['B3'].font.bold)
            self.assertEqual(target['B3'].number_format, '0.00')
            self.assertEqual(target.row_dimensions[3].height, 35)
            self.assertTrue(target.row_dimensions[3].font.italic)
            self.assertTrue(target.column_dimensions['B'].font.bold)
            self.assertIn('$A$1:$B$4', target.print_area)
            self.assertEqual(target.auto_filter.ref, 'A1:B3')
            self.assertEqual(target.freeze_panes, 'B2')
            self.assertTrue(target.protection.sheet)
            self.assertEqual(target.row_breaks.brk[0].id, 3)
            self.assertEqual(result.epoch, wb.epoch)
            self.assertEqual(result.calculation.calcMode, 'auto')
        finally:
            result.close()

    def test_filename_collisions_and_existing_files(self):
        self.save([['部门'], ['A/B'], ['A:B'], ['abc'], ['ABC'], ['x' * 100 + '1'], ['x' * 100 + '2']])
        first = self.split()
        originals = {path: path.read_bytes() for path in first}
        second = self.split()
        self.assertEqual(len(set(path.name.casefold() for path in first + second)), 12)
        self.assertTrue(all(path.read_bytes() == data for path, data in originals.items()))

    def test_blanks_are_kept_and_distinct_from_literal_label(self):
        self.save([['部门', '值'], [None, 10], ['（空白关键字）', 20], [' ', 30]])
        files = self.split()
        values = []
        for file in files:
            wb = openpyxl.load_workbook(file)
            values.extend(row[1] for row in wb.active.iter_rows(min_row=2, values_only=True))
            wb.close()
        self.assertEqual(sorted(values), [10, 20, 30])
        self.assertEqual(len(files), 2)

    def test_merge_anchor_survives_removed_row(self):
        wb, ws = self.save([['部门', '说明'], ['甲', '共用说明'], ['乙'], ['乙']])
        ws.merge_cells('B2:B4')
        ws['B2'].font = Font(bold=True)
        wb.save(self.source)
        files = self.split()
        wb = openpyxl.load_workbook(files[1])
        self.assertEqual(wb.active['B2'].value, '共用说明')
        self.assertTrue(wb.active['B2'].font.bold)
        self.assertIn('B2:B3', wb.active.merged_cells)
        wb.close()

    def test_column_split_totals_merges_and_filter_ids(self):
        wb, ws = self.save([['项目', '甲', '乙', '甲', '合计'], ['销量', 2, 3, 5, '=SUM(B2:D2)'], ['标题', '=B2+D2']])
        ws.merge_cells('B3:D3')
        ws.auto_filter.ref = 'A1:E2'
        ws.auto_filter.add_filter_column(3, ['5'])
        ws.column_dimensions.group('B', 'D', hidden=True)
        ws.print_area = 'A1:E3'
        wb.save(self.source)
        files = app.split_workbook_by_row(self.source, None, 1, 1, self.out, 1)
        wb = openpyxl.load_workbook(files[0])
        self.assertEqual(wb.active['D2'].value, '=SUM(B2:C2)')
        self.assertEqual(wb.active['B3'].value, '=B2+C2')
        self.assertEqual(wb.active.auto_filter.filterColumn[0].colId, 2)
        self.assertTrue(wb.active.column_dimensions['C'].hidden)
        self.assertIn('$A$1:$D$3', wb.active.print_area)
        wb.close()

    def test_no_header_preview_and_output_agree(self):
        self.save([['甲', 1], ['乙', 2], ['甲', 3]])
        preview = app.build_sheet_preview(self.source, 'Sheet', 0, 1)
        self.assertEqual(preview['split_objects'], [{'name': '甲', 'count': 2}, {'name': '乙', 'count': 1}])
        self.assertEqual(len(self.split(header_rows=0)), 2)

    def test_invalid_parameters_rejected_by_preview_and_split(self):
        self.save([['标题', '值'], ['甲', 1]])
        for header, key, footer in [(-1, 1, 0), (1, 0, 0), (1, 3, 0), (1, 1, -1), (1, 1, 1)]:
            for function in (app.split_workbook, app.build_sheet_preview):
                with self.subTest(header=header, key=key, footer=footer, function=function):
                    with self.assertRaises(ValueError):
                        if function is app.split_workbook:
                            function(self.source, 'Sheet', header, key, self.out, footer)
                        else:
                            function(self.source, 'Sheet', header, key, footer)

    def test_formula_group_key_is_rejected(self):
        self.save([['部门'], ['="甲"']])
        with self.assertRaisesRegex(ValueError, 'A2'):
            self.split()

    def test_hidden_sheet_becomes_visible(self):
        wb, ws = self.save([['部门'], ['甲']])
        wb.create_sheet('Visible')
        ws.sheet_state = 'hidden'
        wb.save(self.source)
        file = self.split()[0]
        wb = openpyxl.load_workbook(file)
        self.assertEqual(wb.active.sheet_state, 'visible')
        wb.close()

    def test_literal_formula_like_text_is_not_executed(self):
        wb, ws = self.save([['部门', '内容'], ['甲', '=hello']])
        ws['B2'].data_type = 's'
        wb.save(self.source)
        file = self.split()[0]
        wb = openpyxl.load_workbook(file)
        self.assertEqual(wb.active['B2'].data_type, 's')
        self.assertEqual(wb.active['B2'].value, '=hello')
        wb.close()

    def test_late_save_failure_publishes_nothing(self):
        self.save([['部门'], ['甲'], ['乙']])
        original = openpyxl.Workbook.save
        count = 0
        def failing_save(wb, path):
            nonlocal count
            count += 1
            if count == 2:
                raise OSError('disk full')
            return original(wb, path)
        with patch.object(openpyxl.Workbook, 'save', failing_save):
            with self.assertRaises(OSError):
                self.split()
        self.assertEqual(list(self.out.iterdir()), [])

    def test_rules_and_named_range_follow_rows(self):
        wb, ws = self.save([['部门', '值'], ['甲', 2], ['乙', 3], ['甲', '=SUM(Amounts)']])
        wb.defined_names.add(DefinedName('Amounts', attr_text="'Sheet'!$B$2:$B$3"))
        wb.defined_names.add(DefinedName('Unrelated', attr_text="'Missing'!$A$1"))
        dv = DataValidation(type='whole', operator='between', formula1='0', formula2='10')
        dv.add('B2:B4')
        ws.add_data_validation(dv)
        ws.conditional_formatting.add('B2:B4', FormulaRule(formula=['B2>1'], fill=PatternFill('solid', fgColor='FFFF0000')))
        wb.save(self.source)
        file = self.split()[0]
        wb = openpyxl.load_workbook(file)
        self.assertEqual(wb.defined_names['Amounts'].attr_text, "'Sheet'!$B$2:$B$2")
        self.assertNotIn('Unrelated', wb.defined_names)
        self.assertEqual(str(wb.active.data_validations.dataValidation[0].sqref), 'B2:B3')
        self.assertEqual(str(next(iter(wb.active.conditional_formatting)).sqref), 'B2:B3')
        wb.close()

    def test_template_outputs_real_xlsx(self):
        wb, ws = self.save([['部门'], ['甲']])
        wb.template = True
        self.source = self.source.with_suffix('.xltx')
        wb.save(self.source)
        wb = openpyxl.load_workbook(self.split()[0])
        self.assertFalse(wb.template)
        wb.close()

    def test_unsupported_features_do_not_leave_outputs(self):
        wb, ws = self.save([['部门', '值'], ['甲', '=Other!A1']])
        with self.assertRaisesRegex(ValueError, '其他工作表'):
            self.split()
        self.assertEqual(list(self.out.iterdir()), [])

    def test_missing_sheet_closes_workbook(self):
        self.save([['部门'], ['甲']])
        original = app.close_workbook_compatible
        with patch.object(app, 'close_workbook_compatible', wraps=original) as close:
            with self.assertRaisesRegex(ValueError, '找不到工作表'):
                app.split_workbook(self.source, '不存在', 1, 1, self.out)
            close.assert_called_once()

    def test_publish_failure_rolls_back_only_this_run(self):
        self.save([['部门'], ['甲'], ['乙']])
        self.out.mkdir()
        existing = self.out / 'keep.xlsx'
        existing.write_bytes(b'original')
        original = Path.open
        count = 0
        def failing_open(path, mode='r', *args, **kwargs):
            nonlocal count
            if mode == 'xb':
                count += 1
                if count == 2:
                    raise PermissionError('locked')
            return original(path, mode, *args, **kwargs)
        with patch.object(Path, 'open', failing_open):
            with self.assertRaises(PermissionError):
                self.split()
        self.assertEqual(list(self.out.iterdir()), [existing])
        self.assertEqual(existing.read_bytes(), b'original')

    def test_missing_formula_reference_is_reported(self):
        self.save([['部门', '值'], ['甲', '=B3'], ['乙', 10]])
        result = self.split()
        self.assertEqual(len(result.warnings), 1)
        self.assertIn('B2', result.warnings[0])

    def test_merged_key_and_zero_fixed_columns(self):
        wb, ws = self.save([['甲', None, '乙'], [1, 2, 3]])
        ws.merge_cells('A1:B1')
        wb.save(self.source)
        preview = app.build_sheet_preview_by_row(self.source, 'Sheet', 0, 1)
        files = app.split_workbook_by_row(self.source, None, 0, 1, self.out)
        self.assertEqual([item['count'] for item in preview['split_objects']], [2, 1])
        wb = openpyxl.load_workbook(files[0])
        self.assertEqual(list(wb.active.values), [('甲', None), (1, 2)])
        wb.close()

    def test_preview_size_is_bounded(self):
        wb, ws = self.save([['甲']])
        ws.cell(5000, 1000, '乙')
        wb.save(self.source)
        preview = app.build_sheet_preview(self.source, 'Sheet', 4000, 1000)
        self.assertLess(len(preview['column_headers']), 60)
        self.assertLess(len(preview['preview_rows']), 60)
        self.assertEqual(sum(item['count'] for item in preview['split_objects']), 1000)

    def test_conversion_failure_cleans_temp_directory(self):
        self.source = self.source.with_suffix('.xls')
        self.source.write_bytes(b'old-format')
        converted = tempfile.TemporaryDirectory(dir=self.root)
        directory = Path(converted.name)
        with patch.object(app, 'convert_to_xlsx', return_value=(directory / 'bad.xlsx', converted)):
            with self.assertRaises(FileNotFoundError):
                app.load_workbook_compatible(self.source)
        self.assertFalse(directory.exists())

    def test_converter_timeout_tries_fallback(self):
        def fallback(source, output):
            result = output / 'converted.xlsx'
            result.write_bytes(b'converted')
            return result
        with patch.object(app, 'convert_with_libreoffice', side_effect=app.subprocess.TimeoutExpired('lo', 120)):
            with patch.object(app, 'convert_with_windows_com', side_effect=fallback) as convert:
                result, temporary = app.convert_to_xlsx(self.source)
                try:
                    self.assertEqual(result.read_bytes(), b'converted')
                    convert.assert_called_once()
                finally:
                    temporary.cleanup()


    def test_xlsx_core_et_and_macro_free_xlsm_split(self):
        # These load with a VBA archive attached. A sheet without <pageSetup> keeps openpyxl's
        # back-link from page_setup to the sheet, which must not drag the workbook into a copy.
        for orientation in (None, 'landscape'):
            wb, ws = self.save([['部门', '值'], ['甲', 1], ['乙', 2], ['甲', '=B2*2']])
            if orientation:
                ws.page_setup.orientation = orientation
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.print_title_rows = '1:1'
            ws.auto_filter.ref = 'A1:B4'
            ws.freeze_panes = 'A2'
            ws['A2'].comment = Comment('备注', '作者')
            ws['A3'].hyperlink = '#Sheet!A1'
            dv = DataValidation(type='whole', operator='between', formula1='0', formula2='10')
            dv.add('B2:B4')
            ws.add_data_validation(dv)
            ws.conditional_formatting.add('B2:B4', FormulaRule(formula=['B2>1'], fill=PatternFill('solid', fgColor='FFFF0000')))
            wb.defined_names.add(DefinedName('Amounts', attr_text="'Sheet'!$B$2:$B$3"))
            ws['C2'] = '=SUM(Amounts)'
            wb.save(self.source)
            for suffix in ('.et', '.xlsm'):
                source = self.source.with_name(f'源表{suffix}')
                shutil.copy(self.source, source)
                self.assertEqual(app.build_sheet_preview(source, 'Sheet', 1, 1)['group_count'], 2)
                for single in (False, True):
                    with self.subTest(orientation=orientation, suffix=suffix, single=single):
                        folder = self.out / f'{orientation}{suffix}{single}'
                        files = app.split_workbook(source, None, 1, 1, folder, single_workbook=single)
                        result = openpyxl.load_workbook(files[0])
                        target = result.worksheets[0]
                        self.assertEqual([row[1] for row in target.iter_rows(min_row=2, values_only=True)], [1, '=B2*2'])
                        self.assertEqual(target.page_setup.orientation, orientation)
                        self.assertTrue(target.sheet_properties.pageSetUpPr.fitToPage)
                        self.assertEqual(target['A2'].comment.text, '备注')
                        result.close()

    def test_page_setup_links_to_output_sheet(self):
        self.save([['部门'], ['甲'], ['乙']])
        wb = openpyxl.load_workbook(self.source)
        target = core.build_target(wb.active, [2], 1, wb).active
        self.assertIs(target.page_setup._parent, target)
        wb.close()

    def test_sheet_named_like_default_keeps_title(self):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet('sheet')
        for row in [['部门', '值'], ['甲', '=sheet!B3'], ['甲', 5]]:
            ws.append(row)
        wb.save(self.source)
        result = openpyxl.load_workbook(self.split()[0])
        self.assertEqual(result.sheetnames, ['sheet'])
        self.assertEqual(result.active['B2'].value, '=sheet!B3')
        result.close()

    def test_selected_groups_and_name_template(self):
        self.save([['部门', '值'], ['甲', 1], ['乙', 2], ['丙', 3], ['乙', 4]])
        files = self.split(groups=['丙', '乙'], name_template='{序号}_{关键字}_{工作表}')
        self.assertEqual([path.name for path in files], ['1_乙_Sheet.xlsx', '2_丙_Sheet.xlsx'])
        result = openpyxl.load_workbook(files[0])
        self.assertEqual([row[1] for row in result.active.iter_rows(min_row=2, values_only=True)], [2, 4])
        result.close()
        with self.assertRaisesRegex(ValueError, '找不到拆分对象'):
            self.split(groups=['丁'])
        with self.assertRaisesRegex(ValueError, '未知占位符'):
            self.split(name_template='{部门}')

    def test_single_workbook_has_one_sheet_per_group(self):
        wb, ws = self.save([['部门', '值'], ['甲', 2], ['乙/丙', 3], ['甲', '=SUM(Amounts)+Sheet!B2'], ['合计', '=SUM(B2:B4)']])
        wb.defined_names.add(DefinedName('Amounts', attr_text="'Sheet'!$B$2:$B$3"))
        ws.print_area = 'A1:B5'
        ws.auto_filter.ref = 'A1:B4'
        wb.save(self.source)
        files = self.split(footer_rows=1, single_workbook=True)
        self.assertEqual([path.name for path in files], ['源表_拆分.xlsx'])
        self.assertEqual(files.sheet_titles, ['甲', '乙_丙'])
        result = openpyxl.load_workbook(files[0])
        try:
            first, second = result.worksheets
            self.assertEqual([[cell.value for cell in row] for row in first.iter_rows()],
                             [['部门', '值'], ['甲', 2], ['甲', "=SUM(Amounts)+'甲'!B2"], ['合计', '=SUM(B2:B3)']])
            self.assertEqual(second['B3'].value, '=SUM(B2:B2)')
            self.assertEqual(first.defined_names['Amounts'].attr_text, "'甲'!$B$2:$B$2")
            self.assertNotIn('Amounts', result.defined_names)
            self.assertIn("'乙_丙'!$A$1:$B$3", second.print_area)
            self.assertEqual(second.auto_filter.ref, 'A1:B2')
            self.assertEqual([sheet.views.sheetView[0].tabSelected for sheet in result.worksheets], [True, False])
        finally:
            result.close()

    def test_single_workbook_by_columns(self):
        self.save([['项目', '甲', '乙', '甲'], ['销量', 1, 2, 3]])
        files = app.split_workbook_by_row(self.source, None, 1, 1, self.out, single_workbook=True)
        result = openpyxl.load_workbook(files[0])
        self.assertEqual({sheet.title: list(sheet.values) for sheet in result.worksheets},
                         {'甲': [('项目', '甲', '甲'), ('销量', 1, 3)], '乙': [('项目', '乙'), ('销量', 2)]})
        result.close()

    def test_progress_reports_each_group_and_cancel_publishes_nothing(self):
        self.save([['部门'], ['甲'], ['乙'], ['丙']])
        seen = []
        self.split(progress=lambda done, total, name: seen.append((done, total, name)))
        self.assertEqual(seen, [(0, 3, '甲'), (1, 3, '乙'), (2, 3, '丙'), (3, 3, None)])
        cancel = threading.Event()

        def stop_after_first(done, total, name):
            if done == 1:
                cancel.set()

        before = sorted(self.out.iterdir())
        with self.assertRaises(core.SplitCancelled):
            self.split(progress=stop_after_first, cancel=cancel)
        self.assertEqual(sorted(self.out.iterdir()), before)

    def test_workbook_cache_reuses_until_file_changes(self):
        self.save([['部门'], ['甲']])
        cache = app.WorkbookCache()
        try:
            first = cache.get(self.source)
            self.assertIs(cache.get(self.source), first)
            self.assertEqual(app.build_sheet_preview(self.source, 'Sheet', 1, 1, cache=cache)['group_count'], 1)
            self.assertIs(cache.get(self.source), first)
            self.save([['部门'], ['甲'], ['乙']])
            os.utime(self.source, ns=(time.time_ns(), time.time_ns() + 10**9))
            self.assertIsNot(cache.get(self.source), first)
            self.assertEqual(len(app.split_workbook(self.source, None, 1, 1, self.out, cache=cache)), 2)
        finally:
            cache.clear()
        with self.assertRaisesRegex(ValueError, '找不到输入文件'):
            cache.get(self.root / 'missing.xlsx')

    def test_preview_lists_every_key_column_and_marks_gaps(self):
        wb, ws = self.save([[f'列{index}' for index in range(1, 41)]] + [[f'值{row}'] + [row] * 39 for row in range(40)])
        ws.append([None, 1])
        wb.save(self.source)
        preview = app.build_sheet_preview(self.source, 'Sheet', 1, 1)
        self.assertEqual([number for number, _label in preview['key_choices']], list(range(1, 41)))
        self.assertEqual(preview['key_choices'][39][1], 'AN · 列40')
        self.assertLess(len(preview['preview_columns']), 40)
        self.assertEqual(preview['preview_rows'][-1]['kind'], 'gap')
        self.assertEqual(preview['data_count'], 41)
        self.assertEqual(preview['blank_group'], '（空白关键字）')
        self.assertTrue(preview['column_headers'][0].startswith('★ A'))


class InterfaceTests(unittest.TestCase):
    def test_cli_new_options(self):
        args = app.parse_args(['--input', 'x.xlsx', '--header-rows', '1', '--key-column', '2', '--group', '甲',
                               '--group', '乙', '--name-template', '{关键字}', '--list-groups'])
        self.assertEqual((args.group, args.name_template, args.list_groups), (['甲', '乙'], '{关键字}', True))
        for extra in (['--single-workbook', '--name-template', '{关键字}'], ['--name-template', '{部门}']):
            with self.subTest(extra=extra), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit):
                    app.parse_args(['--input', 'x.xlsx', '--header-rows', '1', '--key-column', '1'] + extra)

    def test_cli_list_groups_writes_nothing(self):
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / '表.xlsx'
            wb = openpyxl.Workbook()
            for row in [['部门'], ['甲'], ['乙'], ['甲']]:
                wb.active.append(row)
            wb.save(source)
            args = app.parse_args(['--input', str(source), '--header-rows', '1', '--key-column', '1', '--list-groups'])
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                app.run_cli(args)
            self.assertEqual(output.getvalue().splitlines()[1:], ['关键字\t行数', '甲\t2', '乙\t1'])
            self.assertEqual(sorted(path.name for path in Path(folder).iterdir()), ['表.xlsx'])

    def test_cli_zero_headers_and_missing_parameters(self):
        self.assertEqual(app.parse_args(['--input', 'x.xlsx', '--header-rows', '0', '--key-column', '1']).header_rows, 0)
        with contextlib.redirect_stderr(io.StringIO()):
            with self.assertRaises(SystemExit) as error:
                app.parse_args(['--input', 'x.xlsx'])
        self.assertEqual(error.exception.code, 2)

    def test_windows_filename_rules(self):
        for name in ['CON', 'aux.txt', 'LPT1', 'com¹']:
            self.assertTrue(core.safe_file_name(name).startswith('_'))
        self.assertEqual(core.safe_file_name('x. '), 'x')
        self.assertEqual(core.safe_file_name('a\x01b'), 'a_b')


class BackgroundInterfaceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.folder = Path(self.temp.name)
        self.settings = self.folder / 'settings.json'
        self.gui = self.open_gui()

    def open_gui(self):
        try:
            self.root = app.tk.Tk()
        except app.tk.TclError as exc:
            self.skipTest(f'Tk display unavailable: {exc}')
        self.root.withdraw()
        with patch.object(app.tk, 'Tk', return_value=self.root):
            return app.SplitterApp(settings_path=self.settings)

    def tearDown(self):
        if getattr(self, 'gui', None) is not None:
            self.gui._splitting = False
            self.gui._close()
        self.temp.cleanup()

    def make_source(self, name='源表.xlsx', rows=None):
        path = self.folder / name
        wb = openpyxl.Workbook()
        wb.active.title = '分配'
        for row in rows or [['部门', '值'], ['甲', 1], ['乙', 2], ['丙', 3], ['甲', 4]]:
            wb.active.append(row)
        wb.save(path)
        return path

    def load(self, path):
        self.gui.header_rows_var.set('1')
        self.gui.footer_rows_var.set('0')
        self.gui.key_column_var.set('1')
        self.gui.load_input_path(path)
        self.pump_until(lambda: self.gui.groups is not None and not self.gui._preview_after)

    def pump_until(self, condition):
        deadline = time.monotonic() + 5
        while not condition() and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.005)
        self.assertTrue(condition(), 'background result did not reach the interface')

    def test_background_job_keeps_ui_responsive_and_discards_stale_preview(self):
        started, release = threading.Event(), threading.Event()
        delivered, responsive = [], []
        def slow():
            started.set()
            release.wait(3)
            return 'old'
        try:
            self.gui._start_job('preview', slow, lambda value, error: delivered.append(value))
            self.pump_until(started.is_set)
            self.root.after(1, lambda: responsive.append(True))
            self.pump_until(lambda: bool(responsive))
            self.gui._start_job('preview', lambda: 'new', lambda value, error: delivered.append(value))
            release.set()
            self.pump_until(lambda: bool(delivered))
            self.assertEqual(delivered, ['new'])
        finally:
            release.set()

    def test_failed_split_restores_controls(self):
        self.gui.input_var.set('nonexistent.xlsx')
        with patch.object(app.messagebox, 'showerror') as show:
            self.gui.run_split()
            self.assertTrue(self.gui._splitting)
            self.pump_until(lambda: not self.gui._splitting)
            show.assert_called_once()
            self.assertEqual(str(self.gui.sheet_combo.cget('state')), 'readonly')

    def test_group_selection_and_single_workbook_split(self):
        source = self.make_source()
        self.load(source)
        self.assertEqual([item['name'] for item in self.gui.groups], ['甲', '乙', '丙'])
        first = next(item for item, name in self.gui.group_items.items() if name == '乙')
        self.gui._toggle_groups([first])
        self.assertEqual(self.gui._selected_names(), ['甲', '丙'])
        self.assertEqual(self.gui.split_preview_table.set(first, 'output'), '不输出')
        self.gui.output_mode_var.set('workbook')
        self.gui.update_group_names()
        self.assertEqual(self.gui.plan_label.cget('text'), '将生成 1 个工作簿（2 个工作表）')
        with patch.object(app.messagebox, 'askyesno', return_value=False) as ask, \
                patch.object(app.messagebox, 'showerror') as show_error:
            self.gui.run_split()
            self.pump_until(lambda: not self.gui._splitting)
            show_error.assert_not_called()
            ask.assert_called_once()
        result = openpyxl.load_workbook(source.parent / 'split_output' / '源表_拆分.xlsx')
        self.assertEqual(result.sheetnames, ['甲', '丙'])
        result.close()
        self.assertEqual(json.loads(self.settings.read_text(encoding='utf-8'))['output_mode'], 'workbook')

    def test_cancelled_split_publishes_nothing(self):
        source = self.make_source()
        self.load(source)
        started, proceed = threading.Event(), threading.Event()
        original = app.save_groups

        def paused_save(*args, **kwargs):
            started.set()
            proceed.wait(3)
            return original(*args, **kwargs)

        with patch.object(app, 'save_groups', paused_save), patch.object(app.messagebox, 'showerror') as show_error:
            self.gui.run_split()
            self.pump_until(started.is_set)
            self.gui.cancel_split()
            proceed.set()
            self.pump_until(lambda: not self.gui._splitting)
            show_error.assert_not_called()
        self.assertIn('已取消', self.gui.status_var.get())
        self.assertEqual(list((source.parent / 'split_output').glob('*.xlsx')), [])
        self.assertEqual(str(self.gui.start_button.cget('state')), 'normal')

    def test_invalid_template_disables_start(self):
        self.load(self.make_source())
        self.gui.name_template_var.set('{部门}')
        self.assertEqual(str(self.gui.start_button.cget('state')), 'disabled')
        self.assertIn('未知占位符', self.gui.template_hint.cget('text'))
        self.gui.name_template_var.set('{序号}')
        self.assertEqual(str(self.gui.start_button.cget('state')), 'normal')

    def test_output_dir_follows_input_until_customized(self):
        first = self.make_source('一.xlsx')
        (self.folder / 'sub').mkdir()
        second = self.make_source('sub/二.xlsx')
        self.gui.load_input_path(first)
        self.assertEqual(self.gui.output_var.get(), str(self.folder / 'split_output'))
        self.gui.load_input_path(second)
        self.assertEqual(self.gui.output_var.get(), str(self.folder / 'sub' / 'split_output'))
        self.gui.output_var.set(str(self.folder / 'custom'))
        self.gui.load_input_path(first)
        self.assertEqual(self.gui.output_var.get(), str(self.folder / 'custom'))

    def test_settings_are_remembered(self):
        self.gui.mode_var.set('row')
        self.gui.header_cols_var.set('2')
        self.gui.name_template_var.set('{序号}_{关键字}')
        self.gui._close()
        self.gui = self.open_gui()
        self.assertEqual((self.gui.mode_var.get(), self.gui.header_cols_var.get(), self.gui.name_template_var.get()),
                         ('row', '2', '{序号}_{关键字}'))
        for content in ('not json', '{"mode": "sideways", "header_rows": "-1", "name_template": 5}'):
            self.gui._close()
            self.settings.write_text(content, encoding='utf-8')
            self.gui = self.open_gui()
            self.assertEqual((self.gui.mode_var.get(), self.gui.header_rows_var.get(), self.gui.name_template_var.get()),
                             ('column', '1', core.DEFAULT_NAME_TEMPLATE))

    def test_key_row_highlight_is_rendered(self):
        self.gui.render_preview_table(['A列/1: 甲'], [{'row_no': 1, 'kind': 'keyrow', 'values': ['甲']}])
        item = self.gui.preview_table.get_children()[0]
        self.assertEqual(self.gui.preview_table.item(item, 'tags'), ('keyrow',))


if __name__ == '__main__':
    unittest.main()

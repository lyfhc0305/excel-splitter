from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple


class _App:
    """Look up excel_splitter attributes when called, so test patches on that module apply."""

    def __getattr__(self, name):
        import excel_splitter
        return getattr(excel_splitter, name)


app = _App()


class SplitterMixin3:
    def update_group_names(self) -> None:
        """Refresh checkboxes, planned output names, counts and the start button."""
        tree = self.split_preview_table
        selected = self._selected_names()
        single = self.output_mode_var.get() == "workbook"
        input_text = self.input_var.get().strip()
        stem = app.safe_file_name(Path(input_text).stem) if input_text else "源文件名"
        planned, error = {}, None
        reserved = app.occupied_output_names(Path(input_text), self._output_dir()) if input_text else set()
        if single:
            planned = dict(zip(selected, app.plan_sheet_titles(selected)))
        else:
            try:
                sheet = (self._preview_info or {}).get("sheet_title", self.sheet_var.get())
                planned = dict(zip(selected, app.plan_file_names(self.name_template_var.get(), Path(input_text).stem or "源文件名",
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
            workbook_name = app.unique_name(f"{stem}_拆分", set(reserved))
            self.workbook_hint.configure(text=f"{workbook_name}（每个对象一个工作表）")
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

    def on_group_click(self, event: app.tk.Event) -> None:
        tree = self.split_preview_table
        if self._splitting or tree.identify("region", event.x, event.y) not in ("tree", "cell"):
            return
        item = tree.identify_row(event.y)
        if item in self.group_items:
            self._toggle_groups([item])

    def on_group_space(self, _event: app.tk.Event) -> str:
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
                    if isinstance(child, (app.ttk.Entry, app.ttk.Button, app.ttk.Radiobutton)) and child not in (self.cancel_button, self.help_button):
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
            app.messagebox.showwarning("缺少文件", "请先选择待拆分的 Excel 文件。", parent=self.root)
            return
        mode = self.mode_var.get()
        try:
            leading, trailing, key = self._read_parameters(mode)
        except ValueError:
            app.messagebox.showwarning("参数错误", "固定行列数和关键行列序号必须是整数。", parent=self.root)
            return
        single = self.output_mode_var.get() == "workbook"
        template = self.name_template_var.get()
        if not single:
            try:
                app.validate_name_template(template)
            except ValueError as exc:
                app.messagebox.showwarning("文件命名规则有误", str(exc), parent=self.root)
                return
        selected = self._selected_names() if self.groups is not None and self.excluded else None
        if selected == []:
            app.messagebox.showwarning("未选择拆分对象", "请至少勾选一个拆分对象。", parent=self.root)
            return
        split_function = app.split_workbook_by_row if mode == "row" else app.split_workbook
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
            if isinstance(error, app.SplitCancelled):
                self._set_status("已取消拆分，未生成任何文件。", "warning")
                return
            if error:
                self._set_status(f"拆分失败：{error}", "error")
                app.messagebox.showerror("拆分失败", str(error), parent=self.root)
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
            if app.messagebox.askyesno(title, message + "\n\n是否打开输出目录？", icon=icon, parent=self.root):
                self.open_output_dir(output_parent)

        self._start_job("split", work, done)

    def run(self) -> None:
        self.root.mainloop()



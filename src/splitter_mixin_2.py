from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple


class _App:
    """Look up excel_splitter attributes when called, so test patches on that module apply."""

    def __getattr__(self, name):
        import excel_splitter
        return getattr(excel_splitter, name)


app = _App()


class SplitterMixin2:
    def _build_groups_card(self, parent: app.tk.Widget) -> app.ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "拆分对象")
        self.group_summary = app.ttk.Label(head, style="Muted.TLabel")
        self.group_summary.grid(row=0, column=1, sticky="w", padx=(px(12), 0))
        tools = app.ttk.Frame(head, style="CardBody.TFrame")
        tools.grid(row=0, column=2, sticky="e")
        for text, command in (("全选", self.select_all_groups), ("全不选", self.select_no_groups), ("反选", self.invert_groups)):
            app.ttk.Button(tools, text=text, style="Link.TButton", command=command).pack(side="left", padx=(px(2), 0))
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
        tree.tag_configure("excluded", foreground=app.PALETTE["disabled"])
        tree.tag_configure("blank", foreground=app.PALETTE["warning"])
        tree.bind("<ButtonRelease-1>", self.on_group_click)
        tree.bind("<space>", self.on_group_space)
        self.groups_message = app.ttk.Label(tree.master, style="Empty.TLabel", justify="center", anchor="center")
        self._show_groups_message("生成预览后，在这里勾选需要输出的拆分对象")

        options = app.ttk.Frame(body, style="CardBody.TFrame")
        options.grid(row=1, column=0, sticky="ew", pady=(px(10), 0))
        options.columnconfigure(3, weight=1)
        app.ttk.Label(options, text="输出方式", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, px(8)))
        self._segmented(
            options, self.output_mode_var, (("files", "每个对象一个文件"), ("workbook", "合并为一个工作簿")),
            self.update_group_names,
        ).grid(row=0, column=1, sticky="w")
        self.template_label = app.ttk.Label(options, text="文件命名", style="Field.TLabel")
        self.template_label.grid(row=0, column=2, sticky="w", padx=(px(20), px(8)))
        self.template_entry = app.ttk.Entry(options, textvariable=self.name_template_var, width=12)
        self.template_entry.grid(row=0, column=3, sticky="ew")
        app.Tooltip(self.template_entry, app.TIPS["template"], self)
        self.workbook_hint = app.ttk.Label(options, style="Field.TLabel")
        self.workbook_hint.grid(row=0, column=3, sticky="w")
        self.workbook_hint.grid_remove()
        self.template_hint = app.ttk.Label(options, style="Error.TLabel", justify="left")
        self.template_hint.grid(row=1, column=0, columnspan=4, sticky="e", pady=(px(4), 0))
        self.template_hint.grid_remove()
        options.bind("<Configure>", lambda event: self.template_hint.configure(wraplength=max(event.width, 1)))
        return card

    def _build_status_bar(self) -> None:
        px = self.px
        app.tk.Frame(self.root, height=1, bg=app.PALETTE["border"]).grid(row=3, column=0, sticky="ew")
        bar = app.ttk.Frame(self.root, style="Bar.TFrame", padding=(px(20), px(10)))
        bar.grid(row=4, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)
        self.status_dot = app.ttk.Label(bar, text="●", style="Idle.Dot.TLabel")
        self.status_dot.grid(row=0, column=0, padx=(0, px(8)))
        self.status_label = app.ttk.Label(bar, textvariable=self.status_var, style="Status.TLabel", width=1)
        self.status_label.grid(row=0, column=1, sticky="ew")
        self.status_label.bind("<Configure>", lambda event: self.status_label.configure(wraplength=max(event.width, 1)))
        self.progress_bar = app.ttk.Progressbar(bar, style="Accent.Horizontal.TProgressbar", length=px(220))
        self.progress_bar.grid(row=0, column=2, padx=(px(12), px(10)))
        self.progress_bar.grid_remove()
        self.cancel_button = app.ttk.Button(bar, text="取消", command=self.cancel_split)
        self.cancel_button.grid(row=0, column=3, padx=(0, px(10)))
        self.cancel_button.grid_remove()
        self.plan_label = app.ttk.Label(bar, style="Muted.TLabel")
        self.plan_label.grid(row=0, column=4, padx=(px(12), px(14)))
        self.start_button = app.ttk.Button(bar, text="开始拆分", style="Accent.TButton", command=self.run_split)
        self.start_button.grid(row=0, column=5)
        self._update_start_state()

    def _bind_shortcuts(self) -> None:
        modifier = "Command" if app.sys.platform == "darwin" else "Control"
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
            if marker in app.PALETTE:
                app.tk.Frame(self.legend, width=px(12), height=px(12), bg=app.PALETTE[marker], highlightthickness=1,
                         highlightbackground=app.PALETTE[marker + "_edge"]).pack(side="left", padx=(px(14), px(5)))
            else:
                app.ttk.Label(self.legend, text=marker, style="Field.TLabel").pack(side="left", padx=(px(14), px(3)))
            app.ttk.Label(self.legend, text=text, style="Muted.TLabel").pack(side="left")

    def show_help(self) -> None:
        app.messagebox.showinfo("使用说明", app.HELP_TEXT, parent=self.root)

    # ---- 文件与目录 ----

    def choose_input(self) -> None:
        path = app.filedialog.askopenfilename(
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
        path = app.filedialog.askdirectory(title="选择输出目录", initialdir=initial_dir, mustexist=False)
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
            app.messagebox.showinfo("输出目录", "请先选择 Excel 文件或输出目录。", parent=self.root)
            return
        if not target.is_dir():
            app.messagebox.showinfo("输出目录", f"目录尚未创建：{target}\n开始拆分后会自动创建。", parent=self.root)
            return
        try:
            app.open_path(target)
        except OSError as exc:
            app.messagebox.showerror("无法打开目录", str(exc), parent=self.root)

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
        except app.queue.Empty:
            pass
        finally:
            latest = None
            try:
                while True:
                    latest = self._progress.get_nowait()
            except app.queue.Empty:
                pass
            if latest is not None and self._splitting:
                self._show_progress(*latest)
            if not self._closed:
                self._poll_after = self.root.after(75, self._poll_jobs)

    def _close(self):
        if self._splitting:
            if app.messagebox.askyesno(
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
        app.save_settings(self.settings_path, {
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
                app.messagebox.showerror("读取失败", str(error), parent=self.root)
                return
            sheet_names = info["sheet_names"]
            self.sheet_combo["values"] = sheet_names
            if self.sheet_var.get().strip() not in sheet_names:
                self.sheet_var.set(sheet_names[0])
            self.refresh_sheet_preview()

        def work():
            if force:
                self._cache.clear()
            return app.load_workbook_info(Path(input_path), self._cache)

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
        except (ValueError, app.tk.TclError):
            combo.set("")

    def _on_key_selected(self, mode: str, combo: app.ttk.Combobox, variable: app.tk.StringVar) -> None:
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
            self.preview_hint.configure(text=app.TIPS["row_mode"])
            self.split_preview_table.heading("count", text="列数")
        else:
            self.row_options.grid_remove()
            self.column_options.grid()
            self.mode_hint.configure(text="关键列相同的数据行放入同一个文件，表头表尾复制到每个文件。")
            self._set_legend([("row_header", "表头"), ("row_footer", "表尾"), ("★", "关键列"), ("⋯", "省略")])
            self.preview_hint.configure(text=app.TIPS["column_mode"])
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
        preview_function = app.build_sheet_preview_by_row if mode == "row" else app.build_sheet_preview
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

    def on_preview_table_click(self, event: app.tk.Event) -> None:
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

    def on_preview_context_menu(self, event: app.tk.Event) -> None:
        info = self._preview_info
        table = self.preview_table
        region = table.identify("region", event.x, event.y)
        if self._splitting or not info or region not in ("heading", "cell"):
            return
        column = self.preview_column_map.get(table.identify_column(event.x))
        row = self.preview_row_map.get(table.identify_row(event.y)) if region == "cell" else None
        menu = app.tk.Menu(self.root, **self.menu_options)
        if self.mode_var.get() == "row":
            if row is not None:
                menu.add_command(label=f"设为关键行（第 {row} 行）", command=lambda: self.key_row_var.set(str(row)))
            if column is not None:
                letter, right = app.get_column_letter(column), info["max_column"] - column + 1
                menu.add_command(label=f"左侧固定到 {letter} 列（共 {column} 列）",
                                 command=lambda: self.header_cols_var.set(str(column)))
                menu.add_command(label=f"从 {letter} 列起固定到最右侧（共 {right} 列）",
                                 command=lambda: self.footer_cols_var.set(str(right)))
        else:
            if column is not None:
                menu.add_command(label=f"设为关键列（{app.get_column_letter(column)} 列）",
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


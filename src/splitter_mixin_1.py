from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple


class _App:
    """Look up excel_splitter attributes when called, so test patches on that module apply."""

    def __getattr__(self, name):
        import excel_splitter
        return getattr(excel_splitter, name)


app = _App()


class SplitterMixin1:
    def __init__(self, settings_path: Optional[Path] = None) -> None:
        app.enable_dpi_awareness()
        self.root = app.tk.Tk()
        self.root.title(app.APP_TITLE)
        self.scale = min(3.0, max(1.0, self.root.winfo_fpixels("1i") / 96))
        self.settings_path = Path(settings_path) if settings_path else app.default_settings_path()
        settings = app.load_settings(self.settings_path)

        def number(key: str, default: str) -> str:
            value = settings.get(key, "")
            return value if value.isdigit() else default

        def choice(key: str, options: Tuple[str, ...]) -> str:
            return settings.get(key) if settings.get(key) in options else options[0]

        self.input_var = app.tk.StringVar()
        self.sheet_var = app.tk.StringVar()
        self.mode_var = app.tk.StringVar(value=choice("mode", ("column", "row")))
        self.header_rows_var = app.tk.StringVar(value=number("header_rows", "1"))
        self.footer_rows_var = app.tk.StringVar(value=number("footer_rows", "0"))
        self.key_column_var = app.tk.StringVar(value=number("key_column", "1"))
        self.header_cols_var = app.tk.StringVar(value=number("header_cols", "1"))
        self.footer_cols_var = app.tk.StringVar(value=number("footer_cols", "0"))
        self.key_row_var = app.tk.StringVar(value=number("key_row", "1"))
        self.output_var = app.tk.StringVar()
        self.output_mode_var = app.tk.StringVar(value=choice("output_mode", ("files", "workbook")))
        self.name_template_var = app.tk.StringVar(value=settings.get("name_template") or app.DEFAULT_NAME_TEMPLATE)
        self.status_var = app.tk.StringVar()
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

        self._jobs = app.queue.Queue()
        self._results = app.queue.Queue()
        self._progress = app.queue.Queue()
        self._versions = {}
        self._cache = app.WorkbookCache()
        self._cancel = app.threading.Event()
        self._preview_after = None
        self._splitting = False
        self._closed = False
        self._close_requested = False
        self._place_window()
        self._build_ui()
        self.on_mode_changed(refresh=False)
        self._set_status("先选择 Excel 文件，界面会自动读取工作表并生成预览。")
        app.threading.Thread(target=self._worker, daemon=True).start()
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
        c, px = app.PALETTE, self.px
        family = app.pick_font_family(self.root)
        size = 13 if app.sys.platform == "darwin" else 10
        self.fonts = {
            "base": (family, size),
            "small": (family, size - 1),
            "bold": (family, size, "bold"),
            "card": (family, size + 1, "bold"),
            "title": (family, size + 5, "bold"),
            "button": (family, size + 1, "bold"),
        }
        self.table_font = app.tkfont.Font(root=self.root, family=family, size=size)
        self.heading_font = app.tkfont.Font(root=self.root, family=family, size=size, weight="bold")
        if app.sys.platform != "darwin":
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkTooltipFont"):
                app.tkfont.nametofont(name, self.root).configure(family=family, size=size)
        self.root.option_add("*TCombobox*Listbox.font", self.fonts["base"])
        self.root.option_add("*TCombobox*Listbox.background", c["field"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", c["accent_soft"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", c["text"])
        self.menu_options = dict(
            tearoff=0, font=self.fonts["base"], bg=c["card"], fg=c["text"], activebackground=c["accent_soft"],
            activeforeground=c["accent"], relief="solid", borderwidth=1,
        )

        style = app.ttk.Style(self.root)
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

    def _card(self, parent: app.tk.Widget, title: str):
        px = self.px
        card = app.ttk.Frame(parent, style="Card.TFrame", padding=(px(16), px(12), px(16), px(14)))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        head = app.ttk.Frame(card, style="CardBody.TFrame")
        head.grid(row=0, column=0, sticky="ew", pady=(0, px(10)))
        head.columnconfigure(1, weight=1)
        app.ttk.Label(head, text=title, style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        body = app.ttk.Frame(card, style="CardBody.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        return card, head, body

    def _build_ui(self) -> None:
        self._setup_style()
        px = self.px
        self.root.configure(bg=app.PALETTE["bg"])
        self.app_icon = app.draw_app_icon(self.root, 64)
        self.header_icon = app.draw_app_icon(self.root, px(30))
        self.check_images = {state: app.draw_checkbox(self.root, max(12, px(15)), state) for state in ("on", "off", "mixed")}
        try:
            self.root.iconphoto(True, self.app_icon)
        except app.tk.TclError:
            pass
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(2, weight=1)

        header = app.ttk.Frame(self.root, style="Bar.TFrame", padding=(px(20), px(10)))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(2, weight=1)
        app.ttk.Label(header, image=self.header_icon).grid(row=0, column=0, padx=(0, px(10)))
        app.ttk.Label(header, text=app.APP_TITLE, style="Title.TLabel").grid(row=0, column=1, sticky="w")
        app.ttk.Label(
            header, style="Subtitle.TLabel",
            text="按关键列拆分数据行，或按关键行拆分数据列；保留样式、公式、合并单元格和打印设置",
        ).grid(row=0, column=2, sticky="w", padx=(px(14), 0))
        self.help_button = app.ttk.Button(header, text="使用说明", style="Link.TButton", command=self.show_help)
        self.help_button.grid(row=0, column=3, sticky="e")
        app.tk.Frame(self.root, height=1, bg=app.PALETTE["border"]).grid(row=1, column=0, sticky="ew")

        body = app.ttk.Frame(self.root, padding=(px(16), px(14)))
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, minsize=px(340))
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        sidebar = app.ttk.Frame(body)
        sidebar.grid(row=0, column=0, sticky="nsew", padx=(0, px(14)))
        sidebar.columnconfigure(0, weight=1)
        workspace = app.ttk.Frame(body)
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

    def _build_file_card(self, parent: app.tk.Widget) -> app.ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "① 选择文件")
        open_button = app.ttk.Button(head, text="打开输出目录", style="Link.TButton", command=self.open_output_dir)
        open_button.grid(row=0, column=2, sticky="e")
        app.Tooltip(open_button, app.TIPS["open"], self)
        body.columnconfigure(1, weight=1)
        app.ttk.Label(body, text="源文件", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, px(8)))
        self.input_entry = app.ttk.Entry(body, textvariable=self.input_var, width=12)
        self.input_entry.grid(row=0, column=1, sticky="ew", padx=(0, px(8)))
        self.input_entry.bind("<Return>", lambda _event: self.reload_workbook_info(force=True))
        self.input_entry.bind("<FocusOut>", self._on_input_focus_out)
        app.ttk.Button(body, text="选择…", command=self.choose_input).grid(row=0, column=2, sticky="ew")
        app.ttk.Label(body, text="工作表", style="Field.TLabel").grid(row=1, column=0, sticky="w", pady=(px(8), 0))
        self.sheet_combo = app.ttk.Combobox(body, textvariable=self.sheet_var, state="readonly", width=12)
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=(0, px(8)), pady=(px(8), 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_changed)
        reload_button = app.ttk.Button(body, text="重新读取", command=lambda: self.reload_workbook_info(force=True))
        reload_button.grid(row=1, column=2, sticky="ew", pady=(px(8), 0))
        app.Tooltip(reload_button, app.TIPS["reload"], self)
        self.sheet_info = app.ttk.Label(body, text="尚未选择文件", style="Muted.TLabel")
        self.sheet_info.grid(row=2, column=1, columnspan=2, sticky="w", pady=(px(5), 0))
        app.tk.Frame(body, height=1, bg=app.PALETTE["border"]).grid(row=3, column=0, columnspan=3, sticky="ew", pady=px(12))
        app.ttk.Label(body, text="输出到", style="Field.TLabel").grid(row=4, column=0, sticky="w")
        self.output_entry = app.ttk.Entry(body, textvariable=self.output_var, width=12)
        self.output_entry.grid(row=4, column=1, sticky="ew", padx=(0, px(8)))
        app.ttk.Button(body, text="选择…", command=self.choose_output_dir).grid(row=4, column=2, sticky="ew")
        app.Tooltip(self.output_entry, app.TIPS["output"], self)
        return card

    def _parameter_frame(self, parent, fields, key_label, key_var, key_tip, on_key_selected):
        """Two fixed-size spinboxes on one line, then the key index with a picker by content."""
        px = self.px
        frame = app.ttk.Frame(parent, style="CardBody.TFrame")
        frame.columnconfigure(6, weight=1)
        column = 0
        for index, (label, variable, unit, tip) in enumerate(fields):
            text = app.ttk.Label(frame, text=label, style="Field.TLabel")
            text.grid(row=0, column=column, sticky="w", padx=(0 if index == 0 else px(16), px(6)))
            spin = app.ttk.Spinbox(frame, from_=0, to=1048576, textvariable=variable, width=5)
            spin.grid(row=0, column=column + 1, sticky="w")
            app.ttk.Label(frame, text=unit, style="Field.TLabel").grid(row=0, column=column + 2, sticky="w", padx=(px(5), 0))
            for widget in (text, spin):
                app.Tooltip(widget, tip, self)
            column += 3
        text = app.ttk.Label(frame, text=key_label, style="Field.TLabel")
        text.grid(row=1, column=0, sticky="w", pady=(px(10), 0), padx=(0, px(6)))
        spin = app.ttk.Spinbox(frame, from_=1, to=1048576, textvariable=key_var, width=5)
        spin.grid(row=1, column=1, sticky="w", pady=(px(10), 0))
        combo = app.ttk.Combobox(frame, state="readonly", width=8)
        combo.grid(row=1, column=2, columnspan=5, sticky="ew", padx=(px(8), 0), pady=(px(10), 0))
        combo.bind("<<ComboboxSelected>>", on_key_selected)
        for widget in (text, spin, combo):
            app.Tooltip(widget, key_tip, self)
        return frame, combo

    def _segmented(self, parent: app.tk.Widget, variable: app.tk.StringVar, options, command) -> app.ttk.Frame:
        frame = app.ttk.Frame(parent, style="CardBody.TFrame")
        frame.columnconfigure(tuple(range(len(options))), weight=1, uniform="segment")
        for column, (value, text) in enumerate(options):
            app.ttk.Radiobutton(
                frame, text=text, value=value, variable=variable, command=command, style="Segment.TRadiobutton",
            ).grid(row=0, column=column, sticky="ew")
        return frame

    def _build_options_card(self, parent: app.tk.Widget) -> app.ttk.Frame:
        px = self.px
        card, _head, body = self._card(parent, "② 拆分设置")
        self._segmented(
            body, self.mode_var, (("column", "按关键列拆分行"), ("row", "按关键行拆分列")), self.on_mode_changed,
        ).grid(row=0, column=0, sticky="ew")
        self.mode_hint = app.ttk.Label(body, style="Muted.TLabel", wraplength=px(300), justify="left")
        self.mode_hint.grid(row=1, column=0, sticky="ew", pady=(px(8), px(12)))
        self.column_options, self.key_column_combo = self._parameter_frame(
            body,
            [("表头", self.header_rows_var, "行", app.TIPS["header_rows"]), ("表尾", self.footer_rows_var, "行", app.TIPS["footer_rows"])],
            "关键列", self.key_column_var, app.TIPS["key_column"], self.on_key_column_selected,
        )
        self.column_options.grid(row=2, column=0, sticky="ew")
        self.row_options, self.key_row_combo = self._parameter_frame(
            body,
            [("左侧固定", self.header_cols_var, "列", app.TIPS["header_cols"]), ("右侧固定", self.footer_cols_var, "列", app.TIPS["footer_cols"])],
            "关键行", self.key_row_var, app.TIPS["key_row"], self.on_key_row_selected,
        )
        self.row_options.grid(row=2, column=0, sticky="ew")
        self.row_options.grid_remove()
        return card

    def _build_tips_card(self, parent: app.tk.Widget) -> app.ttk.Frame:
        card, _head, body = self._card(parent, "说明")
        app.ttk.Label(body, text=app.TIPS["notes"], style="Note.TLabel", justify="left").grid(row=0, column=0, sticky="w")
        return card

    def _scrolled_tree(self, parent: app.tk.Widget, horizontal: bool, **options) -> app.ttk.Treeview:
        frame = app.ttk.Frame(parent, style="Card.TFrame", padding=1)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        tree = app.ttk.Treeview(frame, **options)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = app.ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=scroll_y.set)
        if horizontal:
            scroll_x = app.ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
            scroll_x.grid(row=1, column=0, sticky="ew")
            tree.configure(xscrollcommand=scroll_x.set)
            tree.bind("<Shift-MouseWheel>", lambda event: tree.xview_scroll(-1 if event.delta > 0 else 1, "units"))
            tree.bind("<Shift-Button-4>", lambda _event: tree.xview_scroll(-1, "units"))
            tree.bind("<Shift-Button-5>", lambda _event: tree.xview_scroll(1, "units"))
        return tree

    def _build_preview_card(self, parent: app.tk.Widget) -> app.ttk.Frame:
        px = self.px
        card, head, body = self._card(parent, "表格预览")
        self.legend = app.ttk.Frame(head, style="CardBody.TFrame")
        self.legend.grid(row=0, column=2, sticky="e")
        body.rowconfigure(0, weight=1)
        self.preview_table = self._scrolled_tree(body, True, show="headings", selectmode="none", height=6)
        self.preview_table.bind("<ButtonRelease-1>", self.on_preview_table_click)
        self.preview_table.bind("<Button-3>", self.on_preview_context_menu)
        if app.sys.platform == "darwin":
            self.preview_table.bind("<Button-2>", self.on_preview_context_menu)
            self.preview_table.bind("<Control-Button-1>", self.on_preview_context_menu)
        self.preview_table.tag_configure("header", background=app.PALETTE["row_header"])
        self.preview_table.tag_configure("footer", background=app.PALETTE["row_footer"])
        self.preview_table.tag_configure("keyrow", background=app.PALETTE["row_key"])
        self.preview_table.tag_configure("split", background=app.PALETTE["row_split"])
        self.preview_table.tag_configure("gap", foreground=app.PALETTE["disabled"])
        self.preview_table.tag_configure("odd", background=app.PALETTE["row_stripe"])
        self.preview_message = app.ttk.Label(self.preview_table.master, style="Empty.TLabel", justify="center",
                                         anchor="center", wraplength=px(460))
        self._show_preview_message("尚未选择文件\n\n点击左侧“选择…”或按 Ctrl+O 打开 Excel 文件")
        self.preview_hint = app.ttk.Label(body, style="Muted.TLabel")
        self.preview_hint.grid(row=1, column=0, sticky="w", pady=(px(8), 0))
        return card


"""
PDF·Word·JPG 互转工具 — 图形界面
武汉纺织大学管理学院媒体运营部

基于 tkinter 的 GUI，支持：
  - 批量文件列表（Treeview）
  - 每行独立设置目标格式（嵌入 Combobox）
  - 输出设置面板（质量/DPI/页码范围/冲突策略）
  - 双进度条 + 后台线程转换
  - 启动欢迎窗口
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD

from converter_engine import ConverterEngine

# ── 常量 ──────────────────────────────────────────────────────

APP_TITLE = "PDF·Word·JPG互转工具-武汉纺织大学管理学院媒体运营部"
WELCOME_TITLE = "PDF·Word·JPG互转工具"
WELCOME_SUBTITLE = "武汉纺织大学管理学院媒体运营部"
WELCOME_TEXT = "本软件是自由软件，基于MIT协议开源。"
AUTHOR_GITHUB = "https://github.com/istoryofspring"
SKIP_WELCOME_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".converter_skip_welcome"
)

SUPPORTED_EXTS = [("文档与图片", "*.pdf;*.docx;*.doc;*.jpg;*.jpeg;*.png")]

TARGET_FORMATS = ("PDF", "Word", "JPEG")

# 每行的后台数据：{iid: {"path": str, "source": str, "target": str, "size": int}}
_file_data: dict = {}


# ── 启动欢迎窗口 ──────────────────────────────────────────────

def show_welcome(parent):
    """如果标记文件不存在，弹出欢迎 Toplevel。返回 True 表示用户确认继续。"""
    if os.path.exists(SKIP_WELCOME_FILE):
        return True

    result = {"proceed": False}

    win = tk.Toplevel(parent)
    win.title(WELCOME_TITLE)
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    # 内容框架
    frame = ttk.Frame(win, padding=24)
    frame.pack()

    ttk.Label(frame, text=WELCOME_TITLE, font=("Microsoft YaHei UI", 16, "bold")).pack(pady=(0, 4))
    ttk.Label(frame, text=WELCOME_SUBTITLE, font=("Microsoft YaHei UI", 10)).pack(pady=(0, 12))
    ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=8)
    ttk.Label(frame, text=WELCOME_TEXT, font=("Microsoft YaHei UI", 10)).pack(pady=(4, 0))
    ttk.Label(frame, text="作者GitHub:", font=("Microsoft YaHei UI", 10)).pack(pady=(4, 0))

    link = tk.Label(
        frame,
        text=AUTHOR_GITHUB,
        fg="blue",
        cursor="hand2",
        font=("Microsoft YaHei UI", 10, "underline"),
    )
    link.pack(pady=(0, 12))
    link.bind("<Button-1>", lambda e: os.startfile(AUTHOR_GITHUB))

    skip_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(frame, text="下次不再显示", variable=skip_var).pack(pady=(4, 8))

    def on_click():
        if skip_var.get():
            try:
                with open(SKIP_WELCOME_FILE, "w", encoding="utf-8") as f:
                    f.write("skip")
            except OSError:
                pass
        result["proceed"] = True
        win.destroy()

    ttk.Button(frame, text="开始使用", command=on_click).pack(pady=(4, 0))

    win.update_idletasks()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    pw, ph = parent.winfo_width(), parent.winfo_height()
    px, py = parent.winfo_x(), parent.winfo_y()
    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

    win.wait_window()
    return result["proceed"]


# ── 主窗口 ────────────────────────────────────────────────────

class ConverterApp(TkinterDnD.Tk):
    """主应用程序窗口。"""

    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.minsize(600, 400)

        # 主题
        style = ttk.Style(self)
        available = style.theme_names()
        if "vista" in available:
            style.theme_use("vista")
        elif "clam" in available:
            style.theme_use("clam")

        self.engine = ConverterEngine()
        self._conversion_running = False
        self._cancel_requested = False

        # 第一批"全部→"提示标记
        self._batch_hint_shown = False

        # 行计数器
        self._row_counter = 0

        self._build_ui()
        self._setup_drag_drop()
        self._center_window(800, 620)

        # 启动后弹出欢迎窗口
        self.after(100, lambda: show_welcome(self))

    # ── 布局 ────────────────────────────────────────────────

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=0)  # 工具栏
        self.rowconfigure(1, weight=1)  # 文件列表
        self.rowconfigure(2, weight=0)  # 输出设置
        self.rowconfigure(3, weight=0)  # 进度条
        self.rowconfigure(4, weight=0)  # 底部按钮

        self._build_toolbar()
        self._build_file_list()
        self._build_output_settings()
        self._build_progress_area()
        self._build_bottom_bar()

    def _build_toolbar(self):
        """顶部工具栏。"""
        bar = ttk.Frame(self, padding=4)
        bar.grid(row=0, column=0, sticky="ew")
        bar.columnconfigure(4, weight=1)  # 中间弹簧

        ttk.Button(bar, text="+ 添加文件", command=self._add_files).grid(row=0, column=0, padx=2)
        ttk.Button(bar, text="+ 导入文件夹", command=self._import_folder).grid(row=0, column=1, padx=2)
        ttk.Button(bar, text="— 移除选中", command=self._remove_selected).grid(row=0, column=2, padx=2)
        ttk.Button(bar, text="清空列表", command=self._clear_all).grid(row=0, column=3, padx=2)

        # 弹簧
        ttk.Label(bar, text="").grid(row=0, column=4, sticky="ew", padx=8)

        ttk.Button(bar, text="全部→PDF", command=lambda: self._batch_set_target("PDF")).grid(
            row=0, column=5, padx=2
        )
        ttk.Button(bar, text="全部→Word", command=lambda: self._batch_set_target("Word")).grid(
            row=0, column=6, padx=2
        )
        ttk.Button(bar, text="全部→JPEG", command=lambda: self._batch_set_target("JPEG")).grid(
            row=0, column=7, padx=2
        )

        self._only_selected_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            bar, text="仅选中的行", variable=self._only_selected_var
        ).grid(row=0, column=8, padx=(12, 2))

    def _build_file_list(self):
        """Treeview 文件列表。"""
        list_frame = ttk.Frame(self, padding=4)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        columns = ("#", "filename", "source", "target", "size", "status")
        self._tree = ttk.Treeview(
            list_frame,
            columns=columns,
            show="headings",
            selectmode="extended",
        )

        self._tree.heading("#", text="#")
        self._tree.heading("filename", text="文件名")
        self._tree.heading("source", text="源格式")
        self._tree.heading("target", text="目标格式")
        self._tree.heading("size", text="大小")
        self._tree.heading("status", text="状态")

        self._tree.column("#", width=36, anchor="center", stretch=False)
        self._tree.column("filename", width=220, stretch=True)
        self._tree.column("source", width=64, anchor="center", stretch=False)
        self._tree.column("target", width=72, anchor="center", stretch=False)
        self._tree.column("size", width=72, anchor="center", stretch=False)
        self._tree.column("status", width=90, anchor="center", stretch=False)

        # 滚动条
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # 右键菜单
        self._context_menu = tk.Menu(self, tearoff=0)
        self._context_menu.add_command(label="移除本行", command=self._remove_selected)
        self._context_menu.add_command(label="清空已完成", command=self._clear_completed)
        self._context_menu.add_separator()
        self._context_menu.add_command(label="打开文件位置", command=self._open_file_location)
        self._tree.bind("<Button-3>", self._on_right_click)

        # 双击切换目标格式
        self._tree.bind("<Double-1>", self._on_double_click)

    def _setup_drag_drop(self):
        """注册文件拖放支持。"""
        self.drop_target_register("*")
        self.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drop(self, event):
        """处理拖放文件事件。"""
        # tkinterdnd2 的 event.data 是花括号包裹的路径列表
        raw = event.data
        # 解析花括号包裹的路径（如 '{C:/a.pdf} {D:/b.jpg}'）
        import re
        paths = re.findall(r'\{([^}]+)\}', raw)
        if not paths:
            # 如果没有花括号，尝试直接按空格分割
            paths = raw.split()
        for p in paths:
            p = p.strip()
            if os.path.isfile(p):
                self._add_single_file(p)
            elif os.path.isdir(p):
                # 如果是文件夹，扫描里面所有支持格式的文件
                for name in os.listdir(p):
                    ext = os.path.splitext(name)[1].lower()
                    if ext in ConverterEngine.SUPPORTED_EXT:
                        self._add_single_file(os.path.join(p, name))

    def _build_output_settings(self):
        """输出设置面板。"""
        group = ttk.LabelFrame(self, text="输出设置", padding=6)
        group.grid(row=2, column=0, sticky="ew", padx=4, pady=2)

        # Row 0: 输出目录
        self._same_dir_var = tk.BooleanVar(value=True)
        cb = ttk.Checkbutton(
            group, text="输出到源文件所在目录", variable=self._same_dir_var, command=self._toggle_out_dir
        )
        cb.grid(row=0, column=0, sticky="w", padx=4, pady=2)

        ttk.Label(group, text="或指定目录:").grid(row=0, column=1, padx=(12, 2))
        self._out_dir_var = tk.StringVar()
        self._out_dir_entry = ttk.Entry(group, textvariable=self._out_dir_var, width=28, state="disabled")
        self._out_dir_entry.grid(row=0, column=2, padx=2)
        self._out_dir_btn = ttk.Button(
            group, text="浏览", command=self._browse_out_dir, state="disabled"
        )
        self._out_dir_btn.grid(row=0, column=3, padx=2)

        # 分隔
        ttk.Separator(group, orient="horizontal").grid(
            row=1, column=0, columnspan=4, sticky="ew", pady=8
        )

        # Row 2: JPEG 质量
        ttk.Label(group, text="JPEG 质量 (1-100):").grid(row=2, column=0, sticky="e", padx=4, pady=2)
        self._quality_var = tk.IntVar(value=90)
        quality_spin = ttk.Spinbox(
            group, from_=1, to=100, textvariable=self._quality_var, width=5
        )
        quality_spin.grid(row=2, column=1, sticky="w", padx=2, pady=2)
        quality_scale = ttk.Scale(
            group, from_=1, to=100, variable=self._quality_var, orient="horizontal", length=120
        )
        quality_scale.grid(row=2, column=1, sticky="w", padx=(52, 0), pady=2)

        # DPI
        ttk.Label(group, text="渲染 DPI:").grid(row=2, column=2, sticky="e", padx=(12, 4), pady=2)
        self._dpi_var = tk.IntVar(value=200)
        ttk.Spinbox(group, from_=72, to=600, textvariable=self._dpi_var, width=5).grid(
            row=2, column=3, sticky="w", padx=2, pady=2
        )

        # Row 3: PDF 页面范围
        ttk.Label(group, text="PDF 页面范围:").grid(row=3, column=0, sticky="e", padx=4, pady=2)
        self._page_mode_var = tk.StringVar(value="all")
        page_frame = ttk.Frame(group)
        page_frame.grid(row=3, column=1, columnspan=3, sticky="w", padx=2, pady=2)
        ttk.Radiobutton(page_frame, text="全部页面", variable=self._page_mode_var, value="all").pack(
            side="left", padx=4
        )
        ttk.Radiobutton(page_frame, text="指定范围:", variable=self._page_mode_var, value="range").pack(
            side="left", padx=(12, 4)
        )
        self._page_start_var = tk.StringVar(value="1")
        ttk.Entry(page_frame, textvariable=self._page_start_var, width=4).pack(side="left")
        ttk.Label(page_frame, text="—").pack(side="left", padx=2)
        self._page_end_var = tk.StringVar(value="10")
        ttk.Entry(page_frame, textvariable=self._page_end_var, width=4).pack(side="left")
        ttk.Label(page_frame, text="(仅 PDF→JPEG 生效)").pack(side="left", padx=6)

        # Row 4: 文件冲突
        ttk.Label(group, text="文件冲突策略:").grid(row=4, column=0, sticky="e", padx=4, pady=2)
        self._conflict_var = tk.StringVar(value="rename")
        cf_frame = ttk.Frame(group)
        cf_frame.grid(row=4, column=1, columnspan=3, sticky="w", padx=2, pady=2)
        ttk.Radiobutton(cf_frame, text="自动重命名", variable=self._conflict_var, value="rename").pack(
            side="left", padx=4
        )
        ttk.Radiobutton(cf_frame, text="覆盖", variable=self._conflict_var, value="overwrite").pack(
            side="left", padx=4
        )
        ttk.Radiobutton(cf_frame, text="跳过", variable=self._conflict_var, value="skip").pack(
            side="left", padx=4
        )

    def _build_progress_area(self):
        """双进度条区域。"""
        prog_frame = ttk.Frame(self, padding=4)
        prog_frame.grid(row=3, column=0, sticky="ew", padx=4)
        prog_frame.columnconfigure(0, weight=1)

        # 总进度
        top_row = ttk.Frame(prog_frame)
        top_row.grid(row=0, column=0, sticky="ew")
        top_row.columnconfigure(0, weight=1)
        self._overall_bar = ttk.Progressbar(top_row, mode="determinate", length=400)
        self._overall_bar.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self._overall_label = ttk.Label(top_row, text="0/0 个文件")
        self._overall_label.grid(row=0, column=1)

        # 当前文件详情
        self._detail_label = ttk.Label(prog_frame, text="就绪", foreground="gray")
        self._detail_label.grid(row=1, column=0, sticky="w", pady=(2, 0))

    def _build_bottom_bar(self):
        """底部按钮栏。"""
        bottom = ttk.Frame(self, padding=6)
        bottom.grid(row=4, column=0, sticky="ew")
        bottom.columnconfigure(1, weight=1)

        self._start_btn = ttk.Button(
            bottom, text="▶ 开始全部转换", command=self._start_conversion
        )
        self._start_btn.grid(row=0, column=0, padx=4)

        self._cancel_btn = ttk.Button(
            bottom, text="取消", command=self._cancel_conversion, state="disabled"
        )
        self._cancel_btn.grid(row=0, column=1, padx=4, sticky="w")

        self._open_out_btn = ttk.Button(
            bottom, text="打开输出目录", command=self._open_output_dir
        )
        self._open_out_btn.grid(row=0, column=2, padx=4)

        self._status_label = ttk.Label(bottom, text="状态: 就绪", foreground="gray")
        self._status_label.grid(row=0, column=3, padx=(20, 4))

    # ── 窗口工具 ────────────────────────────────────────────

    def _center_window(self, width, height):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = (screen_w - width) // 2
        y = (screen_h - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ── 工具栏回调 ──────────────────────────────────────────

    def _add_files(self):
        """弹出多文件选择对话框并加入列表。"""
        paths = filedialog.askopenfilenames(
            title="选择文件",
            filetypes=SUPPORTED_EXTS,
        )
        for p in paths:
            self._add_single_file(p)

    def _import_folder(self):
        """扫描文件夹中所有支持格式的文件。"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return
        for name in os.listdir(folder):
            ext = os.path.splitext(name)[1].lower()
            if ext in ConverterEngine.SUPPORTED_EXT:
                full = os.path.join(folder, name)
                self._add_single_file(full)

    def _add_single_file(self, path: str):
        """将单个文件加入 Treeview。去重检查。"""
        # 去重：已存在路径则跳过
        for iid, data in _file_data.items():
            if data["path"] == path:
                return

        try:
            source = ConverterEngine.detect_format(path)
        except ValueError:
            messagebox.showwarning("格式不支持", f"无法识别的文件格式：\n{path}")
            return

        size = os.path.getsize(path)
        size_str = ConverterEngine._format_size(size)

        # 默认目标格式：选第一个非源格式的目标
        available_targets = [f for f in TARGET_FORMATS if f.lower() != source.lower()]
        default_target = available_targets[0] if available_targets else TARGET_FORMATS[0]

        self._row_counter += 1
        iid = str(self._row_counter)

        values = (self._row_counter, os.path.basename(path), source, default_target, size_str, "就绪")
        self._tree.insert("", "end", iid=iid, values=values)

        _file_data[iid] = {
            "path": path,
            "source": source,
            "target": default_target,
            "size": size,
        }

        self._update_status(f"已添加 {len(_file_data)} 个文件")

    def _remove_selected(self):
        """移除 Treeview 中选中的行。"""
        selected = self._tree.selection()
        for iid in selected:
            self._tree.delete(iid)
            _file_data.pop(iid, None)
        self._update_status(f"列表中 {len(_file_data)} 个文件")

    def _clear_all(self):
        """清空全部列表。"""
        if not _file_data:
            return
        if messagebox.askyesno("确认清空", "确定要清空全部文件列表吗？"):
            for iid in list(_file_data.keys()):
                self._tree.delete(iid)
            _file_data.clear()
            self._update_status("列表已清空")

    def _clear_completed(self):
        """清空已完成的行。"""
        to_remove = [
            iid for iid, data in _file_data.items()
            if data.get("status") in ("完成 ✓", "失败 ✗", "已跳过")
        ]
        for iid in to_remove:
            self._tree.delete(iid)
            _file_data.pop(iid, None)
        self._update_status(f"清除了 {len(to_remove)} 条已完成记录")

    def _open_file_location(self):
        """在资源管理器中打开选中文件所在目录。"""
        selected = self._tree.selection()
        if selected:
            data = _file_data.get(selected[0])
            if data:
                folder = os.path.dirname(data["path"])
                os.startfile(folder)

    def _batch_set_target(self, fmt: str):
        """一键设置所有行（或选中行）的目标格式。"""

        # 首次使用时弹出提示
        if not self._batch_hint_shown:
            self._batch_hint_shown = True
            messagebox.showinfo(
                "提示",
                "请单击左下角\"开始全部转换\"按钮开始转换。",
            )

        selected = self._tree.selection()
        target_iids = selected if selected else list(_file_data.keys())

        count = 0
        for iid in target_iids:
            data = _file_data.get(iid)
            if data and data["source"].lower() == fmt.lower():
                continue  # 与源格式相同，跳过
            if data:
                data["target"] = fmt
                self._tree.set(iid, "target", fmt)
                count += 1

        self._update_status(f"已设置 {count} 行的目标格式为 {fmt}")

    def _toggle_out_dir(self):
        """输出目录 Enable/Disable 联动。"""
        if self._same_dir_var.get():
            self._out_dir_entry.configure(state="disabled")
            self._out_dir_btn.configure(state="disabled")
        else:
            self._out_dir_entry.configure(state="readonly")
            self._out_dir_btn.configure(state="normal")

    def _browse_out_dir(self):
        """浏览自定义输出目录。"""
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self._out_dir_var.set(folder)

    # ── Treeview 交互 ───────────────────────────────────────

    def _on_right_click(self, event):
        """右键菜单。"""
        item = self._tree.identify_row(event.y)
        if item:
            self._tree.selection_set(item)
            self._context_menu.post(event.x_root, event.y_root)

    def _on_double_click(self, event):
        """双击目标格式列 → 弹出 Combobox 切换。"""
        column = self._tree.identify_column(event.x)
        if column != "#4":  # 第 4 列 = target
            return
        iid = self._tree.identify_row(event.y)
        if not iid:
            return

        data = _file_data.get(iid)
        if not data:
            return

        # 可选目标格式 = 非源格式
        available = [f for f in TARGET_FORMATS if f.lower() != data["source"].lower()]

        # 获取单元格位置放置 Combobox
        bbox = self._tree.bbox(iid, column="target")
        if not bbox:
            return

        x, y, w, h = bbox

        combo = ttk.Combobox(self._tree, values=available, state="readonly", width=8)
        combo.set(data["target"])
        combo.place(x=x - 2, y=y, width=w + 4, height=h + 4)

        def on_select(event):
            new_val = combo.get()
            if new_val:
                data["target"] = new_val
                self._tree.set(iid, "target", new_val)
            combo.destroy()

        def on_focus_out(event):
            combo.destroy()

        combo.bind("<<ComboboxSelected>>", on_select)
        combo.bind("<FocusOut>", on_focus_out)
        combo.focus_set()

    # ── 状态更新（线程安全）─────────────────────────────────

    def _update_status(self, msg: str, color: str = "gray"):
        self.after(0, lambda: (
            self._status_label.configure(text=f"状态: {msg}", foreground=color)
        ))

    def _update_detail(self, msg: str):
        self.after(0, lambda: self._detail_label.configure(text=msg))

    def _update_overall_progress(self, value: int, text: str):
        self.after(0, lambda: (
            self._overall_bar.configure(value=value),
            self._overall_label.configure(text=text),
        ))

    def _update_tree_status(self, iid: str, status: str):
        self.after(0, lambda: self._tree.set(iid, "status", status))

    def _set_buttons_state(self, running: bool):
        self.after(0, lambda: (
            self._start_btn.configure(state="disabled" if running else "normal"),
            self._cancel_btn.configure(state="normal" if running else "disabled"),
        ))

    # ── 转换逻辑 ────────────────────────────────────────────

    def _resolve_output_dir(self, data: dict) -> str:
        """解析单个文件的输出目录。"""
        if self._same_dir_var.get():
            return os.path.dirname(data["path"])
        custom = self._out_dir_var.get().strip()
        if custom:
            return custom
        return os.path.dirname(data["path"])

    def _get_page_range(self):
        """解析页码范围，返回 (start, end) 或 (None, None)。"""
        if self._page_mode_var.get() == "all":
            return None, None
        try:
            s = int(self._page_start_var.get())
            e = int(self._page_end_var.get())
            return s, e
        except ValueError:
            return None, None  # 无效输入视为全部

    def _start_conversion(self):
        """开始转换全部文件（或仅选中的行）。"""
        if not _file_data:
            messagebox.showinfo("提示", "请先添加文件。")
            return

        # 确定要转换的行
        if self._only_selected_var.get():
            selected = self._tree.selection()
            if not selected:
                messagebox.showinfo("提示", "请先选择要转换的行，或取消勾选'仅选中的行'。")
                return
            work_iids = [iid for iid in selected if iid in _file_data]
        else:
            work_iids = list(_file_data.keys())

        if not work_iids:
            return

        # 恢复所有行的状态为"就绪"
        for iid in work_iids:
            _file_data[iid].pop("status", None)
            self._tree.set(iid, "status", "就绪")

        self._conversion_running = True
        self._cancel_requested = False
        self._set_buttons_state(True)

        thread = threading.Thread(target=self._run_conversion, args=(work_iids,), daemon=True)
        thread.start()

    def _cancel_conversion(self):
        """请求取消转换。"""
        self._cancel_requested = True
        self._update_status("正在取消...", "orange")

    def _run_conversion(self, work_iids: list):
        """在后台线程中依次转换。"""
        total = len(work_iids)
        success_count = 0
        fail_count = 0
        skip_count = 0
        failed_items = []

        dpi = self._dpi_var.get()
        quality = self._quality_var.get()
        conflict = self._conflict_var.get()
        start_page, end_page = self._get_page_range()

        self._update_overall_progress(0, f"0/{total} 个文件")

        for idx, iid in enumerate(work_iids):
            if self._cancel_requested:
                self._update_detail("已取消")
                break

            data = _file_data.get(iid)
            if not data:
                continue

            # 检查文件是否存在
            if not os.path.exists(data["path"]):
                self._update_tree_status(iid, "文件缺失")
                fail_count += 1
                failed_items.append((data["path"], "文件不存在"))
                self._update_overall_progress(
                    int((idx + 1) / total * 100), f"{idx + 1}/{total} 个文件"
                )
                continue

            output_dir = self._resolve_output_dir(data)
            os.makedirs(output_dir, exist_ok=True)

            self._update_tree_status(iid, "转换中...")
            self._update_detail(f"{os.path.basename(data['path'])} → {data['target']}")

            # 执行转换（带进度回调）
            result = self.engine.convert(
                input_path=data["path"],
                source_fmt=data["source"],
                target_fmt=data["target"],
                output_dir=output_dir,
                conflict=conflict,
                dpi=dpi,
                quality=quality,
                start_page=start_page,
                end_page=end_page,
                progress_callback=None,  # 可在之后细化
            )

            if result["success"]:
                self._update_tree_status(iid, "完成 ✓")
                # 记录上一次输出目录
                self._last_output_dir = output_dir
                success_count += 1
            else:
                self._update_tree_status(iid, "失败 ✗")
                fail_count += 1
                failed_items.append((data["path"], result.get("error", "未知错误")))

            self._update_overall_progress(
                int((idx + 1) / total * 100), f"{idx + 1}/{total} 个文件"
            )

        # 完成
        self._conversion_running = False
        self._set_buttons_state(False)
        self._update_detail("转换完成")

        # 汇总弹窗
        self.after(100, lambda: self._show_summary(success_count, fail_count, skip_count, failed_items))

    def _show_summary(self, success, fail, skip, failed_items):
        """弹出汇总结果对话框。"""
        win = tk.Toplevel(self)
        win.title("转换完成")
        win.resizable(False, False)
        win.transient(self)
        win.grab_set()

        frame = ttk.Frame(win, padding=20)
        frame.pack()

        ttk.Label(frame, text="转换完成", font=("Microsoft YaHei UI", 14, "bold")).pack(pady=(0, 12))

        ttk.Label(frame, text=f"✅ 成功: {success} 个文件").pack(anchor="w")
        ttk.Label(frame, text=f"❌ 失败: {fail} 个文件").pack(anchor="w")
        ttk.Label(frame, text=f"⏭️ 跳过: {skip} 个文件").pack(anchor="w")

        if failed_items:
            ttk.Separator(frame, orient="horizontal").pack(fill="x", pady=8)
            ttk.Label(frame, text="失败详情:", font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")
            text = tk.Text(frame, width=50, height=min(8, len(failed_items) + 2), wrap="word")
            text.pack(pady=4)
            for path, err in failed_items:
                text.insert("end", f"• {os.path.basename(path)}: {err}\n")
            text.configure(state="disabled")

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(12, 0))

        if hasattr(self, "_last_output_dir") and success > 0:
            ttk.Button(
                btn_frame,
                text="打开输出目录",
                command=lambda: os.startfile(self._last_output_dir),
            ).pack(side="left", padx=4)

        ttk.Button(btn_frame, text="确定", command=win.destroy).pack(side="left", padx=4)

        win.update_idletasks()
        w, h = win.winfo_reqwidth(), win.winfo_reqheight()
        pw, ph = self.winfo_width(), self.winfo_height()
        px, py = self.winfo_x(), self.winfo_y()
        x = px + (pw - w) // 2
        y = py + (ph - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")

    def _open_output_dir(self):
        """打开上一次的输出目录。"""
        if hasattr(self, "_last_output_dir"):
            os.startfile(self._last_output_dir)
        else:
            # 尝试使用第一个文件所在目录
            if _file_data:
                first = next(iter(_file_data.values()))
                os.startfile(os.path.dirname(first["path"]))
            else:
                messagebox.showinfo("提示", "尚无输出目录。")

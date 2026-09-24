# -*- coding: utf-8 -*-
"""
图片表格转 Excel 助手 —— 精致版 GUI（离线）
入口：选择文件 / 选择文件夹 -> 一键转换 -> 合并输出单 sheet xlsx
"""
import os
import queue
import threading
import traceback
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import core

APP_TITLE = "图片表格转 Excel 助手"
APP_VERSION = core.APP_VERSION
PRIMARY = "#2F5597"
PRIMARY_HOVER = "#4472C4"
BG = "#F4F6FA"
CARD = "#FFFFFF"
TEXT = "#1F2937"
MUTED = "#6B7280"
OK_GREEN = "#2E7D32"
ERR_RED = "#C0392B"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x640")
        self.minsize(760, 560)
        self.configure(bg=BG)

        self.files = []          # 待处理文件
        self.out_path = tk.StringVar()
        self.keep_raw = tk.BooleanVar(value=True)   # 是否输出「原文」列
        self.per_file = tk.BooleanVar(value=False)  # 是否每个文件单独 sheet
        self.worker = None
        self.msg_q = queue.Queue()
        self.done_count = 0

        self._build_styles()
        self._build_ui()
        self._set_icon()
        self.after(100, self._poll_queue)

    def _set_icon(self):
        """设置窗口/任务栏图标（打包后从解包目录读取）"""
        try:
            ico = Path(core.app_base_dir()) / "assets" / "app.ico"
            if ico.exists():
                self.iconbitmap(default=str(ico))
        except Exception:
            pass

    # ------------------------------------------------ UI 构建
    def _build_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", background=BG, foreground=TEXT,
                        font=("Microsoft YaHei UI", 10))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"),
                        foreground="#FFFFFF", background=PRIMARY,
                        borderwidth=0, focusthickness=0, padding=(18, 9))
        style.map("Primary.TButton",
                  background=[("active", PRIMARY_HOVER), ("disabled", "#AEB8CC")])
        style.configure("Ghost.TButton", font=("Microsoft YaHei UI", 10),
                        foreground=PRIMARY, background=CARD,
                        borderwidth=1, focusthickness=0, padding=(14, 8))
        style.map("Ghost.TButton",
                  background=[("active", "#E8EEF9")],
                  bordercolor=[("!active", "#C9D3E8"), ("active", PRIMARY)])
        style.configure("Horizontal.TProgressbar", troughcolor="#E3E8F2",
                        background=PRIMARY, borderwidth=0, thickness=10)
        style.configure("Card.TFrame", background=CARD, relief="flat")
        style.configure("Treeview", background=CARD, fieldbackground=CARD,
                        rowheight=30, borderwidth=0, font=("Microsoft YaHei UI", 10))
        style.configure("Treeview.Heading", background="#E8EEF9",
                        foreground=PRIMARY, font=("Microsoft YaHei UI", 10, "bold"),
                        borderwidth=0, padding=(6, 6))
        style.map("Treeview", background=[("selected", "#D6E2F7")])

    def _card(self, parent, **pack):
        frame = tk.Frame(parent, bg=CARD, highlightthickness=1,
                         highlightbackground="#DCE3F0")
        frame.pack(**pack)
        return frame

    def _build_ui(self):
        # 头部横幅
        header = tk.Frame(self, bg=PRIMARY, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=APP_TITLE, bg=PRIMARY, fg="#FFFFFF",
                 font=("Microsoft YaHei UI", 16, "bold")).pack(side="left", padx=20)
        tk.Label(header, text="本地离线识别 · 支持 PDF / JPG / PNG / BMP",
                 bg=PRIMARY, fg="#C9D6F2",
                 font=("Microsoft YaHei UI", 10)).pack(side="right", padx=20)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # 文件区
        card1 = self._card(body)
        card1.pack(fill="x", ipady=8, ipadx=8)
        row1 = tk.Frame(card1, bg=CARD)
        row1.pack(fill="x", padx=8, pady=6)
        ttk.Button(row1, text="选择文件", style="Primary.TButton",
                   command=self.pick_files).pack(side="left")
        ttk.Button(row1, text="选择文件夹", style="Primary.TButton",
                   command=self.pick_folder).pack(side="left", padx=(10, 0))
        ttk.Button(row1, text="移除选中", style="Ghost.TButton",
                   command=self.remove_selected).pack(side="left", padx=(10, 0))
        ttk.Button(row1, text="清空列表", style="Ghost.TButton",
                   command=self.clear_files).pack(side="left", padx=(10, 0))
        self.lbl_count = tk.Label(row1, text="共 0 个文件", bg=CARD, fg=MUTED,
                                  font=("Microsoft YaHei UI", 10))
        self.lbl_count.pack(side="right")

        cols = ("name", "type", "status")
        self.tree = ttk.Treeview(card1, columns=cols, show="headings", height=8)
        for cid, text, w, anchor in (
            ("name", "文件名", 380, "w"),
            ("type", "类型", 90, "center"),
            ("status", "状态", 260, "w"),
        ):
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=anchor)
        self.tree.pack(fill="x", padx=8, pady=(4, 2))

        # 输出与进度
        card2 = self._card(body)
        card2.pack(fill="x", ipady=8, ipadx=8, pady=(12, 0))
        row2 = tk.Frame(card2, bg=CARD)
        row2.pack(fill="x", padx=8, pady=6)
        tk.Label(row2, text="输出位置", bg=CARD, fg=TEXT,
                 font=("Microsoft YaHei UI", 10, "bold")).pack(side="left")
        ttk.Entry(row2, textvariable=self.out_path,
                  font=("Microsoft YaHei UI", 10)).pack(
            side="left", fill="x", expand=True, padx=10)
        ttk.Button(row2, text="浏览…", style="Ghost.TButton",
                   command=self.pick_output).pack(side="left")
        self.btn_go = ttk.Button(row2, text="开始转换", style="Primary.TButton",
                                 command=self.start_convert)
        self.btn_go.pack(side="left", padx=(10, 0))
        self.btn_open = ttk.Button(row2, text="打开输出", style="Ghost.TButton",
                                   command=self.open_output, state="disabled")
        self.btn_open.pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(card2, mode="determinate", maximum=100)
        self.progress.pack(fill="x", padx=8, pady=(4, 0))

        row2b = tk.Frame(card2, bg=CARD)
        row2b.pack(fill="x", padx=8, pady=(6, 0))
        tk.Checkbutton(
            row2b, text="输出「原文」列（便于核对自动修正，如 2.903.00 → 2903.00）",
            variable=self.keep_raw, bg=CARD, fg=MUTED, activebackground=CARD,
            selectcolor=CARD, font=("Microsoft YaHei UI", 9),
            bd=0, highlightthickness=0).pack(side="left")
        tk.Checkbutton(
            row2b, text="每个文件单独工作表（版式不同时推荐）",
            variable=self.per_file, bg=CARD, fg=MUTED, activebackground=CARD,
            selectcolor=CARD, font=("Microsoft YaHei UI", 9),
            bd=0, highlightthickness=0).pack(side="left", padx=(18, 0))

        self.lbl_progress = tk.Label(card2, text="等待任务", bg=CARD, fg=MUTED,
                                     font=("Microsoft YaHei UI", 9))
        self.lbl_progress.pack(anchor="w", padx=8)

        # 日志区
        card3 = self._card(body)
        card3.pack(fill="both", expand=True, ipady=4, pady=(12, 0))
        self.log_text = tk.Text(card3, height=8, bg="#FBFCFE", fg=TEXT,
                                relief="flat", padx=10, pady=8,
                                font=("Microsoft YaHei UI", 9),
                                state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.log_text.tag_config("err", foreground=ERR_RED)
        self.log_text.tag_config("ok", foreground=OK_GREEN)

        # 状态栏
        status = tk.Frame(self, bg="#E8EEF9", height=26)
        status.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(status, text="就绪 · 识别引擎：RapidOCR + SLANet 表格结构（全离线）",
                                   bg="#E8EEF9", fg=MUTED,
                                   font=("Microsoft YaHei UI", 9))
        self.lbl_status.pack(side="left", padx=12)

    # ------------------------------------------------ 文件管理
    def _add_paths(self, paths):
        added = 0
        for p in paths:
            if p in self.files:
                continue
            ext = Path(p).suffix.lower()
            if ext in core.IMG_EXTS | core.PDF_EXTS:
                self.files.append(p)
                kind = "PDF 文档" if ext == ".pdf" else "图片"
                self.tree.insert("", "end", values=(Path(p).name, kind, "待处理"))
                added += 1
        self._refresh_count()
        if added and not self.out_path.get():
            self.out_path.set(core.default_output_path(self.files))

    def _refresh_count(self):
        self.lbl_count.config(text=f"共 {len(self.files)} 个文件")

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="选择图片或 PDF",
            filetypes=[("支持的文件", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp *.pdf"),
                       ("图片", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff *.webp"),
                       ("PDF 文档", "*.pdf"), ("全部文件", "*.*")])
        self._add_paths(list(paths))

    def pick_folder(self):
        folder = filedialog.askdirectory(title="选择包含图片/PDF 的文件夹")
        if folder:
            self._add_paths(core.collect_files(folder))

    def remove_selected(self):
        """删除选中项。按文件名反查索引，不能用 tree.index()——
        多选删除时索引会随删除而位移，导致删错文件。"""
        sel = self.tree.selection()
        if not sel:
            return
        names = {self.tree.item(i, "values")[0] for i in sel}
        # 保持 files 与 tree 顺序一致地重建
        keep_files, removed = [], 0
        for i, item in enumerate(self.tree.get_children()):
            if item in sel:
                removed += 1
                continue
            if i < len(self.files):
                keep_files.append(self.files[i])
        self.files = keep_files
        for item in sel:
            self.tree.delete(item)
        self._refresh_count()
        self._log(f"已移除 {removed} 个文件")

    def clear_files(self):
        self.files.clear()
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._refresh_count()

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel 位置", defaultextension=".xlsx",
            initialfile=f"识别结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if path:
            self.out_path.set(path)

    def open_output(self):
        out = self.out_path.get()
        target = out if os.path.exists(out) else str(Path(out).parent)
        os.startfile(target)  # noqa: S606

    # ------------------------------------------------ 转换
    def start_convert(self):
        if self.worker is not None and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, "正在转换中，请稍候。")
            return
        if not self.files:
            messagebox.showwarning(APP_TITLE, "请先选择要处理的文件或文件夹。")
            return
        out = self.out_path.get().strip()
        if not out.lower().endswith(".xlsx"):
            messagebox.showwarning(APP_TITLE, "输出位置必须是一个 .xlsx 文件。")
            return

        self.btn_go.config(state="disabled")
        self.btn_open.config(state="disabled")
        self.progress["value"] = 0
        self._log(f"开始转换 {len(self.files)} 个文件…")
        self.lbl_status.config(text="正在识别（首次运行需加载模型，请耐心等待）…")
        self.worker = threading.Thread(
            target=self._work, args=(list(self.files), out), daemon=True)
        self.worker.start()

    def _work(self, paths, out):
        try:
            def on_prog(i, n):
                self.msg_q.put(("progress", i, n))

            def on_log(msg):
                self.msg_q.put(("log", msg))

            rows, errors = core.convert_many(paths, progress_cb=on_prog, log_cb=on_log)
            if rows:
                core.write_excel(rows, out, paths,
                                 keep_raw=self.keep_raw.get(),
                                 per_file=self.per_file.get())
                self.msg_q.put(("done", out, len(rows), errors))
            else:
                self.msg_q.put(("fail", "未识别出任何表格内容。" + ("；".join(errors) if errors else "")))
        except Exception:
            self.msg_q.put(("fail", traceback.format_exc()))

    # ------------------------------------------------ 消息泵
    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    i, n = msg[1], msg[2]
                    self.progress["value"] = i * 100.0 / max(n, 1)
                    self.lbl_progress.config(text=f"进度：{i}/{n}")
                elif kind == "log":
                    self._log(msg[1])
                elif kind == "done":
                    out, nrows, errors = msg[1], msg[2], msg[3]
                    self._refresh_tree_done()
                    self.progress["value"] = 100
                    self.lbl_progress.config(text="转换完成")
                    self.lbl_status.config(text=f"完成 · 输出：{out}")
                    self.btn_go.config(state="normal")
                    self.btn_open.config(state="normal")
                    self._log(f"转换完成，共写入 {nrows} 行，输出：{out}", "ok")
                    for e in errors:
                        self._log(f"部分文件处理异常：{e}", "err")
                    messagebox.showinfo(APP_TITLE,
                                        f"转换完成。\n共写入 {nrows} 行数据。\n"
                                        f"输出文件：{out}"
                                        + (f"\n\n有 {len(errors)} 个文件处理异常，详见日志。" if errors else ""))
                elif kind == "fail":
                    self.progress["value"] = 0
                    self.lbl_progress.config(text="转换失败")
                    self.lbl_status.config(text="失败，详见日志")
                    self.btn_go.config(state="normal")
                    self._log(msg[1], "err")
                    messagebox.showerror(APP_TITLE, "转换失败，详见日志窗口。")
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _refresh_tree_done(self):
        for item in self.tree.get_children():
            vals = list(self.tree.item(item, "values"))
            vals[2] = "已完成"
            self.tree.item(item, values=vals)

    def _log(self, text, tag=None):
        self.log_text.config(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{ts}] {text}\n", tag or ())
        self.log_text.see("end")
        self.log_text.config(state="disabled")


def selftest(argv):
    """无界面自检：--selftest <输入路径> <输出xlsx> [--no-raw]"""
    import sys
    keep_raw = "--no-raw" not in argv
    argv = [a for a in argv if a != "--no-raw"]
    log_path = str(Path(argv[1]).with_suffix(".selftest.log"))
    try:
        with open(log_path, "w", encoding="utf-8") as lf:
            sys.stdout = lf
            sys.stderr = lf
            rc = _selftest_impl(argv, keep_raw)
        return rc
    except SystemExit:
        raise
    except Exception:
        with open(log_path, "a", encoding="utf-8") as lf:
            lf.write(traceback.format_exc())
        return 2


def _selftest_impl(argv, keep_raw=True):
    import sys
    target, out = argv[0], argv[1]
    paths = core.collect_files(target) if Path(target).is_dir() else [target]
    print(f"待处理 {len(paths)} 个文件")
    t0 = datetime.now()
    rows, errors = core.convert_many(paths)
    print(f"识别 {len(rows)} 行，异常 {len(errors)} 项")
    for e in errors:
        print("ERR:", e)
    if rows:
        core.write_excel(rows, out, paths, keep_raw=keep_raw)
        print("输出：", out)
    print("耗时：%.1f 秒" % (datetime.now() - t0).total_seconds())
    return 0 if rows else 1


def main():
    import sys
    args = sys.argv[1:]
    if args and args[0] == "--version":
        print(f"{APP_TITLE} v{core.APP_VERSION}")
        return
    if args and args[0] == "--selftest":
        raise SystemExit(selftest(args[1:]))
    App().mainloop()


if __name__ == "__main__":
    main()

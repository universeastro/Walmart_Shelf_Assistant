import json
import os
import subprocess
import sys
import threading
import tempfile
from datetime import datetime
import ctypes
from ctypes import wintypes
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
MAPPER = APP_DIR / "excel_mapper.ps1"
SETTINGS_PATH = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".config") / "WalmartShelfAssistant" / "settings.json"


class CompletionDialog(tk.Toplevel):
    def __init__(self, parent, message):
        super().__init__(parent)
        self.withdraw()
        self.title("处理完成")
        self.transient(parent)
        self.resizable(False, False)
        self.parent = parent
        self._pending = None
        body = ttk.Frame(self, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=message, wraplength=320).pack(pady=(0, 22))
        self.confirm = ttk.Button(body, text="确定", command=self.destroy)
        self.confirm.pack(anchor="e")
        self.bind("<Return>", lambda event: self.destroy())
        self.bind("<Escape>", lambda event: self.destroy())
        self._parent_binding = parent.bind("<Configure>", self._schedule_center, add="+")
        self.bind("<Configure>", self._schedule_center)
        self.update_idletasks()
        self._center()
        self.deiconify()
        self.confirm.focus_set()

    def _schedule_center(self, event):
        if event.widget in (self.parent, self) and self._pending is None:
            self._pending = self.after_idle(self._center)

    def _center(self):
        self._pending = None
        if sys.platform == "win32":
            # Tk sizes describe client areas; Windows rectangles include the
            # title bar and borders, which must participate in centering.
            user32 = ctypes.WinDLL("user32", use_last_error=True)
            user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
            user32.GetAncestor.restype = wintypes.HWND
            user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
            user32.SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int,
                                          ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.UINT]
            owner = user32.GetAncestor(self.parent.winfo_id(), 2)
            dialog = user32.GetAncestor(self.winfo_id(), 2)
            parent_rect, dialog_rect = wintypes.RECT(), wintypes.RECT()
            if user32.GetWindowRect(owner, ctypes.byref(parent_rect)) and user32.GetWindowRect(dialog, ctypes.byref(dialog_rect)):
                x = (parent_rect.left + parent_rect.right - dialog_rect.right + dialog_rect.left) // 2
                y = (parent_rect.top + parent_rect.bottom - dialog_rect.bottom + dialog_rect.top) // 2
                if (dialog_rect.left, dialog_rect.top) != (x, y):
                    user32.SetWindowPos(dialog, None, x, y, 0, 0, 0x0015)
                return
        x = self.parent.winfo_rootx() + (self.parent.winfo_width() - self.winfo_width()) // 2
        y = self.parent.winfo_rooty() + (self.parent.winfo_height() - self.winfo_height()) // 2
        if (self.winfo_x(), self.winfo_y()) != (x, y):
            self.geometry(f"+{x}+{y}")

    def destroy(self):
        if self._pending is not None:
            self.after_cancel(self._pending)
            self._pending = None
        self.parent.unbind("<Configure>", self._parent_binding)
        super().destroy()


class ShelfAssistant(tk.Tk):
    def __init__(self, settings_path=None):
        super().__init__()
        self.title("沃尔玛上架助手")
        self.geometry("880x680")
        self.minsize(720, 620)
        self.configure(bg="#f6f7f9")
        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.row_mode_var = tk.StringVar(value="Visible")
        self.status_var = tk.StringVar(value="请选择文件 A 和文件 B")
        self.running = False
        self.completion_dialog = None
        self.last_output = None
        self.summary_var = tk.StringVar(value="尚无处理结果")
        self._path_tip = None
        self._tip_job = None
        self.settings_path = Path(settings_path) if settings_path is not None else SETTINGS_PATH
        self._paths_dirty = False
        self._path_vars = {
            "source": self.source_var,
            "target": self.target_var,
            "output": self.output_var,
        }
        self._load_paths()
        for variable in self._path_vars.values():
            variable.trace_add("write", self._mark_paths_dirty)
        self._build_ui()
        self._refresh_ready_status()

    def _load_paths(self):
        try:
            settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        if isinstance(settings, dict):
            for key, variable in self._path_vars.items():
                value = settings.get(key)
                if isinstance(value, str):
                    if key in ("source", "target") and value:
                        try:
                            exists = Path(value.strip()).is_file()
                        except (OSError, ValueError):
                            exists = False
                        if not exists:
                            value = ""
                            self._paths_dirty = True
                    variable.set(value)

    def _mark_paths_dirty(self, *_):
        self._paths_dirty = True
        if hasattr(self, "status_label") and not self.running:
            self._refresh_ready_status()

    def _refresh_ready_status(self):
        ready = all(variable.get().strip() for variable in self._path_vars.values())
        self.status_var.set("就绪" if ready else "请选择文件 A、文件 B 和输出位置")
        self.status_label.configure(style="Hint.TLabel")

    def _save_paths(self):
        if not self._paths_dirty:
            return
        temporary = None
        try:
            self.settings_path.parent.mkdir(parents=True, exist_ok=True)
            settings = {key: variable.get() for key, variable in self._path_vars.items()}
            # Replace only after the complete JSON has been written successfully.
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.settings_path.parent,
                                             suffix=".tmp", delete=False) as stream:
                temporary = Path(stream.name)
                json.dump(settings, stream, ensure_ascii=False, indent=2)
            os.replace(temporary, self.settings_path)
            self._paths_dirty = False
        except OSError as exc:
            messagebox.showwarning("路径保存失败", f"无法保存上次使用的路径：{exc}", parent=self)
        finally:
            if temporary is not None:
                try:
                    temporary.unlink(missing_ok=True)
                except OSError:
                    pass

    def destroy(self):
        self._hide_path_tip()
        self._save_paths()
        super().destroy()

    def _build_ui(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        font = ("Microsoft YaHei UI", 10)
        self.option_add("*Font", font)
        style.configure(".", font=font, background="#f6f7f9", foreground="#252b32")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 22, "bold"))
        style.configure("Section.TLabel", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Hint.TLabel", foreground="#66717e")
        style.configure("TEntry", padding=9, fieldbackground="#ffffff", bordercolor="#cbd2dc",
                        lightcolor="#ffffff", darkcolor="#ffffff")
        style.map("TEntry", bordercolor=[("focus", "#0071ce")],
                  fieldbackground=[("disabled", "#edf0f3")])
        style.configure("TButton", padding=(14, 8), background="#ffffff", bordercolor="#cbd2dc")
        style.map("TButton", background=[("active", "#e8f2fc"), ("disabled", "#edf0f3")])
        style.configure("Run.TButton", font=("Microsoft YaHei UI", 11, "bold"),
                        padding=(24, 10), background="#0071ce", foreground="#ffffff", borderwidth=0)
        style.map("Run.TButton", background=[("disabled", "#dce2e9"), ("pressed", "#004f99"), ("active", "#005eae")],
                  foreground=[("disabled", "#7a8592")])
        style.configure("TRadiobutton", padding=(0, 5))
        style.map("TRadiobutton", background=[("active", "#f6f7f9")])
        style.configure("Horizontal.TProgressbar", background="#0071ce", troughcolor="#e4e9ef",
                        borderwidth=0, lightcolor="#0071ce", darkcolor="#0071ce")
        style.configure("Success.TLabel", foreground="#197348")
        style.configure("Error.TLabel", foreground="#b63838")

        outer = ttk.Frame(self, padding=(28, 22))
        outer.pack(fill="both", expand=True)
        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 18))
        ttk.Label(header, text="沃尔玛上架助手", style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="商品上架 / Excel", style="Hint.TLabel").pack(side="right", anchor="s", pady=6)
        ttk.Separator(outer).pack(fill="x", pady=(0, 18))
        ttk.Label(outer, text="文件", style="Section.TLabel").pack(anchor="w", pady=(0, 4))
        panel = ttk.Frame(outer)
        panel.pack(fill="x")
        panel.columnconfigure(1, weight=1)
        self.file_controls = []
        self._file_row(panel, 0, "文件 A（源数据）", self.source_var, self._choose_source)
        self._file_row(panel, 1, "文件 B（沃尔玛模板）", self.target_var, self._choose_target)
        self._file_row(panel, 2, "输出文件", self.output_var, self._choose_output)

        ttk.Label(outer, text="导出范围", style="Section.TLabel").pack(anchor="w", pady=(18, 4))
        options = ttk.Frame(outer)
        options.pack(fill="x")
        self.row_mode_buttons = []
        for value, label in (
            ("Visible", "仅可见行（跳过隐藏行）"),
            ("All", "全部数据行（包含隐藏行）"),
        ):
            button = ttk.Radiobutton(options, text=label, variable=self.row_mode_var, value=value)
            button.pack(side="left", padx=(0, 24))
            self.row_mode_buttons.append(button)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(18, 0))
        self.run_button = ttk.Button(bottom, text="开始填充", style="Run.TButton", command=self.run_mapping)
        self.run_button.pack(side="right")
        self.status_label = ttk.Label(bottom, textvariable=self.status_var, style="Hint.TLabel")
        self.status_label.pack(side="left")
        self.progress = ttk.Progressbar(outer, mode="indeterminate", maximum=100)
        self.progress.pack(fill="x", pady=(10, 10))
        log_toolbar = ttk.Frame(outer)
        log_toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(log_toolbar, text="处理记录", style="Section.TLabel").pack(side="left")
        self.open_output_button = ttk.Button(log_toolbar, text="打开输出位置", command=self._open_output_location,
                                             state="disabled")
        self.open_output_button.pack(side="right")
        ttk.Label(outer, textvariable=self.summary_var, style="Hint.TLabel").pack(anchor="w", pady=(0, 8))
        log_frame = ttk.Frame(outer)
        log_frame.pack(fill="both", expand=True)
        self.log = tk.Text(log_frame, height=5, width=1, state="disabled", bg="#ffffff", fg="#394552",
                           font=("Microsoft YaHei UI", 10), relief="flat", wrap="word",
                           padx=12, pady=10, spacing1=3, spacing3=3,
                           highlightthickness=1, highlightbackground="#dce2e9")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)

    def _file_row(self, parent, row, label, variable, command):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 18), pady=6)
        entry = ttk.Entry(parent, textvariable=variable, width=12)
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        entry.bind("<Enter>", lambda event: self._schedule_path_tip(entry, variable))
        for sequence in ("<Leave>", "<ButtonPress>", "<KeyPress>", "<FocusOut>"):
            entry.bind(sequence, lambda event: self._hide_path_tip(), add="+")
        entry.xview_moveto(1)
        button = ttk.Button(parent, text="浏览...", command=command)
        button.grid(row=row, column=2, padx=(10, 0), pady=6)
        self.file_controls.extend((entry, button))

    def _schedule_path_tip(self, entry, variable):
        self._hide_path_tip()
        def show():
            self._tip_job = None
            if not variable.get():
                return
            tip = self._path_tip = tk.Toplevel(self)
            tip.withdraw()
            tip.overrideredirect(True)
            ttk.Label(tip, text=variable.get(), padding=10,
                      wraplength=min(560, self.winfo_width() - 40)).pack()
            tip.update_idletasks()
            x = min(entry.winfo_rootx(), self.winfo_rootx() + self.winfo_width() - tip.winfo_reqwidth() - 10)
            tip.geometry(f"+{max(0, x)}+{entry.winfo_rooty() + entry.winfo_height() + 4}")
            tip.deiconify()
        self._tip_job = self.after(450, show)

    def _hide_path_tip(self):
        if self._tip_job is not None:
            self.after_cancel(self._tip_job)
            self._tip_job = None
        if self._path_tip is not None:
            self._path_tip.destroy()
            self._path_tip = None

    def _browse_options(self, variable, save=False):
        options = {"parent": self}
        for value in (variable.get(), self.target_var.get(), self.source_var.get()):
            if not value.strip():
                continue
            try:
                path = Path(value.strip()).absolute()
                directory = path if path.is_dir() else path.parent
                for candidate in (directory, *directory.parents):
                    if candidate.is_dir():
                        options["initialdir"] = str(candidate)
                        break
                if "initialdir" in options:
                    break
            except (OSError, ValueError):
                continue
        if save and variable.get().strip():
            options["initialfile"] = Path(variable.get().strip()).name
        return options

    def _open_output_location(self):
        if self.running or self.last_output is None:
            return
        try:
            if not self.last_output.is_file():
                messagebox.showwarning("输出文件不存在", "输出文件已移动或删除。", parent=self)
                return
            os.startfile(str(self.last_output.parent))
        except OSError as exc:
            messagebox.showerror("无法打开输出位置", str(exc), parent=self)

    def _choose_source(self):
        path = filedialog.askopenfilename(
            title="选择文件 A",
            filetypes=[("Excel 文件", "*.xls;*.xlsx;*.xlsm"), ("所有文件", "*.*")],
            **self._browse_options(self.source_var),
        )
        if path:
            self.source_var.set(path)
            self._set_default_output()

    def _choose_target(self):
        path = filedialog.askopenfilename(
            title="选择文件 B",
            filetypes=[("Excel 文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")],
            **self._browse_options(self.target_var),
        )
        if path:
            self.target_var.set(path)
            self._set_default_output()

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
            **self._browse_options(self.output_var, save=True),
        )
        if path:
            self.output_var.set(path)

    def _set_default_output(self):
        target = self.target_var.get().strip()
        if target and not self.output_var.get().strip():
            target_path = Path(target)
            self.output_var.set(str(target_path.with_name(target_path.stem + "_已填充.xlsx")))

    def _write_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", f"{datetime.now():%H:%M:%S}  {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def run_mapping(self):
        if self.running:
            return
        self._save_paths()
        source = Path(self.source_var.get().strip())
        target = Path(self.target_var.get().strip())
        output = Path(self.output_var.get().strip())
        if not source.is_file() or not target.is_file():
            messagebox.showwarning("缺少文件", "请选择有效的文件 A 和文件 B。")
            return
        if not self.output_var.get().strip():
            messagebox.showwarning("缺少输出路径", "请选择输出文件路径。")
            return
        if output.resolve() in (target.resolve(), source.resolve()):
            messagebox.showwarning("输出路径无效", "输出文件不能覆盖文件 A 或文件 B。")
            return
        self.status_var.set("正在处理...")
        self._hide_path_tip()
        self.last_output = None
        self.open_output_button.configure(state="disabled")
        self.summary_var.set("正在生成输出文件...")
        self.status_label.configure(style="Hint.TLabel")
        self.progress.start(12)
        self.running = True
        self.run_button.configure(state="disabled", text="正在填充...")
        for control in self.file_controls:
            control.configure(state="disabled")
        row_mode = self.row_mode_var.get()
        for button in self.row_mode_buttons:
            button.configure(state="disabled")
        self._write_log(f"开始处理：{source.name} -> {target.name}")
        self._write_log("导出范围：" + ("仅可见行" if row_mode == "Visible" else "全部数据行（包含隐藏行）"))
        threading.Thread(target=self._run_worker, args=(source, target, output, row_mode), daemon=True).start()

    def _run_worker(self, source, target, output, row_mode="Visible"):
        command = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(MAPPER),
            "-SourcePath",
            str(source),
            "-TargetPath",
            str(target),
            "-OutputPath",
            str(output),
            "-RowMode",
            row_mode,
        ]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace")
            raw = completed.stdout.strip().splitlines()
            payload = json.loads(raw[-1]) if raw else {"success": False, "message": completed.stderr.strip()}
            self.after(0, self._finish, payload, completed.returncode)
        except Exception as exc:  # pragma: no cover - UI error path
            self.after(0, self._finish, {"success": False, "message": str(exc)}, 1)

    def _finish(self, payload, returncode):
        self.running = False
        self.progress.stop()
        self.progress.configure(value=0)
        self.run_button.configure(state="normal", text="开始填充")
        for control in self.file_controls:
            control.configure(state="normal")
        for button in self.row_mode_buttons:
            button.configure(state="normal")
        if payload.get("success") and returncode == 0:
            self.summary_var.set(f"读取 {payload.get('rowsRead', 0)} 行    写入 {payload.get('rowsWritten', 0)} 行    跳过隐藏 {payload.get('rowsHiddenSkipped', 0)} 行")
            output_path = payload.get("output") or self.output_var.get()
            if output_path:
                self.last_output = Path(output_path).resolve()
                self.open_output_button.configure(state="normal")
            self.status_var.set("处理完成")
            self.status_label.configure(style="Success.TLabel")
            self._write_log(payload.get("message", "映射完成。"))
            self._write_log(f"输出文件：{payload.get('output', self.output_var.get())}")
            self._write_log(f"读取 {payload.get('rowsRead', 0)} 行，写入 {payload.get('rowsWritten', 0)} 行。")
            self._write_log(f"跳过隐藏行：{payload.get('rowsHiddenSkipped', 0)} 行。")
            for item in payload.get("skipped", []):
                self._write_log("跳过：" + item)
            if self.completion_dialog is not None and self.completion_dialog.winfo_exists():
                self.completion_dialog.destroy()
            self.completion_dialog = CompletionDialog(self, payload.get("message", "映射完成。"))
        else:
            self.summary_var.set("本次处理失败，详情见处理记录")
            self.last_output = None
            self.open_output_button.configure(state="disabled")
            self.status_var.set("处理失败")
            self.status_label.configure(style="Error.TLabel")
            message = payload.get("message") or "Excel 处理失败，请确认本机已安装 Microsoft Excel。"
            if payload.get("stage"):
                self._write_log("失败阶段：" + payload["stage"])
            if payload.get("errorLine"):
                self._write_log(f"脚本行号：{payload['errorLine']}；类型：{payload.get('errorType', '')}")
            self._write_log("错误：" + message)
            messagebox.showerror("处理失败", message)


if __name__ == "__main__":
    app = ShelfAssistant()
    app.mainloop()

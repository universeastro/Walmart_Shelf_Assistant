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

try:
    import windnd
except ImportError:
    windnd = None


APP_DIR = Path(__file__).resolve().parent
MAPPER = APP_DIR / "excel_mapper.ps1"
SETTINGS_PATH = Path(os.environ.get("LOCALAPPDATA") or Path.home() / ".config") / "WalmartShelfAssistant" / "settings.json"
PROFILE_LABELS = {"主线 01": "Mainline", "支线 02": "Req02"}
PROFILE_KEYS = {value: label for label, value in PROFILE_LABELS.items()}
PATH_KEYS = ("source", "target", "output")
ROW_MODES = ("Visible", "All")


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


class OutputModeDialog(tk.Toplevel):
    def __init__(self, parent, output_name):
        super().__init__(parent)
        self.withdraw()
        self.title("输出文件已存在")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.bind("<Escape>", lambda _event: self.destroy())

        body = ttk.Frame(self, padding=24)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=f"{output_name} 已存在。请选择处理方式：", wraplength=440).pack(anchor="w")
        ttk.Label(body, text="替换文件：以文件 B 重新生成输出。", style="Hint.TLabel").pack(anchor="w", pady=(12, 2))
        ttk.Label(body, text="继续写入：保留现有内容，在末尾留 3 行后追加。", style="Hint.TLabel").pack(anchor="w")

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(22, 0))
        ttk.Button(actions, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(actions, text="继续写入", command=lambda: self._select("AppendExisting")).pack(side="right", padx=8)
        replace = ttk.Button(actions, text="替换文件", command=lambda: self._select("Replace"))
        replace.pack(side="right")

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.deiconify()
        self.grab_set()
        replace.focus_set()

    def _select(self, result):
        self.result = result
        self.destroy()

    @classmethod
    def ask(cls, parent, output_name):
        dialog = cls(parent, output_name)
        parent.wait_window(dialog)
        return dialog.result


class ShelfAssistant(tk.Tk):
    def __init__(self, settings_path=None):
        super().__init__()
        self.title("沃尔玛上架助手")
        self.geometry("880x680")
        self.minsize(720, 620)
        self.configure(bg="#f6f7f9")
        self.profile_var = tk.StringVar(value="主线 01")
        self._profile_path_vars = {
            label: {key: tk.StringVar() for key in PATH_KEYS}
            for label in PROFILE_LABELS
        }
        self._profile_row_mode_vars = {
            label: tk.StringVar(value="Visible")
            for label in PROFILE_LABELS
        }
        self._profile_output_labels = {
            label: tk.StringVar(value="输出文件")
            for label in PROFILE_LABELS
        }
        self._active_profile_label = "主线 01"
        self._profile_change_guard = False
        self._activate_profile_vars(self._active_profile_label)
        self.status_var = tk.StringVar(value="请选择文件 A 和文件 B")
        self.running = False
        self.completion_dialog = None
        self.last_output = None
        self._profile_last_outputs = {label: None for label in PROFILE_LABELS}
        self.summary_var = tk.StringVar(value="尚无处理结果")
        self.settings_path = Path(settings_path) if settings_path is not None else SETTINGS_PATH
        self._paths_dirty = False
        self._load_paths()
        self._activate_profile_vars(self.profile_var.get())
        for label, path_vars in self._profile_path_vars.items():
            for variable in path_vars.values():
                variable.trace_add("write", lambda *_args, profile_label=label: self._mark_paths_dirty(profile_label))
            self._profile_row_mode_vars[label].trace_add(
                "write", lambda *_args, profile_label=label: self._mark_paths_dirty(profile_label)
            )
        self._build_ui()
        self.profile_var.trace_add("write", self._on_profile_var_changed)
        self._refresh_ready_status()

    def _activate_profile_vars(self, label):
        if label not in self._profile_path_vars:
            label = "主线 01"
        self._active_profile_label = label
        if self.profile_var.get() != label:
            self.profile_var.set(label)
        self._path_vars = self._profile_path_vars[label]
        self.source_var = self._path_vars["source"]
        self.target_var = self._path_vars["target"]
        self.output_var = self._path_vars["output"]
        self.row_mode_var = self._profile_row_mode_vars[label]
        self.output_label_var = self._profile_output_labels[label]

    def _validated_path_value(self, key, value):
        if key in ("source", "target") and value:
            try:
                exists = Path(value.strip()).is_file()
            except (OSError, ValueError):
                exists = False
            if not exists:
                self._paths_dirty = True
                return ""
        return value

    def _load_paths(self):
        try:
            settings = json.loads(self.settings_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        if not isinstance(settings, dict):
            return

        profiles = settings.get("profiles")
        if isinstance(profiles, dict):
            for profile_key, label in PROFILE_KEYS.items():
                values = profiles.get(profile_key)
                if not isinstance(values, dict):
                    continue
                for key, variable in self._profile_path_vars[label].items():
                    value = values.get(key)
                    if isinstance(value, str):
                        variable.set(self._validated_path_value(key, value))
                row_mode = values.get("row_mode")
                if isinstance(row_mode, str) and row_mode in ROW_MODES:
                    self._profile_row_mode_vars[label].set(row_mode)
            last_profile = settings.get("last_profile")
            if isinstance(last_profile, str) and last_profile in PROFILE_KEYS:
                self.profile_var.set(PROFILE_KEYS[last_profile])
            return

        legacy_target = settings.get("target")
        legacy_label = self._profile_label_for_target(legacy_target) if isinstance(legacy_target, str) else None
        legacy_label = legacy_label or "主线 01"
        for key, variable in self._profile_path_vars[legacy_label].items():
            value = settings.get(key)
            if isinstance(value, str):
                variable.set(self._validated_path_value(key, value))
        legacy_row_mode = settings.get("row_mode")
        if isinstance(legacy_row_mode, str) and legacy_row_mode in ROW_MODES:
            self._profile_row_mode_vars[legacy_label].set(legacy_row_mode)
        if any(key in settings for key in PATH_KEYS):
            self.profile_var.set(legacy_label)
            self._paths_dirty = True

    def _mark_paths_dirty(self, profile_label=None):
        self._paths_dirty = True
        if (profile_label is None or profile_label == self._active_profile_label) and hasattr(self, "status_label") and not self.running:
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
            settings = {
                "version": 2,
                "last_profile": PROFILE_LABELS[self._active_profile_label],
                "profiles": {
                    profile_key: {
                        **{
                            key: variable.get()
                            for key, variable in self._profile_path_vars[label].items()
                        },
                        "row_mode": self._profile_row_mode_vars[label].get(),
                    }
                    for label, profile_key in PROFILE_LABELS.items()
                },
            }
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
        style.configure("TEntry", padding=9, foreground="#252b32", fieldbackground="#ffffff", bordercolor="#cbd2dc",
                        lightcolor="#ffffff", darkcolor="#ffffff")
        style.map("TEntry", bordercolor=[("focus", "#0071ce")],
                  foreground=[("disabled", "#252b32"), ("!disabled", "#252b32")],
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
        self.profile_combo = ttk.Combobox(header, textvariable=self.profile_var,
                                          values=tuple(PROFILE_LABELS), state="readonly", width=12)
        self.profile_combo.pack(side="right", padx=(0, 18), pady=4)
        self.profile_combo.bind("<<ComboboxSelected>>", self._on_profile_selected)
        ttk.Label(header, text="方案", style="Hint.TLabel").pack(side="right", padx=(0, 6), pady=6)
        ttk.Separator(outer).pack(fill="x", pady=(0, 18))
        self.pages = ttk.Frame(outer)
        self.pages.pack(fill="x")
        self.pages.columnconfigure(0, weight=1)
        self._profile_pages = {}
        self._profile_file_controls = {}
        for label in PROFILE_LABELS:
            page = ttk.Frame(self.pages)
            page.grid(row=0, column=0, sticky="nsew")
            page.columnconfigure(0, weight=1)
            ttk.Label(page, text=f"{label} 文件", style="Section.TLabel").grid(
                row=0, column=0, sticky="w", pady=(0, 4)
            )
            panel = ttk.Frame(page)
            panel.grid(row=1, column=0, sticky="ew")
            panel.columnconfigure(1, weight=1)
            controls = []
            path_vars = self._profile_path_vars[label]
            self._file_row(panel, 0, "文件 A（源数据）", path_vars["source"], self._choose_source, "source", label, controls)
            self._file_row(panel, 1, "文件 B（目标文件）", path_vars["target"], self._choose_target, "target", label, controls)
            self._file_row(panel, 2, self._profile_output_labels[label], path_vars["output"], self._choose_output,
                           "output", label, controls)
            self._profile_pages[label] = page
            self._profile_file_controls[label] = controls
        self.all_file_controls = [
            control
            for controls in self._profile_file_controls.values()
            for control in controls
        ]
        self.file_controls = []
        self._show_profile_page(self._active_profile_label)

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

    def _file_row(self, parent, row, label, variable, command, kind, profile_label, controls):
        label_options = {"textvariable": label} if isinstance(label, tk.Variable) else {"text": label}
        ttk.Label(parent, **label_options).grid(row=row, column=0, sticky="w", padx=(0, 18), pady=6)
        entry = ttk.Entry(parent, textvariable=variable, width=12,
                          font=("Microsoft YaHei UI", 10), foreground="#252b32")
        entry.grid(row=row, column=1, sticky="ew", pady=6)
        entry.xview_moveto(1)
        button = ttk.Button(parent, text="浏览...", command=command)
        button.grid(row=row, column=2, padx=(10, 0), pady=6)
        if windnd is not None:
            try:
                windnd.hook_dropfiles(
                    entry,
                    func=lambda paths, v=variable, k=kind, p=profile_label: self.after(
                        0, self._handle_drop_paths, paths, v, k, p
                    ),
                    force_unicode=True,
                )
            # windnd uses a process-wide hook counter and raises TypeError
            # (from an invalid string raise) after its internal limit. Drag
            # and drop is optional, so a failed hook must not block the UI.
            except (OSError, RuntimeError, TypeError):
                pass
        controls.extend((entry, button))

    def _show_profile_page(self, label):
        if self.running and label != self._active_profile_label:
            return False
        self._activate_profile_vars(label)
        if hasattr(self, "row_mode_buttons"):
            for button in self.row_mode_buttons:
                button.configure(variable=self.row_mode_var)
        if hasattr(self, "_profile_pages"):
            self._profile_pages[label].tkraise()
            self.file_controls = self._profile_file_controls[label]
        if hasattr(self, "open_output_button"):
            self.last_output = self._profile_last_outputs[label]
            state = "normal" if self.last_output is not None else "disabled"
            self.open_output_button.configure(state=state)
        if hasattr(self, "status_label") and not self.running:
            self._refresh_ready_status()
        return True

    def _on_profile_selected(self, _event=None):
        self._on_profile_var_changed()

    def _on_profile_var_changed(self, *_):
        if self._profile_change_guard:
            return
        requested = self.profile_var.get()
        previous = self._active_profile_label
        if requested not in PROFILE_LABELS or (self.running and requested != previous):
            self._profile_change_guard = True
            try:
                self.profile_var.set(previous)
            finally:
                self._profile_change_guard = False
            return
        if self._show_profile_page(requested) and requested != previous:
            self._paths_dirty = True

    def _profile_label_for_target(self, target):
        if not isinstance(target, str):
            return None
        target_norm = target.strip().replace("/", "\\").lower()
        if target_norm.endswith("\\02\\b模板.xls") or "\\02\\b模板.xls" in target_norm:
            return "支线 02"
        if "\\01\\" in target_norm and target_norm.rsplit("\\", 1)[-1].startswith("b模板"):
            return "主线 01"
        return None

    def _parse_drop_paths(self, data):
        if isinstance(data, (list, tuple)):
            return [Path(os.fsdecode(item)) for item in data if item]
        try:
            return [Path(item) for item in self.tk.splitlist(data) if item.strip()]
        except (tk.TclError, TypeError, ValueError):
            return []

    def _handle_drop(self, event, variable, kind, profile_label=None):
        return self._handle_drop_paths(getattr(event, "data", ""), variable, kind, profile_label)

    def _handle_drop_paths(self, data, variable, kind, profile_label=None):
        if self.running:
            return "break"
        paths = self._parse_drop_paths(data)
        if len(paths) != 1:
            messagebox.showwarning("无法导入", "请每次只拖入一个 Excel 文件。", parent=self)
            return "break"
        path = paths[0]
        if path.suffix.lower() not in {".xls", ".xlsx", ".xlsm"}:
            messagebox.showwarning("文件类型不支持", "请选择 .xls、.xlsx 或 .xlsm 文件。", parent=self)
            return "break"
        if kind != "output" and not path.is_file():
            messagebox.showwarning("文件不存在", "拖入的 Excel 文件不存在。", parent=self)
            return "break"
        profile_label = profile_label or self._active_profile_label
        if kind == "target":
            profile_label = self._resolve_target_profile(path, profile_label)
            variable = self._profile_path_vars[profile_label]["target"]
        variable.set(str(path))
        if kind in {"source", "target"}:
            self._set_default_output(profile_label)
        return "break"

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
            profile_label = self._resolve_target_profile(Path(path), self._active_profile_label)
            self._profile_path_vars[profile_label]["target"].set(path)
            self._set_default_output(profile_label)

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            confirmoverwrite=False,
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
            **self._browse_options(self.output_var, save=True),
        )
        if path:
            self.output_var.set(path)

    def _choose_existing_output_mode(self, output):
        return OutputModeDialog.ask(self, output.name)

    def _set_output_action(self, profile_label, output_mode, completed=False, output_exists=True):
        if not output_exists:
            action = "已生成" if completed else "将生成"
        elif output_mode == "AppendExisting":
            action = "已继续写入" if completed else "将继续写入"
        elif output_mode == "Replace":
            action = "已替换" if completed else "将替换"
        else:
            action = ""
        suffix = f"（{action}）" if action else ""
        self._profile_output_labels[profile_label].set("输出文件" + suffix)

    def _effective_output_path(self, output, target):
        if output.suffix or not target.suffix:
            return output
        return output.with_name(output.name + target.suffix)

    def _set_default_output(self, profile_label=None):
        profile_label = profile_label or self._active_profile_label
        path_vars = self._profile_path_vars[profile_label]
        target = path_vars["target"].get().strip()
        if target and not path_vars["output"].get().strip():
            target_path = Path(target)
            path_vars["output"].set(str(target_path.with_name(target_path.stem + "_已填充.xlsx")))

    def _resolve_target_profile(self, path, current_label):
        suggested = self._profile_label_for_target(str(path))
        if suggested and suggested != current_label:
            should_switch = messagebox.askyesno(
                "切换方案",
                f"该目标文件看起来属于“{suggested}”。是否切换到对应页面？",
                parent=self,
            )
            if should_switch:
                self._show_profile_page(suggested)
                self._paths_dirty = True
                return suggested
        return current_label

    def _select_profile_for_target(self, target):
        label = self._profile_label_for_target(target)
        if label:
            self._show_profile_page(label)

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
        effective_output = self._effective_output_path(output, target)
        if effective_output.resolve() in (target.resolve(), source.resolve()):
            messagebox.showwarning("输出路径无效", "输出文件不能覆盖文件 A 或文件 B。")
            return
        if effective_output.exists() and not effective_output.is_file():
            messagebox.showwarning("输出路径无效", "输出路径必须是 Excel 文件，不能是文件夹。")
            return
        output_exists = effective_output.is_file()
        output_mode = "Replace"
        if output_exists:
            output_mode = self._choose_existing_output_mode(effective_output)
            if output_mode is None:
                return
        self._set_output_action(self._active_profile_label, output_mode, output_exists=output_exists)
        self.status_var.set("正在处理...")
        self.last_output = None
        self._profile_last_outputs[self._active_profile_label] = None
        self.open_output_button.configure(state="disabled")
        self.summary_var.set("正在生成输出文件...")
        self.status_label.configure(style="Hint.TLabel")
        self.progress.start(12)
        self.running = True
        self.run_button.configure(state="disabled", text="正在填充...")
        for control in self.all_file_controls:
            control.configure(state="disabled")
        self.profile_combo.configure(state="disabled")
        row_mode = self.row_mode_var.get()
        for button in self.row_mode_buttons:
            button.configure(state="disabled")
        self._write_log(f"开始处理：{source.name} -> {target.name}")
        self._write_log("导出范围：" + ("仅可见行" if row_mode == "Visible" else "全部数据行（包含隐藏行）"))
        profile = PROFILE_LABELS.get(self.profile_var.get(), "Mainline")
        self._write_log("映射方案：" + self.profile_var.get())
        output_action = (
            "在现有输出中继续写入"
            if output_mode == "AppendExisting"
            else ("替换现有输出" if output_exists else "生成新输出")
        )
        self._write_log("输出处理：" + output_action)
        threading.Thread(
            target=self._run_worker,
            args=(source, target, output, row_mode, profile, output_mode),
            kwargs={"output_exists": output_exists},
            daemon=True,
        ).start()

    def _run_worker(self, source, target, output, row_mode="Visible", profile="Mainline",
                    output_mode="Replace", output_exists=True):
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
            "-Profile",
            profile,
            "-RowMode",
            row_mode,
            "-OutputMode",
            output_mode,
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            raw = completed.stdout.strip().splitlines()
            payload = json.loads(raw[-1]) if raw else {"success": False, "message": completed.stderr.strip()}
            self.after(0, self._finish, payload, completed.returncode, output_mode, output_exists)
        except Exception as exc:  # pragma: no cover - UI error path
            self.after(0, self._finish, {"success": False, "message": str(exc)}, 1, output_mode, output_exists)

    def _finish(self, payload, returncode, output_mode=None, output_exists=True):
        self.running = False
        self.progress.stop()
        self.progress.configure(value=0)
        self.run_button.configure(state="normal", text="开始填充")
        for control in self.all_file_controls:
            control.configure(state="normal")
        self.profile_combo.configure(state="readonly")
        for button in self.row_mode_buttons:
            button.configure(state="normal")
        if payload.get("success") and returncode == 0:
            self.summary_var.set(f"读取 {payload.get('rowsRead', 0)} 行    写入 {payload.get('rowsWritten', 0)} 行    跳过隐藏 {payload.get('rowsHiddenSkipped', 0)} 行")
            output_path = payload.get("output") or self.output_var.get()
            if output_path:
                self.last_output = Path(output_path).resolve()
                self._profile_last_outputs[self._active_profile_label] = self.last_output
                self.output_var.set(str(self.last_output))
                self.open_output_button.configure(state="normal")
            self._set_output_action(self._active_profile_label, output_mode or "Replace", completed=True,
                                    output_exists=output_exists)
            self.status_var.set("处理完成")
            self.status_label.configure(style="Success.TLabel")
            self._write_log(payload.get("message", "映射完成。"))
            self._write_log(f"输出文件：{payload.get('output', self.output_var.get())}")
            self._write_log(f"读取 {payload.get('rowsRead', 0)} 行，写入 {payload.get('rowsWritten', 0)} 行。")
            if payload.get('writeStartRow') and payload.get('rowsWritten', 0):
                self._write_log(f"本批次从第 {payload['writeStartRow']} 行开始写入；已有数据末行：{payload.get('existingLastRow', 0)}。")
            self._write_log(f"跳过隐藏行：{payload.get('rowsHiddenSkipped', 0)} 行。")
            for item in payload.get("skipped", []):
                self._write_log("跳过：" + item)
            if self.completion_dialog is not None and self.completion_dialog.winfo_exists():
                self.completion_dialog.destroy()
            self.completion_dialog = CompletionDialog(self, payload.get("message", "映射完成。"))
        else:
            self.summary_var.set("本次处理失败，详情见处理记录")
            self.last_output = None
            self._profile_last_outputs[self._active_profile_label] = None
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

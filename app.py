import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
MAPPER = APP_DIR / "excel_mapper.ps1"


class ShelfAssistant(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("沃尔玛上架助手")
        self.geometry("760x600")
        self.minsize(680, 560)
        self.configure(bg="#f4f6f8")
        self.source_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.row_mode_var = tk.StringVar(value="Visible")
        self.status_var = tk.StringVar(value="请选择文件 A 和文件 B")
        self.running = False
        self._build_ui()

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 20, "bold"))
        style.configure("Hint.TLabel", foreground="#5b6573")
        style.configure("Run.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=(18, 9))

        outer = ttk.Frame(self, padding=28)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="沃尔玛上架助手", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="按多级表头名称将文件 A 的商品信息填入文件 B，不覆盖原始模板。",
            style="Hint.TLabel",
        ).pack(anchor="w", pady=(5, 24))

        panel = ttk.LabelFrame(outer, text="文件")
        panel.pack(fill="x")
        panel.columnconfigure(1, weight=1)
        self._file_row(panel, 0, "文件 A（源数据）", self.source_var, self._choose_source)
        self._file_row(panel, 1, "文件 B（沃尔玛模板）", self.target_var, self._choose_target)
        self._file_row(panel, 2, "输出文件", self.output_var, self._choose_output)

        options = ttk.LabelFrame(outer, text="导出范围")
        options.pack(fill="x", pady=(18, 0))
        self.row_mode_buttons = []
        for value, label in (
            ("Visible", "仅可见行（跳过所有隐藏行，连续排列）"),
            ("All", "全部数据行（包含隐藏行，连续排列）"),
        ):
            button = ttk.Radiobutton(options, text=label, variable=self.row_mode_var, value=value)
            button.pack(anchor="w", padx=12, pady=7)
            self.row_mode_buttons.append(button)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(20, 0))
        self.run_button = ttk.Button(bottom, text="开始填充", style="Run.TButton", command=self.run_mapping)
        self.run_button.pack(side="left")
        ttk.Label(bottom, textvariable=self.status_var, style="Hint.TLabel").pack(side="left", padx=16)

        log_frame = ttk.LabelFrame(outer, text="处理记录")
        log_frame.pack(fill="both", expand=True, pady=(18, 0))
        self.log = tk.Text(log_frame, height=7, state="disabled", bg="#ffffff", relief="flat")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

    def _file_row(self, parent, row, label, variable, command):
        ttk.Label(parent, text=label, width=20).grid(row=row, column=0, sticky="w", padx=(12, 8), pady=9)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=9)
        ttk.Button(parent, text="浏览...", command=command).grid(row=row, column=2, padx=(8, 12), pady=9)

    def _choose_source(self):
        path = filedialog.askopenfilename(
            title="选择文件 A",
            filetypes=[("Excel 文件", "*.xls;*.xlsx;*.xlsm"), ("所有文件", "*.*")],
        )
        if path:
            self.source_var.set(path)
            self._set_default_output()

    def _choose_target(self):
        path = filedialog.askopenfilename(
            title="选择文件 B",
            filetypes=[("Excel 文件", "*.xlsx;*.xlsm;*.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.target_var.set(path)
            self._set_default_output()

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
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
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def run_mapping(self):
        if self.running:
            return
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
        self.running = True
        self.run_button.configure(state="disabled")
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
        self.run_button.configure(state="normal")
        for button in self.row_mode_buttons:
            button.configure(state="normal")
        if payload.get("success") and returncode == 0:
            self.status_var.set("处理完成")
            self._write_log(payload.get("message", "映射完成。"))
            self._write_log(f"输出文件：{payload.get('output', self.output_var.get())}")
            self._write_log(f"读取 {payload.get('rowsRead', 0)} 行，写入 {payload.get('rowsWritten', 0)} 行。")
            self._write_log(f"跳过隐藏行：{payload.get('rowsHiddenSkipped', 0)} 行。")
            for item in payload.get("skipped", []):
                self._write_log("跳过：" + item)
            messagebox.showinfo("处理完成", payload.get("message", "映射完成。"))
        else:
            self.status_var.set("处理失败")
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

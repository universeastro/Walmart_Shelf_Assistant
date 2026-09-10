"""CompletionDialog 的冒烟测试。

app.py 用自绘 Toplevel 替换了 messagebox.showinfo。这个类的风险不在视觉，
而在一条具体的失败路径：__init__ 里 `self._center()` 在 `self.deiconify()`
**之前**被调用，而 _center() 会走进 ctypes/Win32 分支。只要那里抛一次异常，
deiconify() 就被跳过，弹窗永久停在 withdraw 状态——**用户看不到「处理完成」**，
而且不报错。本测试就是钉住这条路径。

用法：py tests/smoke_completion_dialog.py
退出码 0 全过 / 1 有失败。需要可见的桌面会话（tkinter 要建真窗口）。
"""

import sys
import tkinter as tk
from pathlib import Path

# 控制台代码页是 GBK 时中文会乱码。与仓库里的 PowerShell 脚本同样处理。
if sys.platform == "win32":
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# 允许从仓库根目录之外运行
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app  # noqa: E402

failures = []


def check(cond, msg):
    if cond:
        print(f"  OK   {msg}")
    else:
        print(f"  FAIL {msg}")
        failures.append(msg)


root = tk.Tk()
root.title("smoke")
root.geometry("760x600+120+120")
root.update()

print("== 1. 构造后可见（deiconify 没有被跳过）==")
dlg = app.CompletionDialog(root, "映射完成。")
root.update()
check(dlg.winfo_exists() == 1, "dialog exists")
check(
    dlg.winfo_viewable() == 1,
    "dialog is viewable - _center() did not abort before deiconify()",
)

print("== 2. 相对父窗口居中 ==")
root.update_idletasks()
dx = (dlg.winfo_rootx() + dlg.winfo_width() // 2) - (
    root.winfo_rootx() + root.winfo_width() // 2
)
dy = (dlg.winfo_rooty() + dlg.winfo_height() // 2) - (
    root.winfo_rooty() + root.winfo_height() // 2
)
print(f"   offset from parent centre: dx={dx} dy={dy}")
# 容差 40px：Win32 路径用的是含标题栏/边框的窗口矩形，而 winfo_* 是客户区，
# 两者中心天然差十几像素。未居中的话偏差会是几百像素，区分度足够。
check(abs(dx) <= 40 and abs(dy) <= 40, "dialog is centred on the parent window")

print("== 3. Escape / Return 关闭 ==")
dlg.event_generate("<Escape>")
root.update()
check(dlg.winfo_exists() == 0, "Escape destroys the dialog")

dlg = app.CompletionDialog(root, "第二次。")
root.update()
dlg.event_generate("<Return>")
root.update()
check(dlg.winfo_exists() == 0, "Return destroys the dialog")

print("== 4. 连续两次运行不会堆叠弹窗 ==")
# 复刻 app.py 的调用序列：旧弹窗先 destroy，再建新的。
first = app.CompletionDialog(root, "第一次。")
root.update()
if first is not None and first.winfo_exists():
    first.destroy()
second = app.CompletionDialog(root, "第二次。")
root.update()
check(first.winfo_exists() == 0, "previous dialog was destroyed")
check(second.winfo_exists() == 1, "new dialog exists")
check(second.winfo_viewable() == 1, "new dialog is viewable")

print("== 5. destroy 后父窗口的 <Configure> 绑定被摘掉 ==")
# 每建一次弹窗就往父窗口挂一个 <Configure> 回调；不摘掉的话，
# 跑 N 次就会累积 N 个回调，每次移动窗口都全部触发。
# 弹窗活着时应当有绑定，destroy 之后应当没有——两次都要断言，
# 否则「两边都是空」也会让这个检查悄悄通过。
while_alive = root.bind("<Configure>")
print(f"   while dialog alive : {while_alive!r}")
check(while_alive != "", "dialog registered a <Configure> binding on the parent")
second.destroy()
root.update()
after_close = root.bind("<Configure>")
print(f"   after dialog closed: {after_close!r}")
check(after_close == "", "no <Configure> binding leaked onto the parent")

root.destroy()

print("== 6. 真实 _finish 路径（含 completion_dialog 的复用守卫）==")
# 上面几节用的是裸 tk.Tk；这里换成真的 ShelfAssistant，让 _finish 里那段
# 「先 destroy 旧弹窗再建新的」守卫也被真正执行到。两个 Tk 根不能共存，
# 所以放在 root.destroy() 之后。
win = app.ShelfAssistant()
# 挪到屏幕外而不是 withdraw()：winfo_viewable() 要求自身与所有祖先都已映射，
# 对一个 transient 子窗口来说父窗口被 withdraw 会让它必然报告 0——
# 那是测试自己制造的假阳性，不是弹窗的问题。
win.geometry("760x600+4000+4000")
win.update()

payload = {
    "success": True,
    "message": "映射完成。",
    "output": "out.xlsx",
    "rowsRead": 1,
    "rowsWritten": 1,
    "rowsHiddenSkipped": 0,
    "skipped": [],
}

win._finish(payload, 0)
win.update()
check(
    win.completion_dialog is not None and win.completion_dialog.winfo_exists() == 1,
    "first _finish creates a dialog",
)
check(win.completion_dialog.winfo_viewable() == 1, "first dialog is viewable")

first = win.completion_dialog
win._finish(payload, 0)
win.update()
check(first.winfo_exists() == 0, "second _finish destroys the previous dialog")
check(
    win.completion_dialog is not first and win.completion_dialog.winfo_exists() == 1,
    "second _finish creates a fresh dialog instead of stacking",
)

win.destroy()

print()
if failures:
    print(f"FAILED ({len(failures)}):")
    for f in failures:
        print(f"  - {f}")
    sys.exit(1)
print("PASS: CompletionDialog shows, centres, closes, and does not leak bindings.")
sys.exit(0)

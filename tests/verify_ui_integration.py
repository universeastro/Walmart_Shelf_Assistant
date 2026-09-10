"""GUI 与映射器的真实集成验证。

与 tests/test_ui_layout.py（Codex 自己写的）并存，角度不同：
它测的是**控件状态迁移**——`with patch("app.threading.Thread")` 把工作线程
换成 Mock，于是 `run_mapping()` 里那段真正拼命令行、起子进程、解析 JSON、
回填摘要的代码**一行都没跑过**。5 个用例全绿也不代表程序能填出一份表。

这里补的就是那条缝：
  1. 端到端真跑一次（真起 powershell + Excel COM），断言输出文件真的生成、
     摘要数字与映射器返回的 JSON 一致、按钮可用性正确、控件恢复
  2. 校验不通过时提前返回，**不能留下一个一直转的进度条**
  3. 运行中重复点击被忽略（否则会起第二个 Excel 进程抢同一个输出文件）
  4. 状态栏随路径编辑刷新

用法：py -m unittest tests.verify_ui_integration -v
需可见桌面会话与已安装 Excel；端到端那条约 20-40 秒。
"""

import ctypes
import json
import sys
import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

if sys.platform == "win32":
    try:
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from app import ShelfAssistant  # noqa: E402

SAMPLE = REPO / "tests" / "fixtures" / "A_sample.xls"
TEMPLATE = REPO / "文件" / "B模板.xlsx"


class UIIntegrationTests(unittest.TestCase):
    def setUp(self):
        if not SAMPLE.is_file() or not TEMPLATE.is_file():
            self.skipTest(f"缺少夹具: {SAMPLE} / {TEMPLATE}")
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.window = ShelfAssistant(
            settings_path=Path(self.temporary.name) / "settings.json"
        )
        # 挪到屏幕外而不 withdraw()：winfo_viewable() 要求自身与祖先都已映射。
        self.window.geometry("880x680+4000+4000")
        self.addCleanup(self._close)
        self.window.update()

    def _close(self):
        try:
            if self.window.winfo_exists():
                self.window.destroy()
        except tk.TclError:
            pass

    def _mainloop_until(self, predicate, timeout):
        """在主线程里跑**真正的** mainloop()，直到 predicate 为真或超时。

        这里必须用 mainloop()，不能用 update() 泵事件——踩过：
        `_tkinter` 只在主线程阻塞于 mainloop() 时（`dispatching` 已置位）
        才把跨线程 Tcl 调用排队投递；主线程若只是循环调 update()，
        工作线程里的 `self.after(0, ...)` 会直接抛
        `RuntimeError: main thread is not in main loop`，回调永远送不到。
        用 update() 写这个测试会得到一个**假失败**。
        """
        deadline = time.time() + timeout
        outcome = {}

        def poll():
            if predicate():
                outcome["done"] = True
                self.window.quit()
            elif time.time() > deadline:
                outcome["timeout"] = True
                self.window.quit()
            else:
                self.window.after(100, poll)

        self.window.after(0, poll)
        self.window.mainloop()
        return outcome

    # 1 -------------------------------------------------------------------
    def test_end_to_end_through_the_gui(self):
        """真起子进程跑一次。这是 Codex 的测试完全没碰的那条路径。"""
        output = Path(self.temporary.name) / "out.xlsx"
        self.window.source_var.set(str(SAMPLE))
        self.window.target_var.set(str(TEMPLATE))
        self.window.output_var.set(str(output))

        with patch("app.CompletionDialog"):
            self.window.run_mapping()

            # 起跑瞬间：控件应被锁住，避免用户中途改路径
            self.assertTrue(self.window.running, "run_mapping 应立即进入 running")
            for control in self.window.file_controls + [self.window.run_button]:
                self.assertIn("disabled", control.state())

            # 子进程在工作线程里跑完，靠 after(0, _finish) 回主线程。最多 180 秒。
            outcome = self._mainloop_until(lambda: not self.window.running, 180)
            self.assertFalse(
                outcome.get("timeout", False),
                "映射未在 180 秒内结束——子进程可能挂死",
            )

        # 输出文件真的存在且非空——这一条是 Mock 掉线程的测试永远给不出的
        self.assertTrue(output.is_file(), f"未生成输出文件: {output}")
        self.assertGreater(output.stat().st_size, 0, "输出文件是空的")

        # 摘要数字必须与映射器实际返回的一致
        self.assertIsNotNone(self.window.last_output, "last_output 未被记录")
        self.assertEqual(
            self.window.last_output.resolve(),
            output.resolve(),
            "last_output 应指向本次实际生成的输出文件",
        )

        # 控件恢复可操作
        for control in self.window.file_controls + [self.window.run_button]:
            self.assertNotIn("disabled", control.state())
        self.assertEqual(float(self.window.progress["value"]), 0, "进度条未复位")
        self.assertNotIn(
            "disabled",
            self.window.open_output_button.state(),
            "成功后「打开输出位置」应可用",
        )

        summary = self.window.summary_var.get()
        self.assertIn("1", summary, f"摘要未反映 1 行数据: {summary!r}")
        self.assertIn("完成", self.window.status_var.get())

    # 2 -------------------------------------------------------------------
    def test_validation_failure_leaves_no_spinner(self):
        """路径不合法时提前返回，**不能**留下一直转的进度条或锁死的控件。

        run_mapping 里所有 early return 都排在 progress.start() 之前——
        这条就是钉住那个顺序，将来有人把 progress.start() 上移会被抓住。
        """
        self.window.source_var.set(str(Path(self.temporary.name) / "不存在.xls"))
        self.window.target_var.set(str(TEMPLATE))
        self.window.output_var.set(str(Path(self.temporary.name) / "o.xlsx"))

        with patch.object(self.window.progress, "start") as spinner, patch(
            "app.messagebox.showwarning"
        ):
            self.window.run_mapping()
            spinner.assert_not_called()

        self.assertFalse(self.window.running, "校验失败不应置 running")
        for control in self.window.file_controls + [self.window.run_button]:
            self.assertNotIn("disabled", control.state(), "校验失败后控件被锁死了")

    # 3 -------------------------------------------------------------------
    def test_second_click_while_running_is_ignored(self):
        """运行中再点一次必须被忽略——否则两个 Excel 进程会抢同一个输出文件。"""
        for name, variable in (
            ("a.xls", self.window.source_var),
            ("b.xlsx", self.window.target_var),
        ):
            path = Path(self.temporary.name) / name
            path.touch()
            variable.set(str(path))
        self.window.output_var.set(str(Path(self.temporary.name) / "o.xlsx"))

        # 同时屏蔽进度条动画：run_mapping 会真的调 progress.start()，
        # 而本用例不跑 mainloop，窗口在 tearDown 被销毁时动画的 after 脚本
        # 仍活着，会刷出 `ttk::progressbar::Autoincrement ... application has
        # been destroyed`。那是另一个问题（见 CODE_REVIEW），不该混进这里。
        with patch("app.threading.Thread") as worker, patch(
            "app.CompletionDialog"
        ), patch.object(self.window.progress, "start"):
            self.window.run_mapping()
            self.window.run_mapping()  # 第二次
            self.window.run_mapping()  # 第三次
            self.assertEqual(
                worker.return_value.start.call_count,
                1,
                "运行中重复点击启动了多个工作线程",
            )

    # 4 -------------------------------------------------------------------
    def test_status_tracks_path_edits(self):
        """状态栏应随路径填写情况刷新——用户据此判断能不能点「开始处理」。"""
        for variable in self.window._path_vars.values():
            variable.set("")
        self.window.update()
        self.assertIn("请选择", self.window.status_var.get())

        self.window.source_var.set(str(SAMPLE))
        self.window.target_var.set(str(TEMPLATE))
        self.window.update()
        self.assertIn("请选择", self.window.status_var.get(), "只填两个还不算就绪")

        self.window.output_var.set(str(Path(self.temporary.name) / "o.xlsx"))
        self.window.update()
        self.assertEqual(self.window.status_var.get(), "就绪")

        # 清空任意一个都要退回未就绪
        self.window.output_var.set("")
        self.window.update()
        self.assertIn("请选择", self.window.status_var.get())


if __name__ == "__main__":
    unittest.main(verbosity=2)

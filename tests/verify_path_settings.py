"""路径持久化设置的独立验证。

与 tests/test_path_settings.py（Codex 自己写的）并存，但刻意不重复它的角度——
它验证"功能做对了没有"，这里验证"边界会不会咬人"。

Codex 的测试覆盖了：重启恢复、改动与清空被保存、文件被移走后清空、
目录不被当作输入、缺失/损坏/类型错误的配置、保存失败保留旧值。

这里补的是它没覆盖的：
  1. 什么都没改时**不应该**在用户磁盘上生成配置文件（无谓留垃圾）
  2. run_mapping() 里新增的 _save_paths() 调用点——Codex 的测试全走 destroy()，
     没有一条经过这个入口
  3. 落盘的中文必须是 UTF-8 而不是 \\uXXXX 转义（配置文件要人能读）
  4. 成功保存后不留 .tmp
  5. 损坏的配置文件被忽略，但**不能被删掉**（那是用户的文件）

用法：py -m unittest tests.verify_path_settings -v
"""

import json
import os
import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import ShelfAssistant  # noqa: E402


class PathSettingsBoundaryTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.settings = Path(self.temporary.name) / "config" / "settings.json"

    def open_app(self):
        window = ShelfAssistant(settings_path=self.settings)
        window.withdraw()

        def close_window():
            try:
                exists = window.winfo_exists()
            except tk.TclError:
                return
            if exists:
                window.destroy()

        self.addCleanup(close_window)
        return window

    def seed(self, text):
        self.settings.parent.mkdir(parents=True, exist_ok=True)
        self.settings.write_text(text, encoding="utf-8")

    # 1 -------------------------------------------------------------------
    def test_nothing_changed_writes_no_file(self):
        """没动过任何路径就不该写盘——否则每次开一下程序都留下一个文件。"""
        window = self.open_app()
        window.destroy()
        self.assertFalse(
            self.settings.exists(),
            "打开又关闭、且未改动任何路径时，不应生成 settings.json",
        )

    # 2 -------------------------------------------------------------------
    def test_run_mapping_saves_before_validating(self):
        """run_mapping 开头的 _save_paths() 是新增调用点，Codex 的测试没走过。

        用不存在的源文件让它在校验处提前返回——这样既不启动 Excel，
        又能确认保存确实发生在校验之前。
        """
        window = self.open_app()
        window.source_var.set(str(Path(self.temporary.name) / "missing.xls"))
        window.output_var.set("out.xlsx")

        with patch("app.messagebox.showwarning") as warning:
            window.run_mapping()
            warning.assert_called_once()

        self.assertTrue(self.settings.exists(), "run_mapping 应当先保存路径再校验")
        data = json.loads(self.settings.read_text(encoding="utf-8"))
        self.assertEqual(data["output"], "out.xlsx")

    # 3 -------------------------------------------------------------------
    def test_chinese_paths_are_stored_as_readable_utf8(self):
        """ensure_ascii=False 的效果要落到盘上——配置文件是给人看的。"""
        window = self.open_app()
        chinese = str(Path(self.temporary.name) / "中文 目录" / "文件 a.xls")
        window.output_var.set(chinese)
        window.destroy()

        raw = self.settings.read_bytes()
        self.assertIn(
            "中文".encode("utf-8"),
            raw,
            "中文应以 UTF-8 字节存储，而不是 \\uXXXX 转义",
        )
        self.assertEqual(json.loads(raw.decode("utf-8"))["output"], chinese)

    # 4 -------------------------------------------------------------------
    def test_successful_save_leaves_no_temp_file(self):
        window = self.open_app()
        window.source_var.set("a.xls")
        window.destroy()
        self.assertEqual(
            list(self.settings.parent.glob("*.tmp")),
            [],
            "成功保存后不应残留 .tmp 文件",
        )

    # 5 -------------------------------------------------------------------
    def test_corrupt_settings_file_is_ignored_but_not_deleted(self):
        """损坏的配置要被忽略，但文件本身是用户的，不能顺手删掉。"""
        self.seed("{broken")
        window = self.open_app()
        self.assertEqual(window.source_var.get(), "")
        window.destroy()
        self.assertEqual(
            self.settings.read_text(encoding="utf-8"),
            "{broken",
            "损坏的 settings.json 不应被删除或被覆盖",
        )

    # 7 -------------------------------------------------------------------
    def test_wrong_type_is_rejected_by_the_type_check_not_by_existence(self):
        """Codex 的 test_missing_corrupt_and_wrong_type_settings 中，
        {"target": "valid.xlsx"} 的 target 是**合法字符串**，它被清空是因为
        文件不存在，而不是因为类型检查——所以那个用例并没有隔离 isinstance 分支。
        这里用非字符串值把该分支单独钉住：即使路径指向一个真实存在的文件，
        只要类型不对就必须被拒。
        """
        existing = Path(self.temporary.name) / "real.xlsx"
        existing.touch()

        # 非字符串：必须被 isinstance 拦下
        self.seed(json.dumps({"target": 12}))
        window = self.open_app()
        self.assertEqual(window.target_var.get(), "", "非字符串的 target 必须被拒绝")
        window.destroy()

        # 同样一个真实存在的文件，用字符串给出时应当被接受——
        # 这才说明上一条拒绝的是「类型」而不是「文件不存在」。
        self.seed(json.dumps({"target": str(existing)}))
        window = self.open_app()
        self.assertEqual(window.target_var.get(), str(existing))

    # 6 -------------------------------------------------------------------
    def test_output_path_is_kept_even_if_it_does_not_exist(self):
        """README 承诺「输出路径保留，即使输出文件尚未生成」——单独钉一条。"""
        missing = str(Path(self.temporary.name) / "not-created-yet.xlsx")
        self.seed(json.dumps({"source": "", "target": "", "output": missing}))
        window = self.open_app()
        self.assertEqual(window.output_var.get(), missing)


if __name__ == "__main__":
    unittest.main(verbosity=2)

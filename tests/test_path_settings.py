import json
import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import ShelfAssistant


class PathSettingsTests(unittest.TestCase):
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

    def input_paths(self):
        source = Path(self.temporary.name) / "\u6d4b\u8bd5 a.xls"
        target = Path(self.temporary.name) / "template.xlsx"
        source.touch()
        target.touch()
        return {"source": str(source), "target": str(target), "output": str(Path(self.temporary.name) / "new output.xlsx")}

    def saved_profile(self, profile="Mainline"):
        data = json.loads(self.settings.read_text(encoding="utf-8"))
        self.assertEqual(data["version"], 2)
        return data["profiles"][profile]

    def switch(self, window, label):
        window.profile_var.set(label)
        window._on_profile_selected()

    def test_close_and_restart_restores_existing_inputs_and_new_output(self):
        window = self.open_app()
        paths = self.input_paths()
        for key, value in paths.items():
            window._path_vars[key].set(value)
        window.destroy()
        restored = self.open_app()
        self.assertEqual({key: var.get() for key, var in restored._path_vars.items()}, paths)

    def test_changes_and_clearing_are_saved(self):
        paths = self.input_paths()
        self.seed(json.dumps(paths))
        window = self.open_app()
        window.source_var.set("new.xls")
        window.output_var.set("")
        window.destroy()
        self.assertEqual(
            self.saved_profile(),
            {"source": "new.xls", "target": paths["target"], "output": "", "row_mode": "Visible"},
        )

    def test_moved_input_is_cleared_without_changing_other_paths(self):
        for key in ("source", "target"):
            with self.subTest(key=key):
                paths = self.input_paths()
                self.seed(json.dumps(paths))
                original = Path(paths[key])
                original.rename(original.with_name("moved-" + original.name))
                window = self.open_app()
                expected = dict(paths, **{key: ""})
                self.assertEqual({name: var.get() for name, var in window._path_vars.items()}, expected)
                window.destroy()
                self.assertEqual(self.saved_profile(), {**expected, "row_mode": "Visible"})

    def test_directory_is_not_restored_as_input(self):
        self.seed(json.dumps({"source": self.temporary.name}))
        window = self.open_app()
        self.assertEqual(window.source_var.get(), "")

    def test_missing_corrupt_and_wrong_type_settings(self):
        for content in (None, "{broken", "[]", '{"source": 12, "target": "valid.xlsx"}'):
            with self.subTest(content=content):
                if content is not None:
                    self.seed(content)
                window = self.open_app()
                self.assertEqual(window.source_var.get(), "")
                self.assertEqual(window.target_var.get(), "")
                window.destroy()

    def test_save_failure_preserves_previous_settings(self):
        self.seed('{"source": "old.xls"}')
        window = self.open_app()
        window.source_var.set("new.xls")
        with patch("app.os.replace", side_effect=PermissionError("locked")), patch("app.messagebox.showwarning") as warning:
            window._save_paths()
            warning.assert_called_once()
        self.assertEqual(json.loads(self.settings.read_text(encoding="utf-8")), {"source": "old.xls"})
        self.assertEqual(list(self.settings.parent.glob("*.tmp")), [])
        window._save_paths()
        self.assertFalse(window._paths_dirty)

    def test_profiles_keep_independent_paths_and_restore_last_page(self):
        mainline = self.input_paths()
        req_source = Path(self.temporary.name) / "req-source.xls"
        req_target = Path(self.temporary.name) / "req-target.xls"
        req_source.touch()
        req_target.touch()
        req02 = {
            "source": str(req_source),
            "target": str(req_target),
            "output": str(Path(self.temporary.name) / "req-output.xlsx"),
        }

        window = self.open_app()
        for key, value in mainline.items():
            window._path_vars[key].set(value)
        self.switch(window, "支线 02")
        for key, value in req02.items():
            window._path_vars[key].set(value)
        self.switch(window, "主线 01")
        self.assertEqual({key: var.get() for key, var in window._path_vars.items()}, mainline)
        self.switch(window, "支线 02")
        self.assertEqual({key: var.get() for key, var in window._path_vars.items()}, req02)
        window.destroy()

        restored = self.open_app()
        self.assertEqual(restored.profile_var.get(), "支线 02")
        self.assertEqual({key: var.get() for key, var in restored._path_vars.items()}, req02)
        self.switch(restored, "主线 01")
        self.assertEqual({key: var.get() for key, var in restored._path_vars.items()}, mainline)

    def test_legacy_req02_paths_migrate_to_req02_page(self):
        source = Path(self.temporary.name) / "source.xls"
        target = Path(self.temporary.name) / "02" / "B模板.xls"
        source.touch()
        target.parent.mkdir()
        target.touch()
        legacy = {"source": str(source), "target": str(target), "output": "req-output.xlsx"}
        self.seed(json.dumps(legacy))

        window = self.open_app()
        self.assertEqual(window.profile_var.get(), "支线 02")
        self.assertEqual({key: var.get() for key, var in window._path_vars.items()}, legacy)
        self.switch(window, "主线 01")
        self.assertEqual({key: var.get() for key, var in window._path_vars.items()}, dict.fromkeys(("source", "target", "output"), ""))
        window.destroy()
        data = json.loads(self.settings.read_text(encoding="utf-8"))
        self.assertEqual(data["last_profile"], "Mainline")
        self.assertEqual(data["profiles"]["Req02"], {**legacy, "row_mode": "Visible"})

    def test_profiles_keep_independent_row_modes_and_restore_them(self):
        window = self.open_app()
        self.assertEqual(window.row_mode_var.get(), "Visible")
        self.switch(window, "支线 02")
        window.row_mode_var.set("All")
        self.switch(window, "主线 01")
        self.assertEqual(window.row_mode_var.get(), "Visible")
        self.switch(window, "支线 02")
        self.assertEqual(window.row_mode_var.get(), "All")
        window.destroy()

        restored = self.open_app()
        self.assertEqual(restored.profile_var.get(), "支线 02")
        self.assertEqual(restored.row_mode_var.get(), "All")
        self.switch(restored, "主线 01")
        self.assertEqual(restored.row_mode_var.get(), "Visible")

    def test_missing_or_invalid_profile_row_mode_defaults_to_visible(self):
        settings = {
            "version": 2,
            "last_profile": "Req02",
            "profiles": {
                "Mainline": {"source": "", "target": "", "output": "", "row_mode": 12},
                "Req02": {"source": "", "target": "", "output": ""},
            },
        }
        self.seed(json.dumps(settings))
        window = self.open_app()
        self.assertEqual(window.row_mode_var.get(), "Visible")
        self.switch(window, "主线 01")
        self.assertEqual(window.row_mode_var.get(), "Visible")

    def test_req03_paths_and_row_mode_are_independent_and_persisted(self):
        source = Path(self.temporary.name) / "req03-source.xls"
        target = Path(self.temporary.name) / "XPY沃尔玛价格计算.xls"
        source.touch()
        target.touch()
        req03 = {
            "source": str(source),
            "target": str(target),
            "output": str(Path(self.temporary.name) / "req03-output.xls"),
        }

        window = self.open_app()
        self.switch(window, "支线 03")
        for key, value in req03.items():
            window._path_vars[key].set(value)
        window.row_mode_var.set("All")
        self.switch(window, "主线 01")
        self.assertEqual(window.source_var.get(), "")
        self.switch(window, "支线 03")
        self.assertEqual({key: var.get() for key, var in window._path_vars.items()}, req03)
        window.destroy()

        restored = self.open_app()
        self.assertEqual(restored.profile_var.get(), "支线 03")
        self.assertEqual({key: var.get() for key, var in restored._path_vars.items()}, req03)
        self.assertEqual(restored.row_mode_var.get(), "All")


if __name__ == "__main__":
    unittest.main()

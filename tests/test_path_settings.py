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
        self.assertEqual(json.loads(self.settings.read_text(encoding="utf-8")), {"source": "new.xls", "target": paths["target"], "output": ""})

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
                self.assertEqual(json.loads(self.settings.read_text(encoding="utf-8")), expected)

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


if __name__ == "__main__":
    unittest.main()

import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import ShelfAssistant


class UILayoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.window = ShelfAssistant(settings_path=Path(self.temp.name) / "settings.json")
        self.addCleanup(self.window.destroy)
        self.window.update()

    def test_controls_fit_at_supported_sizes(self):
        for size in ("720x620", "880x680", "1100x800"):
            with self.subTest(size=size):
                self.window.geometry(size)
                self.window.update()
                controls = self.window.file_controls + self.window.row_mode_buttons + [
                    self.window.run_button, self.window.status_label, self.window.log]
                for control in controls:
                    self.assertTrue(control.winfo_viewable())
                    left = control.winfo_rootx() - self.window.winfo_rootx()
                    top = control.winfo_rooty() - self.window.winfo_rooty()
                    self.assertGreaterEqual(left, 0)
                    self.assertGreaterEqual(top, 0)
                    self.assertLessEqual(left + control.winfo_width(), self.window.winfo_width())
                    self.assertLessEqual(top + control.winfo_height(), self.window.winfo_height())
                self.assertGreaterEqual(self.window.log.winfo_height(), 60)
                for button in self.window.row_mode_buttons:
                    self.assertGreaterEqual(button.winfo_width(), button.winfo_reqwidth())

    def test_path_entries_use_one_font_and_color(self):
        from tkinter import font as tkfont
        for entry in self.window.file_controls[::2]:
            self.assertEqual(str(entry.cget("foreground")), "#252b32")
            actual = tkfont.Font(root=self.window, font=entry.cget("font")).actual()
            self.assertEqual(actual["family"].lower(), "microsoft yahei ui")
            self.assertEqual(actual["size"], 10)

    def test_busy_controls_and_failure_recovery(self):
        for name, variable in (("a.xls", self.window.source_var), ("b.xlsx", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        self.window.output_var.set(str(Path(self.temp.name) / "output.xlsx"))
        with patch("app.threading.Thread") as worker:
            self.window.run_mapping()
            worker.return_value.start.assert_called_once()
        self.assertTrue(self.window.running)
        for control in self.window.file_controls + self.window.row_mode_buttons + [self.window.run_button]:
            self.assertIn("disabled", control.state())
        with patch("app.messagebox.showerror"):
            self.window._finish({"success": False, "message": "test failure"}, 1)
        self.assertFalse(self.window.running)
        self.assertEqual(float(self.window.progress["value"]), 0)
        for control in self.window.file_controls + self.window.row_mode_buttons + [self.window.run_button]:
            self.assertNotIn("disabled", control.state())

    def test_browse_uses_previous_directory_and_output_name(self):
        output = Path(self.temp.name) / "new.xlsx"
        self.window.output_var.set(str(output))
        options = self.window._browse_options(self.window.output_var, save=True)
        self.assertEqual(options["initialdir"], self.temp.name)
        self.assertEqual(options["initialfile"], "new.xlsx")
        self.assertIs(options["parent"], self.window)
        with patch("app.filedialog.asksaveasfilename", return_value=""):
            self.window._choose_output()
        self.assertEqual(self.window.output_var.get(), str(output))

    def test_output_action_uses_successful_result_not_edited_path(self):
        output = Path(self.temp.name) / "result.xlsx"
        output.touch()
        with patch("app.CompletionDialog"):
            self.window._finish({"success": True, "output": str(output), "rowsRead": 4, "rowsWritten": 3,
                                 "rowsHiddenSkipped": 1}, 0)
            self.window.completion_dialog = None
        self.window.output_var.set(str(Path(self.temp.name) / "different.xlsx"))
        with patch("app.os.startfile", create=True) as open_folder:
            self.window._open_output_location()
            open_folder.assert_called_once_with(str(output.parent))
        self.assertNotIn("disabled", self.window.open_output_button.state())
        self.assertIn("3", self.window.summary_var.get())
        output.unlink()
        with patch("app.os.startfile", create=True) as open_folder, patch("app.messagebox.showwarning") as warning:
            self.window._open_output_location()
            warning.assert_called_once()
            open_folder.assert_not_called()

    def test_path_entries_have_no_hover_popup_binding(self):
        for entry in self.window.file_controls[::2]:
            self.assertEqual(entry.bind("<Enter>"), "")
        self.assertFalse(hasattr(self.window, "_schedule_path_tip"))


if __name__ == "__main__":
    unittest.main()

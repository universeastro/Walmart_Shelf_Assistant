import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from app import OutputModeDialog, ShelfAssistant


class UILayoutTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.window = ShelfAssistant(settings_path=Path(self.temp.name) / "settings.json")
        self.addCleanup(self.window.destroy)
        self.window.update()

    def test_controls_fit_at_supported_sizes(self):
        for profile in ("主线 01", "支线 02", "支线 03"):
            self.window.profile_var.set(profile)
            for size in ("720x620", "880x680", "1100x800"):
                with self.subTest(profile=profile, size=size):
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
        for entry in self.window.all_file_controls[::2]:
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
        for control in self.window.all_file_controls + self.window.row_mode_buttons + [self.window.run_button, self.window.profile_combo]:
            self.assertIn("disabled", control.state())
        with patch("app.messagebox.showerror"):
            self.window._finish({"success": False, "message": "test failure"}, 1)
        self.assertFalse(self.window.running)
        self.assertEqual(float(self.window.progress["value"]), 0)
        for control in self.window.all_file_controls + self.window.row_mode_buttons + [self.window.run_button]:
            self.assertNotIn("disabled", control.state())
        self.assertIn("readonly", self.window.profile_combo.state())

    def test_browse_uses_previous_directory_and_output_name(self):
        output = Path(self.temp.name) / "new.xlsx"
        self.window.output_var.set(str(output))
        options = self.window._browse_options(self.window.output_var, save=True)
        self.assertEqual(options["initialdir"], self.temp.name)
        self.assertEqual(options["initialfile"], "new.xlsx")
        self.assertIs(options["parent"], self.window)
        with patch("app.filedialog.asksaveasfilename", return_value="") as choose:
            self.window._choose_output()
        self.assertEqual(self.window.output_var.get(), str(output))
        self.assertFalse(choose.call_args.kwargs["confirmoverwrite"])

    def test_existing_output_choice_is_passed_to_worker(self):
        for name, variable in (("a.xls", self.window.source_var), ("b.xlsx", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        output = Path(self.temp.name) / "output.xlsx"
        output.touch()
        self.window.output_var.set(str(output))

        with patch("app.OutputModeDialog.ask", return_value="AppendExisting"), patch(
            "app.threading.Thread"
        ) as worker, patch.object(self.window.progress, "start"):
            self.window.run_mapping()
        self.assertEqual(worker.call_args.kwargs["args"][-1], "AppendExisting")
        with patch("app.messagebox.showerror"):
            self.window._finish({"success": False, "message": "test cleanup"}, 1)

    def test_cancel_existing_output_choice_does_not_start(self):
        for name, variable in (("a.xls", self.window.source_var), ("b.xlsx", self.window.target_var),
                               ("output.xlsx", self.window.output_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        with patch("app.OutputModeDialog.ask", return_value=None), patch(
            "app.threading.Thread"
        ) as worker, patch.object(self.window.progress, "start") as spinner:
            self.window.run_mapping()
        worker.assert_not_called()
        spinner.assert_not_called()
        self.assertFalse(self.window.running)

    def test_extensionless_output_checks_effective_existing_file(self):
        for name, variable in (("a.xls", self.window.source_var), ("b.xlsx", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        entered = Path(self.temp.name) / "output"
        effective = entered.with_suffix(".xlsx")
        effective.touch()
        self.window.output_var.set(str(entered))
        with patch("app.OutputModeDialog.ask", return_value=None) as choose, patch(
            "app.threading.Thread"
        ) as worker:
            self.window.run_mapping()
        choose.assert_called_once_with(self.window, effective.name)
        worker.assert_not_called()

    def test_output_mode_dialog_has_explicit_actions(self):
        dialog = OutputModeDialog(self.window, "existing.xlsx")
        self.assertIsNone(dialog.result)
        labels = {child.cget("text") for frame in dialog.winfo_children() for child in frame.winfo_children()
                  if child.winfo_class() == "TFrame" for child in child.winfo_children()
                  if child.winfo_class() == "TButton"}
        self.assertEqual(labels, {"替换文件", "继续写入", "取消"})
        dialog._select("AppendExisting")
        self.assertEqual(dialog.result, "AppendExisting")

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

    def test_successful_replace_shows_actual_output_path_and_action(self):
        output = Path(self.temp.name) / "result.xlsx"
        output.touch()
        self.window.output_var.set(str(output.with_suffix("")))

        with patch("app.CompletionDialog"):
            self.window._finish(
                {"success": True, "output": str(output), "rowsRead": 1, "rowsWritten": 1},
                0,
                "Replace",
            )
            self.window.completion_dialog = None

        self.assertEqual(self.window.output_var.get(), str(output.resolve()))
        self.assertNotEqual(self.window.output_label_var.get(), "输出文件")

    def test_new_output_shows_generated_instead_of_replaced(self):
        output = Path(self.temp.name) / "new-result.xlsx"
        self.window.output_var.set(str(output))

        with patch("app.CompletionDialog"):
            self.window._set_output_action(self.window._active_profile_label, "Replace", output_exists=False)
            self.assertIn("将生成", self.window.output_label_var.get())
            self.window._finish(
                {"success": True, "output": str(output), "rowsRead": 1, "rowsWritten": 1},
                0,
                "Replace",
                False,
            )
            self.window.completion_dialog = None

        self.assertIn("已生成", self.window.output_label_var.get())

    def test_new_output_mode_is_forwarded_as_generation(self):
        for name, variable in (("a.xls", self.window.source_var), ("b.xlsx", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        output = Path(self.temp.name) / "not-created.xlsx"
        self.window.output_var.set(str(output))

        with patch("app.threading.Thread") as worker, patch.object(self.window.progress, "start"):
            self.window.run_mapping()

        self.assertIn("将生成", self.window.output_label_var.get())
        self.assertFalse(worker.call_args.kwargs["kwargs"]["output_exists"])
        with patch("app.messagebox.showerror"):
            self.window._finish({"success": False, "message": "test cleanup"}, 1, "Replace", False)

    def test_output_action_is_scoped_to_profile_page(self):
        output = Path(self.temp.name) / "main-result.xlsx"
        output.touch()
        with patch("app.CompletionDialog"):
            self.window._finish({"success": True, "output": str(output)}, 0)
            self.window.completion_dialog = None
        self.assertEqual(self.window.last_output, output.resolve())
        self.window.profile_var.set("支线 02")
        self.assertIsNone(self.window.last_output)
        self.assertIn("disabled", self.window.open_output_button.state())
        self.window.profile_var.set("主线 01")
        self.assertEqual(self.window.last_output, output.resolve())
        self.assertNotIn("disabled", self.window.open_output_button.state())

    def test_path_entries_have_no_hover_popup_binding(self):
        for entry in self.window.all_file_controls[::2]:
            self.assertEqual(entry.bind("<Enter>"), "")
        self.assertFalse(hasattr(self.window, "_schedule_path_tip"))

    def test_drop_parses_windows_paths_and_rejects_multiple_files(self):
        paths = self.window._parse_drop_paths("{C:/folder with spaces/a.xls} C:/b.xlsx")
        self.assertEqual(paths, [Path("C:/folder with spaces/a.xls"), Path("C:/b.xlsx")])
        with patch("app.messagebox.showwarning") as warning:
            result = self.window._handle_drop(SimpleNamespace(data="a.xls b.xlsx"), self.window.source_var, "source")
        self.assertEqual(result, "break")
        warning.assert_called_once()

    def test_drop_validates_extension_and_selects_req02_profile(self):
        target = Path(self.temp.name) / "02" / "B模板.xls"
        target.parent.mkdir()
        target.touch()
        with patch("app.messagebox.showwarning") as warning, patch("app.messagebox.askyesno", return_value=True) as confirm:
            self.window._handle_drop_paths([str(target)], self.window.target_var, "target")
            self.assertEqual(self.window.target_var.get(), str(target))
            self.assertEqual(self.window.profile_var.get(), "支线 02")
            self.window._handle_drop_paths([str(target.with_suffix(".txt"))], self.window.target_var, "target")
        confirm.assert_called_once()
        warning.assert_called_once()

    def test_drop_is_ignored_while_running(self):
        self.window.running = True
        with patch("app.messagebox.showwarning") as warning:
            self.window._handle_drop(SimpleNamespace(data="C:/new.xls"), self.window.source_var, "source")
        self.assertEqual(self.window.source_var.get(), "")
        warning.assert_not_called()

    def test_selecting_source_does_not_override_manual_profile(self):
        source = Path(self.temp.name) / "source.xls"
        source.touch()
        self.window.target_var.set(str(Path(self.temp.name) / "custom-target.xls"))
        self.window.profile_var.set("支线 02")
        with patch("app.filedialog.askopenfilename", return_value=str(source)):
            self.window._choose_source()
        self.assertEqual(self.window.profile_var.get(), "支线 02")

    def test_profile_pages_keep_independent_paths_and_active_aliases(self):
        self.window.source_var.set("main.xls")
        main_controls = self.window.file_controls
        self.window.profile_var.set("支线 02")
        self.assertEqual(self.window.source_var.get(), "")
        self.assertIs(self.window.source_var, self.window._profile_path_vars["支线 02"]["source"])
        self.assertIsNot(self.window.file_controls, main_controls)
        self.window.source_var.set("req.xls")
        self.window.profile_var.set("主线 01")
        self.assertEqual(self.window.source_var.get(), "main.xls")
        self.window.profile_var.set("支线 02")
        self.assertEqual(self.window.source_var.get(), "req.xls")

    def test_profile_pages_keep_independent_row_modes_and_button_selection(self):
        mainline_var = self.window.row_mode_var
        self.window.row_mode_var.set("All")
        self.assertTrue(self.window.row_mode_buttons[1].instate(["selected"]))

        self.window.profile_var.set("支线 02")
        self.assertIs(self.window.row_mode_var, self.window._profile_row_mode_vars["支线 02"])
        self.assertIsNot(self.window.row_mode_var, mainline_var)
        self.assertEqual(self.window.row_mode_var.get(), "Visible")
        self.assertTrue(self.window.row_mode_buttons[0].instate(["selected"]))

        self.window.profile_var.set("主线 01")
        self.assertIs(self.window.row_mode_var, mainline_var)
        self.assertEqual(self.window.row_mode_var.get(), "All")
        self.assertTrue(self.window.row_mode_buttons[1].instate(["selected"]))

    def test_run_mapping_uses_active_profiles_row_mode(self):
        self.window.profile_var.set("支线 02")
        self.window.row_mode_var.set("All")
        for name, variable in (("a.xls", self.window.source_var), ("b.xls", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        self.window.output_var.set(str(Path(self.temp.name) / "out.xls"))

        with patch("app.threading.Thread") as worker, patch.object(self.window.progress, "start"):
            self.window.run_mapping()

        args = worker.call_args.kwargs["args"]
        self.assertEqual(args[3], "All")
        self.assertEqual(args[4], "Req02")
        with patch("app.messagebox.showerror"):
            self.window._finish({"success": False, "message": "test cleanup"}, 1)

    def test_rejecting_target_page_switch_keeps_current_page(self):
        target = Path(self.temp.name) / "02" / "B模板.xls"
        target.parent.mkdir()
        target.touch()
        with patch("app.messagebox.askyesno", return_value=False):
            self.window._handle_drop_paths([str(target)], self.window.target_var, "target")
        self.assertEqual(self.window.profile_var.get(), "主线 01")
        self.assertEqual(self.window.target_var.get(), str(target))
        self.window.profile_var.set("支线 02")
        self.assertEqual(self.window.target_var.get(), "")

    def test_running_blocks_programmatic_page_switch(self):
        self.window.running = True
        self.window.profile_var.set("支线 02")
        self.assertEqual(self.window.profile_var.get(), "主线 01")
        self.assertEqual(self.window._active_profile_label, "主线 01")

    def test_worker_passes_selected_profile(self):
        completed = SimpleNamespace(stdout='{"success":true}', stderr="", returncode=0)
        with patch("app.subprocess.run", return_value=completed) as run, patch.object(self.window, "after"):
            self.window._run_worker(
                Path("a.xls"), Path("b.xls"), Path("out.xlsx"), "Visible", "Req02", "AppendExisting"
            )
        command = run.call_args.args[0]
        self.assertEqual(command[command.index("-Profile") + 1], "Req02")
        self.assertEqual(command[command.index("-OutputMode") + 1], "AppendExisting")

    def test_req03_target_switches_profile_and_defaults_to_xls_output(self):
        target = Path(self.temp.name) / "XPY沃尔玛价格计算.xls"
        target.touch()
        with patch("app.messagebox.askyesno", return_value=True) as confirm:
            self.window._handle_drop_paths([str(target)], self.window.target_var, "target")
        confirm.assert_called_once()
        self.assertEqual(self.window.profile_var.get(), "支线 03")
        self.assertEqual(self.window.target_var.get(), str(target))
        self.assertEqual(Path(self.window.output_var.get()).name, "XPY沃尔玛价格计算_已填充.xls")

    def test_req03_save_dialog_and_worker_use_req03_contract(self):
        self.window.profile_var.set("支线 03")
        target = Path(self.temp.name) / "renamed-target.xlsx"
        target.touch()
        self.window.target_var.set(str(target))
        with patch("app.filedialog.asksaveasfilename", return_value="") as choose:
            self.window._choose_output()
        self.assertEqual(choose.call_args.kwargs["defaultextension"], ".xls")
        self.assertEqual(choose.call_args.kwargs["filetypes"][0][1], "*.xls")
        self.assertNotIn("*.xlsx", str(choose.call_args.kwargs["filetypes"]))

        completed = SimpleNamespace(stdout='{"success":true}', stderr="", returncode=0)
        with patch("app.subprocess.run", return_value=completed) as run, patch.object(self.window, "after"):
            self.window._run_worker(
                Path("a.xls"), target, Path("out.xls"), "Visible", "Req03", "Replace"
            )
        command = run.call_args.args[0]
        self.assertEqual(command[command.index("-Profile") + 1], "Req03")
        self.assertEqual(command[command.index("-OutputMode") + 1], "Replace")

    def test_req03_rejects_non_xls_output_before_starting(self):
        self.window.profile_var.set("支线 03")
        for name, variable in (("a.xls", self.window.source_var), ("b.xls", self.window.target_var)):
            path = Path(self.temp.name) / name
            path.touch()
            variable.set(str(path))
        self.window.output_var.set(str(Path(self.temp.name) / "bad-output.xlsx"))

        with patch("app.messagebox.showwarning") as warning, patch("app.threading.Thread") as worker:
            self.window.run_mapping()

        warning.assert_called_once()
        self.assertIn(".xls", warning.call_args.args[1])
        worker.assert_not_called()
        self.assertFalse(self.window.running)

    def test_req03_rejects_non_xls_output_drop(self):
        self.window.profile_var.set("支线 03")
        output = Path(self.temp.name) / "bad-output.xlsm"
        with patch("app.messagebox.showwarning") as warning:
            result = self.window._handle_drop_paths(str(output), self.window.output_var, "output")
        self.assertEqual(result, "break")
        self.assertEqual(self.window.output_var.get(), "")
        warning.assert_called_once()


if __name__ == "__main__":
    unittest.main()

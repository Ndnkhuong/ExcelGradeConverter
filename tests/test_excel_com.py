"""Kiểm thử chuẩn hóa sheet đích bằng Microsoft Excel COM Automation (pywin32)."""

from __future__ import annotations

import builtins
import os
import sys
import tempfile
import threading
import time
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.core.errors import AppError
from app.core.excel_com import normalize_sheet_with_excel_com
from app.core.excel_writer import CellUpdate, write_updates_with_report


class _FakeCells:
    def __init__(self, log: list[tuple]) -> None:
        self._log = log
        self._number_format = None

    @property
    def NumberFormat(self):
        return self._number_format

    @NumberFormat.setter
    def NumberFormat(self, value: str) -> None:
        self._number_format = value
        self._log.append(("set_number_format", value))


class _FakeWorksheet:
    def __init__(self, name: str, log: list[tuple]) -> None:
        self.name = name
        self.Cells = _FakeCells(log)


class _FakeWorkbook:
    def __init__(self, path: str, sheet_names: list[str], log: list[tuple], fail_save: bool = False) -> None:
        self.path = path
        self._sheets = {name: _FakeWorksheet(name, log) for name in sheet_names}
        self._log = log
        self._fail_save = fail_save

    def Worksheets(self, name: str):
        self._log.append(("worksheets", name))
        if name not in self._sheets:
            raise RuntimeError(f"Unknown sheet {name}")
        return self._sheets[name]

    def Save(self) -> None:
        self._log.append(("save", self.path))
        if self._fail_save:
            raise RuntimeError("Disk full / locked")

    def Close(self, SaveChanges: bool = False) -> None:
        self._log.append(("close", SaveChanges))


class _FakeWorkbooks:
    def __init__(self, sheet_names: list[str], log: list[tuple], fail_open: bool = False, fail_save: bool = False):
        self._sheet_names = sheet_names
        self._log = log
        self._fail_open = fail_open
        self._fail_save = fail_save

    def Open(self, path: str, **kwargs):
        self._log.append(("open", path))
        if self._fail_open:
            raise RuntimeError("File locked by another user")
        return _FakeWorkbook(path, self._sheet_names, self._log, fail_save=self._fail_save)


class _FakeExcelApp:
    def __init__(self, sheet_names: list[str], log: list[tuple], fail_open: bool = False, fail_save: bool = False):
        self._log = log
        self.Visible = True
        self.DisplayAlerts = True
        self.AskToUpdateLinks = True
        self.EnableEvents = True
        self.ScreenUpdating = True
        self.Hwnd = 0
        self.Workbooks = _FakeWorkbooks(sheet_names, log, fail_open=fail_open, fail_save=fail_save)

    def Quit(self) -> None:
        self._log.append(("quit", None))


def _make_fake_com_modules(
    log: list[tuple],
    sheet_names: list[str] | None = None,
    dispatch_exc: Exception | None = None,
    fail_open: bool = False,
    fail_save: bool = False,
    hang_event: threading.Event | None = None,
) -> dict[str, types.ModuleType]:
    pythoncom_mod = types.ModuleType("pythoncom")
    pythoncom_mod.CoInitialize = lambda: log.append(("co_init", None))
    pythoncom_mod.CoUninitialize = lambda: log.append(("co_uninit", None))

    win32com_mod = types.ModuleType("win32com")
    client_mod = types.ModuleType("win32com.client")

    def _dispatch_ex(prog_id: str):
        log.append(("dispatch_ex", prog_id))
        if dispatch_exc is not None:
            raise dispatch_exc
        if hang_event is not None:
            hang_event.wait(timeout=5.0)
        return _FakeExcelApp(sheet_names or ["Sheet1"], log, fail_open=fail_open, fail_save=fail_save)

    client_mod.DispatchEx = _dispatch_ex
    win32com_mod.client = client_mod
    return {
        "pythoncom": pythoncom_mod,
        "win32com": win32com_mod,
        "win32com.client": client_mod,
    }


class TestExcelComAutomation(unittest.TestCase):
    def test_com_sequence_sets_general_saves_closes_and_releases(self) -> None:
        log: list[tuple] = []
        fake_mods = _make_fake_com_modules(log, sheet_names=["BangDiemB"])
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "temp.xlsx"
            wb = Workbook()
            wb.active.title = "BangDiemB"
            wb.save(path)

            with patch.dict(sys.modules, fake_mods):
                normalize_sheet_with_excel_com(path, "BangDiemB")

            self.assertEqual(
                log,
                [
                    ("co_init", None),
                    ("dispatch_ex", "Excel.Application"),
                    ("open", str(path.resolve())),
                    ("worksheets", "BangDiemB"),
                    ("set_number_format", "General"),
                    ("save", str(path.resolve())),
                    ("close", True),
                    ("quit", None),
                    ("co_uninit", None),
                ],
            )

    def test_missing_pywin32_raises_clear_error_and_does_not_create_output(self) -> None:
        real_import = builtins.__import__

        def _blocked_import(name, *args, **kwargs):
            if name in ("pythoncom", "win32com", "win32com.client"):
                raise ImportError(f"No module named {name}")
            return real_import(name, *args, **kwargs)

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            dest = tmp_dir / "dest.xlsx"
            output = tmp_dir / "out.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(["Họ và tên", "Điểm"])
            ws.append(["Nguyễn Văn A", None])
            wb.save(dest)

            with (
                patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}),
                patch.dict(sys.modules, {"pythoncom": None, "win32com": None, "win32com.client": None}),
                patch("builtins.__import__", side_effect=_blocked_import),
            ):
                with self.assertRaises(AppError) as ctx:
                    write_updates_with_report(dest, output, "Sheet1", [CellUpdate(2, 2, 9.0)])
                self.assertIn("pywin32", ctx.exception.message)
                self.assertFalse(output.exists())

    def test_excel_not_installed_raises_clear_error_and_does_not_create_output(self) -> None:
        log: list[tuple] = []
        fake_exc = RuntimeError(-2147221005, "Invalid class string")
        fake_mods = _make_fake_com_modules(log, dispatch_exc=fake_exc)
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            dest = tmp_dir / "dest.xlsx"
            output = tmp_dir / "out.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(["Họ và tên", "Điểm"])
            ws.append(["Nguyễn Văn A", None])
            wb.save(dest)

            with (
                patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}),
                patch.dict(sys.modules, fake_mods),
            ):
                with self.assertRaises(AppError) as ctx:
                    write_updates_with_report(dest, output, "Sheet1", [CellUpdate(2, 2, 9.0)])
                self.assertIn("Microsoft Excel", ctx.exception.message)
                self.assertFalse(output.exists())
                self.assertIn(("co_uninit", None), log)

    def test_save_failure_closes_without_saving_and_does_not_create_output(self) -> None:
        log: list[tuple] = []
        fake_mods = _make_fake_com_modules(log, sheet_names=["Sheet1"], fail_save=True)
        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            dest = tmp_dir / "dest.xlsx"
            output = tmp_dir / "out.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(["Họ và tên", "Điểm"])
            ws.append(["Nguyễn Văn A", None])
            wb.save(dest)

            with (
                patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}),
                patch.dict(sys.modules, fake_mods),
            ):
                with self.assertRaises(AppError):
                    write_updates_with_report(dest, output, "Sheet1", [CellUpdate(2, 2, 9.0)])
                self.assertFalse(output.exists())
                self.assertIn(("close", False), log)
                self.assertIn(("quit", None), log)
                self.assertIn(("co_uninit", None), log)

    def test_hung_excel_times_out_and_does_not_create_output(self) -> None:
        log: list[tuple] = []
        hang = threading.Event()
        fake_mods = _make_fake_com_modules(log, hang_event=hang)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "temp.xlsx"
            Workbook().save(path)
            try:
                with patch.dict(sys.modules, fake_mods):
                    start = time.monotonic()
                    with self.assertRaises(AppError) as ctx:
                        normalize_sheet_with_excel_com(path, "Sheet1", timeout_seconds=0.15)
                    elapsed = time.monotonic() - start
                    self.assertLess(elapsed, 2.0)
                    self.assertIn("không phản hồi", ctx.exception.message)
            finally:
                hang.set()

    def test_excel_com_is_final_writer_no_openpyxl_save_after_com(self) -> None:
        """Đảm bảo openpyxl không bao giờ save() lại file sau bước Excel COM."""
        events: list[str] = []
        orig_save = Workbook.save

        def _tracked_save(self, filename):
            events.append(f"openpyxl_save:{Path(filename).name}")
            return orig_save(self, filename)

        def _tracked_com(path, sheet_name):
            events.append(f"excel_com:{Path(path).name}:{sheet_name}")

        with tempfile.TemporaryDirectory() as tmp:
            tmp_dir = Path(tmp)
            dest = tmp_dir / "dest.xlsx"
            output = tmp_dir / "final_out.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "Sheet1"
            ws.append(["Họ và tên", "Điểm"])
            ws.append(["Nguyễn Văn A", None])
            wb.save(dest)
            events.clear()

            with (
                patch.dict(os.environ, {"APPDATA": str(tmp_dir / "appdata")}),
                patch.object(Workbook, "save", _tracked_save),
                patch("app.core.excel_writer.normalize_sheet_with_excel_com", side_effect=_tracked_com),
            ):
                write_updates_with_report(dest, output, "Sheet1", [CellUpdate(2, 2, 8.5)])

            self.assertEqual(len(events), 2)
            self.assertTrue(events[0].startswith("openpyxl_save:export_"))
            self.assertTrue(events[1].startswith("excel_com:export_"))
            self.assertTrue(output.exists())


if __name__ == "__main__":
    unittest.main()

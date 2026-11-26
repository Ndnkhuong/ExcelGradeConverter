"""Chuẩn hóa file Excel đích bằng Microsoft Excel COM Automation (pywin32).

Mở file tạm bằng ứng dụng Microsoft Excel thực tế, đặt NumberFormat = "General"
cho toàn bộ worksheet đích (tương đương Ctrl+A -> General), lưu bằng chính Excel
và giải phóng toàn bộ COM objects an toàn.
"""

from __future__ import annotations

import gc
import subprocess
import threading
from pathlib import Path

from app.core.errors import AppError
from app.utils.logger import get_logger

DEFAULT_COM_TIMEOUT_SECONDS = 60.0

# Mã HRESULT phổ biến khi máy chưa cài Microsoft Excel / COM class chưa đăng ký.
_EXCEL_NOT_INSTALLED_HRESULTS = {
    -2147221005,  # CO_E_CLASSSTRING (Invalid class string)
    -2147221164,  # REGDB_E_CLASSNOTREG (Class not registered)
}


def normalize_sheet_with_excel_com(
    file_path: str | Path,
    sheet_name: str,
    timeout_seconds: float = DEFAULT_COM_TIMEOUT_SECONDS,
) -> None:
    """Mở file bằng Microsoft Excel COM, đặt toàn bộ sheet đích về General và lưu lại.

    Nếu không cài pywin32, không cài Microsoft Excel, file đang bị khóa, Excel bị treo
    hoặc không thể lưu, hàm sẽ ném AppError với thông báo rõ ràng.
    """
    target_path = Path(file_path).resolve()
    if not target_path.exists():
        raise AppError("Không tìm thấy file tạm để chuẩn hóa bằng Microsoft Excel.")

    _ensure_pywin32_available()

    error_holder: list[Exception] = []
    pid_holder: list[int] = []

    worker = threading.Thread(
        target=_run_excel_com_worker,
        args=(target_path, sheet_name, error_holder, pid_holder),
        daemon=True,
    )
    worker.start()
    worker.join(timeout=timeout_seconds)

    if worker.is_alive():
        if pid_holder:
            _terminate_excel_process(pid_holder[0])
        get_logger().error(
            "Microsoft Excel COM bị treo quá %ss khi chuẩn hóa file %s",
            timeout_seconds,
            target_path.name,
        )
        raise AppError(
            "Microsoft Excel không phản hồi (bị treo) khi chuẩn hóa định dạng General. "
            "Quá trình xuất file đã bị hủy; hãy kiểm tra lại Microsoft Excel rồi thử lại."
        )

    if error_holder:
        exc = error_holder[0]
        if isinstance(exc, AppError):
            raise exc
        raise AppError(
            "Chuẩn hóa file bằng Microsoft Excel thất bại. File kết quả chưa được xuất."
        ) from exc


def _ensure_pywin32_available() -> None:
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError as exc:
        raise AppError(
            "Chưa cài đặt thư viện pywin32 để chuẩn hóa bằng Microsoft Excel. "
            "Hãy cài đặt bằng lệnh: pip install pywin32"
        ) from exc


def _run_excel_com_worker(
    target_path: Path,
    sheet_name: str,
    error_holder: list[Exception],
    pid_holder: list[int],
) -> None:
    import pythoncom
    import win32com.client

    com_initialized = False
    excel = None
    workbook = None
    worksheet = None
    saved = False

    try:
        pythoncom.CoInitialize()
        com_initialized = True

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise _classify_dispatch_error(exc) from exc

        pid = _get_excel_pid(excel)
        if pid is not None:
            pid_holder.append(pid)

        excel.Visible = False
        excel.DisplayAlerts = False
        for attr, value in (
            ("AskToUpdateLinks", False),
            ("EnableEvents", False),
            ("ScreenUpdating", False),
        ):
            try:
                setattr(excel, attr, value)
            except Exception:
                pass

        try:
            workbook = excel.Workbooks.Open(
                str(target_path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )
        except Exception as exc:
            get_logger().exception("Excel COM không mở được file tạm %s", target_path.name)
            raise AppError(
                "Microsoft Excel không thể mở file để chuẩn hóa. "
                "Hãy đóng các file Excel đang giữ khóa rồi thử lại."
            ) from exc

        try:
            worksheet = workbook.Worksheets(sheet_name)
        except Exception as exc:
            get_logger().exception("Excel COM không tìm thấy sheet %s", sheet_name)
            raise AppError(
                f"Microsoft Excel không tìm thấy sheet “{sheet_name}” khi chuẩn hóa định dạng."
            ) from exc

        try:
            worksheet.Cells.NumberFormat = "General"
        except Exception as exc:
            get_logger().exception("Excel COM lỗi khi đặt NumberFormat = General")
            raise AppError(
                "Microsoft Excel không thể đặt Number Format về General cho sheet đích."
            ) from exc

        try:
            workbook.Save()
            saved = True
        except Exception as exc:
            get_logger().exception("Excel COM không lưu được workbook %s", target_path.name)
            raise AppError(
                "Microsoft Excel không thể lưu file sau khi chuẩn hóa. "
                "Hãy kiểm tra file có đang bị khóa hoặc hết dung lượng ổ đĩa không."
            ) from exc

    except AppError as exc:
        error_holder.append(exc)
    except Exception as exc:
        get_logger().exception("Lỗi không xác định khi chạy Microsoft Excel COM Automation")
        error_holder.append(
            AppError(
                "Chuẩn hóa file bằng Microsoft Excel COM Automation thất bại. "
                "File kết quả chưa được xuất."
            )
        )
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=saved)
            except Exception:
                pass
        worksheet = None
        workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
            excel = None
        gc.collect()
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _classify_dispatch_error(exc: Exception) -> AppError:
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        first = exc.args[0]
        if isinstance(first, int):
            hresult = first
    text = str(exc).lower()
    if (
        hresult in _EXCEL_NOT_INSTALLED_HRESULTS
        or "invalid class string" in text
        or "class not registered" in text
    ):
        get_logger().error("Không tìm thấy Microsoft Excel COM (HRESULT=%s)", hresult)
        return AppError(
            "Không khởi chạy được Microsoft Excel. Máy tính cần cài đặt Microsoft Excel "
            "phiên bản desktop để chuẩn hóa và xuất file."
        )
    get_logger().exception("Lỗi khởi tạo Excel.Application COM")
    return AppError(
        "Không thể kết nối tới Microsoft Excel để chuẩn hóa file. "
        "Hãy kiểm tra cài đặt Microsoft Excel desktop trên máy rồi thử lại."
    )


def _get_excel_pid(excel: object) -> int | None:
    try:
        import win32process

        hwnd = getattr(excel, "Hwnd", None)
        if hwnd:
            _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
            return int(pid) if pid else None
    except Exception:
        return None
    return None


def _terminate_excel_process(pid: int) -> None:
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            capture_output=True,
            timeout=10,
            check=False,
        )
    except Exception:
        pass

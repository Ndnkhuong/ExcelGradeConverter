# Excel Grade Converter — Hướng dẫn cài đặt và chạy dự án

Ứng dụng desktop Windows chạy hoàn toàn local để chuyển dữ liệu điểm từ file Excel nguồn (mẫu A) sang file Excel đích (mẫu B) mà vẫn giữ nguyên cấu trúc, định dạng, thứ tự học sinh và không sửa file gốc.

- **Phiên bản:** `1.0.0`
- **Định dạng hỗ trợ:** `.xlsx`, `.xlsm` (không hỗ trợ `.xls`)
- **Công nghệ:** Python, PySide6, openpyxl, SQLite

---

## 1. Yêu cầu môi trường

- Hệ điều hành: **Windows 10 / 11**
- **Python 3.10 trở lên** (khuyến nghị **Python 3.12**) có kèm `pip`

Kiểm tra phiên bản Python trên máy:

```powershell
python --version
```

---

## 2. Cài đặt và chạy ứng dụng

Mở **PowerShell** hoặc **Command Prompt** tại thư mục gốc chứa dự án, sau đó làm theo các bước sau:

### Bước 1: Di chuyển vào thư mục dự án

```powershell
cd ExcelGradeConverter
```

### Bước 2: Tạo và kích hoạt môi trường ảo (khuyến nghị)

**Trên PowerShell:**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

> Nếu PowerShell báo lỗi chặn chạy script (`ExecutionPolicy`), chạy lệnh sau cho phiên hiện tại rồi kích hoạt lại:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> .\.venv\Scripts\Activate.ps1
> ```

**Trên Command Prompt (CMD):**

```bat
python -m venv .venv
.venv\Scripts\activate.bat
```

### Bước 3: Cài đặt thư viện phụ thuộc

Dự án sử dụng 2 thư viện ngoài: `PySide6` (giao diện) và `openpyxl` (đọc/ghi Excel).

```powershell
pip install --upgrade pip
pip install PySide6 openpyxl
```

### Bước 4: Chạy ứng dụng

```powershell
python main.py
```

Cửa sổ **Excel Grade Converter 1.0.0** sẽ mở lên.

---

## 3. Chạy nhanh cho các lần sau

Sau khi đã cài đặt xong môi trường ảo `.venv`, những lần sau chỉ cần chạy:

```powershell
cd ExcelGradeConverter
.\.venv\Scripts\Activate.ps1
python main.py
```

Hoặc chạy trực tiếp bằng Python trong `.venv` mà không cần kích hoạt:

```powershell
.\.venv\Scripts\python.exe main.py
```

---

## 4. Hướng dẫn sử dụng nhanh

Từ màn hình chính (**Dashboard**), nhấn **Bắt đầu chuyển đổi** (hoặc **Mở cấu hình đã lưu**) và đi qua 6 bước:

1. **Chọn file:** Chọn file Excel nguồn (mẫu A, mặc định dữ liệu từ **dòng 14**) và file Excel đích (mẫu B, mặc định dữ liệu từ **dòng 2**), kiểm tra sheet và dòng tiêu đề/dữ liệu.
2. **Ghép cột:** Chọn cột họ tên; hệ thống mặc định ghép **cột E (mẫu A) → cột E (mẫu B)** và bỏ qua các cột khác để không ghi đè ngoài phạm vi. Có thể lưu cấu hình mapping để dùng lại.
3. **Đối chiếu:** Ứng dụng tự động khớp học sinh duy nhất theo họ tên đã chuẩn hóa (giữ nguyên dấu tiếng Việt, không phân biệt hoa/thường). Các trường hợp trùng tên hoặc chưa khớp cần được ghép/xác nhận thủ công.
4. **Kiểm tra:** Rà soát các lỗi/cảnh báo (điểm không phải số, ngoài thang điểm, công thức chưa có giá trị tính sẵn, nguy cơ ghi đè ô có sẵn...).
5. **Xem trước:** Kiểm tra chi tiết các ô sẽ được cập nhật và các ô bị bỏ qua trước khi xuất.
6. **Xuất file:** Chọn thư mục lưu và tên file mới (mặc định `BangDiem_DaChuyenDoi_YYYYMMDD_HHMM.xlsx`). File nguồn và file đích gốc không bị thay đổi.

Ngoài ra, có thể vào **Cài đặt** để điều chỉnh thang điểm hợp lệ (mặc định `0–10`) hoặc vào **Lịch sử** để xem lại các lần chuyển đổi gần đây.

---

## 5. Dữ liệu local của ứng dụng

Ứng dụng chạy hoàn toàn offline và tự động tạo thư mục lưu cấu hình tại:

```text
%APPDATA%\ExcelGradeConverter\
├── database\app.db     # Lưu cấu hình mapping, lịch sử chuyển đổi và cài đặt
├── logs\app.log        # Nhật ký lỗi kỹ thuật (không ghi điểm hay họ tên học sinh)
├── temp\               # File tạm (tự động dọn khi mở/đóng ứng dụng)
└── settings.json       # Cài đặt thang điểm hợp lệ
```

---

## 6. Đóng gói thành file `.exe` (tùy chọn)

Nếu muốn đóng gói ứng dụng thành file thực thi Windows để chạy trên máy không cài Python:

```powershell
pip install pyinstaller
pyinstaller --noconfirm --windowed --name "ExcelGradeConverter" --add-data "app\resources\styles\app.qss;app\resources\styles" main.py
```

File thực thi sau khi build sẽ nằm trong thư mục `dist\ExcelGradeConverter\ExcelGradeConverter.exe`.

---

## 7. Xử lý sự cố thường gặp

| Lỗi / Tình huống | Cách xử lý |
| --- | --- |
| `ModuleNotFoundError: No module named 'PySide6'` hoặc `'openpyxl'` | Kiểm tra đã kích hoạt đúng `.venv` chưa và chạy lại `pip install PySide6 openpyxl`. |
| Không mở được file Excel (`File đang được chương trình khác giữ`) | Đóng file đó trong Microsoft Excel rồi thử lại. |
| Không chọn được file `.xls` | Ứng dụng chỉ hỗ trợ `.xlsx` và `.xlsm`. Hãy mở file `.xls` bằng Excel và **Save As** sang `.xlsx`. |
| Không tìm thấy cột hoặc lệch dòng dữ liệu | Kiểm tra lại mục **Dòng tiêu đề** và **Dòng bắt đầu dữ liệu** ở bước **1. Chọn file** (dòng dữ liệu phải nằm bên dưới dòng tiêu đề). |

**❤️ Đồng hành cùng dự án**
Công cụ được xây dựng với mong muốn hỗ trợ công việc hiệu quả và thuận tiện hơn.
Nếu bạn thấy công cụ mang lại giá trị thiết thực, hãy ủng hộ dự án để góp phần duy trì, cải tiến và phát triển thêm các tính năng mới.
Mỗi sự đóng góp đều là nguồn động viên quý báu. Xin cảm ơn!
<img width="380" height="374" alt="image" src="https://github.com/user-attachments/assets/d1324039-5044-457e-a214-846d66fcd821" />
1028665913 - Vietcombank


# Excel Grade Converter — Hướng dẫn cài đặt và chạy dự án

Ứng dụng desktop Windows chạy hoàn toàn local để chuyển dữ liệu điểm từ file Excel nguồn (mẫu A) sang file Excel đích (mẫu B) mà vẫn giữ nguyên cấu trúc, định dạng, thứ tự học sinh và không sửa file gốc.

- **Phiên bản:** `1.0.0`
- **Định dạng hỗ trợ:** `.xlsx`, `.xlsm` (không hỗ trợ `.xls`)
- **Công nghệ:** Python, PySide6, openpyxl, pywin32 (Microsoft Excel COM Automation), SQLite

---

## 1. Yêu cầu môi trường

- Hệ điều hành: **Windows 10 / 11**
- **Microsoft Excel phiên bản desktop** đã được cài đặt trên máy (bắt buộc cho bước chuẩn hóa toàn bộ sheet đích về `General` bằng Excel COM Automation khi xuất file)
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

Dự án sử dụng 3 thư viện ngoài: `PySide6` (giao diện), `openpyxl` (đọc/map Excel) và `pywin32` (chuẩn hóa & lưu file bằng Microsoft Excel COM Automation).

```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

Hoặc cài trực tiếp:

```powershell
pip install PySide6 openpyxl pywin32
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

## 4. Hướng dẫn sử dụng

Từ màn hình chính (**Dashboard**), nhấn **Bắt đầu chuyển đổi** (hoặc **Mở cấu hình đã lưu**) và đi qua 6 bước:

1. **Chọn file:**
   - Chọn file Excel nguồn (**Mẫu A**, mặc định dòng tiêu đề `13`, dữ liệu học sinh từ **dòng 14**) và file Excel đích (**Mẫu B**, mặc định dòng tiêu đề `1`, dữ liệu học sinh từ **dòng 2**).
   - Kiểm tra sheet và xem trước bảng dữ liệu (có đánh số dòng và tên cột Excel `A, B, C, D, E...`). Có thể chỉnh lại dòng tiêu đề / dòng bắt đầu nếu file khác mẫu mặc định.
2. **Ghép cột:**
   - Chọn cột **Họ tên nguồn** và **Họ tên đích** (ứng dụng tự nhận diện).
   - Hệ thống mặc định ghép **cột E (mẫu A) → cột E (mẫu B)** và bỏ qua các cột khác để tránh ghi đè ngoài phạm vi. Bạn có thể ghép thêm các cột khác (điểm số, mức đạt, nhận xét...) hoặc dùng nút **Gợi ý theo tên cột**.
   - Với mỗi cột được ghép, chọn **Loại dữ liệu** phù hợp:
     - **Điểm số:** tự động chuyển chuỗi số có dấu `.` hoặc `,` (ví dụ `"8.5"`, `"8,5"`, `"9"`) thành kiểu số thực/nguyên (`Number`) của Excel, không làm tròn và không đổi ô trống thành `0`.
     - **Đánh giá bằng chữ:** dùng cho cột mức đạt (`T`, `H`, `C`) hoặc nhận xét văn bản.
   - Tùy chọn **"Tự động chuyển mã đánh giá"** (mặc định **Bật**): khi bật, các ô trong cột **Đánh giá bằng chữ** chứa đúng 1 ký tự `T`, `H`, `C` (không phân biệt hoa/thường, bỏ khoảng trắng thừa) sẽ tự động chuyển thành:
     - `T` → `Hoàn thành tốt`
     - `H` → `Hoàn thành`
     - `C` → `Chưa hoàn thành`
     - Không chuyển các chuỗi nhiều ký tự (`TH`, `TC`, `Tốt`, `Hoàn thành`...), ô trống, hoặc cột **Điểm số**.
   - Có thể nhấn **Lưu cấu hình** / **Tải cấu hình** (lưu cả danh sách cột ghép, loại dữ liệu và trạng thái bật/tắt chuyển mã đánh giá).
3. **Đối chiếu:**
   - Ứng dụng tự động khớp học sinh duy nhất theo **họ tên đã chuẩn hóa** (Unicode NFC, bỏ khoảng trắng thừa, không phân biệt hoa/thường, giữ nguyên dấu tiếng Việt; **không** ghép theo thứ tự dòng).
   - Các trường hợp trùng tên hoặc chưa khớp phải được ghép/xác nhận thủ công.
4. **Kiểm tra:**
   - Rà soát lỗi nghiêm trọng (điểm số ngoài thang điểm cấu hình, thiếu cột bắt buộc...) và cảnh báo (học sinh chưa khớp, công thức chưa có giá trị, nguy cơ ghi đè, thiếu cột `Mã HCM` trên sheet đích...).
5. **Xem trước:**
   - Xem chính xác giá trị sẽ ghi (sau khi chuyển kiểu số và chuyển mã `T`/`H`/`C`), danh sách ô bỏ qua và cảnh báo.
   - Có thể bật/tắt trực tiếp checkbox **"Tự động chuyển mã đánh giá"** để xem kết quả thay đổi ngay trước khi xuất.
6. **Xuất file:**
   - Chọn thư mục lưu và tên file mới (mặc định `BangDiem_DaChuyenDoi_YYYYMMDD_HHMM.xlsx`). File nguồn và file đích gốc **không bao giờ bị sửa đổi**.
   - **Quy trình chuẩn hóa tương thích hệ thống import:**
     1. `openpyxl` ghi dữ liệu vào bản sao tạm, giữ nguyên thứ tự học sinh, định dạng (font, màu nền, viền, căn lề, độ rộng cột, chiều cao dòng, merge cells, công thức) và các sheet khác.
     2. **Microsoft Excel COM Automation (`pywin32`)** mở file tạm bằng Microsoft Excel thực tế, chuẩn hóa **Number Format toàn bộ sheet đích về `General`** (tương đương `Ctrl + A` → `General`, bao gồm tiêu đề, `Mã HCM`, họ tên, lớp, điểm, mức đánh giá, nhận xét...) và lưu trực tiếp bằng Excel.
     3. Cột `Mã HCM` và các mã định danh (`HSSSO...`, mã lớp, mã học sinh có số `0` ở đầu) được giữ nguyên tuyệt đối dưới dạng chuỗi văn bản.
     4. Xác minh lại file kết quả sau khi lưu (không lưu đè lại bằng `openpyxl`).

Ngoài ra, có thể vào **Cài đặt** để điều chỉnh thang điểm hợp lệ (mặc định `0–10`) hoặc vào **Lịch sử** để xem lại các lần chuyển đổi gần đây.

---

## 5. Chạy kiểm thử tự động (dành cho lập trình viên)

Trong thư mục `ExcelGradeConverter` (đã kích hoạt `.venv`), chạy toàn bộ unit test:

```powershell
python -m unittest discover -s tests -v
```

---

## 6. Dữ liệu local của ứng dụng

Ứng dụng chạy hoàn toàn offline và tự động tạo thư mục lưu cấu hình tại:

```text
%APPDATA%\ExcelGradeConverter\
├── database\app.db     # Lưu cấu hình mapping, lịch sử chuyển đổi và cài đặt
├── logs\app.log        # Nhật ký lỗi kỹ thuật (không ghi điểm hay họ tên học sinh)
├── temp\               # File tạm (tự động dọn khi mở/đóng ứng dụng)
└── settings.json       # Cài đặt thang điểm hợp lệ
```

---

## 7. Đóng gói thành file `.exe` (tùy chọn)

Nếu muốn đóng gói ứng dụng thành file thực thi Windows để chạy trên máy không cài Python (máy chạy `.exe` vẫn cần cài sẵn **Microsoft Excel desktop**):

```powershell
pip install pyinstaller
pyinstaller --noconfirm --windowed --name "ExcelGradeConverter" --add-data "app\resources\styles\app.qss;app\resources\styles" --hidden-import win32com.client --hidden-import pythoncom --hidden-import win32process main.py
```

File thực thi sau khi build sẽ nằm trong thư mục `dist\ExcelGradeConverter\ExcelGradeConverter.exe`.

---

## 8. Xử lý sự cố thường gặp

| Lỗi / Tình huống | Cách xử lý |
| --- | --- |
| `ModuleNotFoundError: No module named 'PySide6'`, `'openpyxl'` hoặc `'win32com'` | Kiểm tra đã kích hoạt đúng `.venv` chưa và chạy lại `pip install -r requirements.txt`. |
| `Không khởi chạy được Microsoft Excel` khi xuất file | Máy tính cần cài đặt **Microsoft Excel phiên bản desktop** (không hỗ trợ chỉ dùng Excel Online/WPS) để chạy Excel COM Automation. |
| `Microsoft Excel không phản hồi` khi xuất file | Đóng các hộp thoại hoặc file Excel đang bị treo trên máy, sau đó xuất lại. |
| Cảnh báo `Không tìm thấy cột "Mã HCM"` | Kiểm tra lại mục **Dòng tiêu đề** của file đích ở bước **1. Chọn file** nếu mẫu đích có cột `Mã HCM`. |
| Không mở được file Excel (`File đang được chương trình khác giữ`) | Đóng file đó trong Microsoft Excel rồi thử lại. |
| Không chọn được file `.xls` | Ứng dụng chỉ hỗ trợ `.xlsx` và `.xlsm`. Hãy mở file `.xls` bằng Excel và **Save As** sang `.xlsx`. |
| Không tìm thấy cột hoặc lệch dòng dữ liệu | Kiểm tra lại mục **Dòng tiêu đề** và **Dòng bắt đầu dữ liệu** ở bước **1. Chọn file** (dòng dữ liệu phải nằm bên dưới dòng tiêu đề). |

**❤️ Đồng hành cùng dự án**
Công cụ được xây dựng với mong muốn hỗ trợ công việc hiệu quả và thuận tiện hơn.
Nếu bạn thấy công cụ mang lại giá trị thiết thực, hãy ủng hộ dự án để góp phần duy trì, cải tiến và phát triển thêm các tính năng mới.
Mỗi sự đóng góp đều là nguồn động viên quý báu. Xin cảm ơn!
<img width="350" height="344" alt="image" src="https://github.com/user-attachments/assets/d1324039-5044-457e-a214-846d66fcd821" />
1028665913 - Vietcombank


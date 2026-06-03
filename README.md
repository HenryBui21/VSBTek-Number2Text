# VSBTek-Number2Text

**VSBTek-Number2Text** là một Excel Add-in chuyên nghiệp giúp chuyển đổi số thành chữ (đọc số) hỗ trợ cả tiếng Việt (VND) và tiếng Anh. Công cụ này được thiết kế để xử lý các con số cực lớn, đảm bảo độ chính xác cao trong các báo cáo tài chính và kế toán.

## 🌟 Tính năng nổi bật

- **Hỗ trợ số cực lớn:** Có khả năng đọc các số lên đến hàng tỷ tỷ (quintillions), vượt xa giới hạn thông thường của Excel.
- **Đa ngôn ngữ:** Hỗ trợ đọc số sang tiếng Việt (VND) và tiếng Anh (USD hoặc số thuần túy).
- **Độ chính xác cao:** Xử lý lỗi làm tròn và các trường hợp số thập phân phức tạp.
- **Hỗ trợ bảng mã:** Chuyển đổi linh hoạt giữa các bảng mã phổ biến tại Việt Nam (Unicode, TCVN3, VNI).
- **Cài đặt tự động:** Đi kèm script cài đặt thông minh, tự động cấu hình và đăng ký Add-in vào Excel chỉ với một cú click.

## 🚀 Hướng dẫn cài đặt

Để cài đặt Add-in, bạn chỉ cần thực hiện các bước sau:

1. Tải toàn bộ thư mục project về máy.
2. Chuột phải vào file `VSBTek-Installer.ps1` và chọn **Run with PowerShell**.
3. Chờ script hoàn tất quá trình build và đăng ký.
4. Mở Excel và bắt đầu sử dụng.

*Lưu ý: Nếu PowerShell hỏi về Execution Policy, hãy chọn 'Yes' để cho phép chạy script.*

## 📖 Cách sử dụng

Sau khi cài đặt thành công, bạn có thể sử dụng các hàm sau trực tiếp trong ô Excel:

### 1. Đọc số sang tiếng Việt (VND)
Cú pháp: `=VND(Số_tiền, [Kiểu_phông], [Sử_dụng_lẻ])`
- `Kiểu_phông`: 1 = Unicode (mặc định), 2 = VNI, 3 = TCVN3.
- `Sử_dụng_lẻ`: `TRUE` (mặc định - dùng "lẻ"), `FALSE` (dùng "linh").
- Ví dụ: `=VND(1250000)` -> *Một triệu hai trăm năm mươi nghìn đồng chẵn.*

### 2. Đọc số thuần túy (Tiếng Việt)
Cú pháp: `=DocSo(Số, [Loại_tiền], [Kiểu_phông], [Sử_dụng_lẻ])`
- `Loại_tiền`: 0 = Số thuần túy Việt Nam (mặc định), 1 = VND, 2 = USD, 3 = Tiếng Anh thường.
- Ví dụ: `=DocSo(15.5)` -> *Mười lăm phẩy năm.*

### 3. Đọc số sang tiếng Anh (USD)
Cú pháp: `=USD(Số_tiền)`
- Phân tách phần nguyên là Dollars và phần thập phân là Cents.
- Ví dụ: `=USD(123.45)` -> *One Hundred and Twenty-Three Dollars and Forty-Five Cents only.*

### 4. Đọc số thuần túy (Tiếng Anh)
Cú pháp: `=ENG(Số)`
- Đọc số thuần túy bằng tiếng Anh, hỗ trợ đọc chữ số lẻ thập phân từng số sau từ "point".
- Ví dụ: `=ENG(123.45)` -> *One Hundred and Twenty-Three point Four Five only.*


## 🛠 Cấu trúc thư mục

- `/src`: Chứa mã nguồn VBA (file `.bas`).
- `VSBTek-Installer.ps1`: Script PowerShell dùng để cài đặt và build Add-in.
- `VSBTek-Number2Text.xlam`: File Add-in đã được build sẵn (có thể dùng ngay).

## 👨‍💻 Phát triển và Tùy chỉnh

Nếu bạn muốn thay đổi logic đọc số, hãy chỉnh sửa các file trong thư mục `src`:
- `modNumber2TextCore.bas`: Logic xử lý chính cho tiếng Việt.
- `modNumber2TextEng.bas`: Logic xử lý tiếng Anh.
- `modFontConverter.bas`: Các hàm chuyển đổi bảng mã.
- `modPublicAPI.bas`: Định nghĩa các hàm công khai gọi từ Excel.

Sau khi sửa code, hãy chạy lại `VSBTek-Installer.ps1` để cập nhật file `.xlam`.

---
**Phát triển bởi VSBTek**
Email: contact@vsbtek.com
Website: [https://vsbtek.com](https://vsbtek.com)

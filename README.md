# Add-in Kiểm tra Độ Dốc Đường ống Revit

Add-in này được viết bằng C# cho Autodesk Revit, có chức năng kiểm tra độ dốc của đường ống thoát nước theo điều kiện đường kính.

## Tính năng

- Kiểm tra độ dốc đường ống dựa trên đường kính:
  - Đường kính ≥ Diameter_1 → Độ dốc phải bằng Slope_1
  - Đường kính ≤ Diameter_2 → Độ dốc phải bằng Slope_2
- Tự động loại trừ đường ống thẳng đứng khỏi việc kiểm tra
- Tạo bảng dự toán (schedule) cho các đường ống không đạt yêu cầu
- Giao diện người dùng thân thiện để nhập thông số

## Cài đặt

1. **Biên dịch project:**
   - Mở project trong Visual Studio
   - Chắc chắn rằng đã tham chiếu đến RevitAPI.dll và RevitAPIUI.dll từ thư mục cài đặt Revit
   - Biên dịch project để tạo file CheckSlopePipe.dll

2. **Cấu hình add-in:**
   - Sao chép file `CheckSlopePipe.addin` vào thư mục add-in của Revit:
     - `%APPDATA%\Autodesk\Revit\Addins\2024\` (cho Revit 2024)
   - Chỉnh sửa đường dẫn trong file `.addin` để trỏ đến vị trí file DLL

3. **Sử dụng trong Revit:**
   - Khởi động Revit
   - Mở dự án có đường ống cần kiểm tra
   - Vào tab Add-Ins → External Tools → "Kiểm tra độ dốc đường ống"

## Hướng dẫn sử dụng

1. **Chạy lệnh:** Trong Revit, chọn lệnh từ menu Add-Ins
2. **Nhập thông số:** 
   - Đường kính 1 (mm): Đường kính ngưỡng lớn
   - Độ dốc 1: Độ dốc yêu cầu cho đường kính lớn (ví dụ: 0.02 cho 2%)
   - Đường kính 2 (mm): Đường kính ngưỡng nhỏ
   - Độ dốc 2: Độ dốc yêu cầu cho đường kính nhỏ (ví dụ: 0.01 cho 1%)
3. **Xem kết quả:** Add-in sẽ:
   - Hiển thị số lượng đường ống đã kiểm tra và không đạt yêu cầu
   - Tạo schedule mới chứa thông tin các đường ống không đạt yêu cầu

## Thông số kỹ thuật

- **Ngôn ngữ:** C#
- **Framework:** .NET 4.8
- **Phiên bản Revit:** 2024 (có thể điều chỉnh cho phiên bản khác)
- **Phụ thuộc:** Revit API, Windows Forms

## Cấu trúc file

- `CheckSlopePipe.addin` - File cấu hình add-in
- `CheckSlopeCommand.cs` - Lớp chính chứa logic xử lý
- `CheckSlopePipe.csproj` - File project Visual Studio
- `README.md` - Hướng dẫn sử dụng

## Lưu ý quan trọng

- Add-in chỉ hoạt động với đường ống (Pipe) trong Revit
- Đường ống thẳng đứng được tự động loại trừ khỏi kiểm tra
- Đơn vị đường kính được chuyển đổi từ feet sang mm để so sánh
- Schedule được tạo tự động trong dự án Revit

## Xử lý lỗi thường gặp

- **Lỗi biên dịch:** Kiểm tra đường dẫn tham chiếu Revit API
- **Lỗi runtime:** Đảm bảo file DLL và add-in được đặt đúng vị trí
- **Không thấy lệnh:** Kiểm tra file .addin có đúng định dạng và vị trí

## Tùy chỉnh

Có thể điều chỉnh các hằng số trong code để thay đổi ngưỡng mặc định:
- `Diameter_1`, `Slope_1` - cho đường kính lớn
- `Diameter_2`, `Slope_2` - cho đường kính nhỏ

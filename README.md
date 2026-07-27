# KiotViet Printer Better

Ứng dụng Windows giúp đọc file Excel xuất từ KiotViet và in bằng mẫu BarTender `.btw`.

## Cơ chế file data cố định

- App đọc file KiotViet do người dùng chọn, không phụ thuộc tên file.
- Trước mỗi lượt in, app ghi sản phẩm hiện tại vào một file `.xls` trung gian có đường dẫn cố định.
- File trung gian chỉ có ba cột: `Tên hàng`, `Giá bán`, `Đơn vị tính`.
- Mẫu BarTender cần được kết nối với file trung gian này một lần. Những lần sau chỉ cần đưa file KiotViet mới vào app và in.
Không cần KiotViet Public API.

## Luồng sử dụng

1. Xuất danh sách hàng hóa từ KiotViet ra `.xls` hoặc `.xlsx`.
2. Kéo file vào app, hoặc bấm **Chọn file**.
3. Kiểm tra dữ liệu ngay trên màn hình.
4. Chọn một mẫu tem có trạng thái **Sẵn sàng**.
5. Xem trước và in.

## Quản lý mẫu tem

Mở **Mẫu tem & cài đặt** để:

- thêm hoặc bỏ một mẫu khỏi app;
- bật/tắt mẫu mà không xóa file thật;
- đổi file BarTender `.btw`;
- kiểm tra mẫu đã đủ điều kiện in hay chưa.

Thao tác **Bỏ khỏi app** chỉ xóa cấu hình. App không xóa file `.btw`, file dữ liệu hoặc
file Excel gốc của người dùng.

### Mẫu tem giá không dùng file data

App luôn truyền dữ liệu thẳng vào BarTender. Trong file `.btw`, hãy tạo ba
**Named Data Sources** đúng tên:

- `Tên hàng`
- `Giá bán`
- `Đơn vị tính`

Chế độ này không cần chọn file data trung gian.

## Sửa dữ liệu và zoom

- Các ô màu vàng có thể sửa trước khi xem trước hoặc in.
- Có thể sửa tên in trên tem, đơn vị tính, số lượng và giá bán.
- Chọn nhiều ô, dùng `Ctrl+C` / `Ctrl+V` để sao chép và dán như Excel.
- Dùng phím mũi tên để chuyển ô; `Enter` xác nhận nội dung và xuống dòng, `Shift+Enter` đi lên, `F2` sửa ô hiện tại.
- `Delete` xóa nội dung ô đã chọn, `Ctrl+Z` hoàn tác lần sửa gần nhất.
- `Ctrl+D` điền nội dung ô trên xuống vùng chọn, `Ctrl+R` điền nội dung ô bên trái sang vùng chọn.
- `Ctrl+S` lưu dữ liệu đã sửa thành file Excel mới.
- Có thể xóa các dòng đã chọn hoặc lưu dữ liệu đã sửa thành file `.xlsx` mới.
- Dùng thanh trượt, nút `−` / `+` hoặc giữ `Ctrl` và lăn chuột để zoom từ 50% đến 200%.
- Có thể dùng `Ctrl++`, `Ctrl+-` và `Ctrl+0` để tăng, giảm hoặc đặt lại zoom 100%.
- Zoom thay đổi cả chữ, chiều rộng cột, chiều cao dòng và tiêu đề; khi phóng lớn bảng có thanh cuộn ngang như Excel.
- App không ghi đè file Excel gốc.

## Lưu cấu hình

Cấu hình và bản sao mẫu `.btw` được lưu trong
`%LocalAppData%\KiotVietPrinterBetter`. Cài lại hoặc cập nhật app không làm mất mẫu.

## Yêu cầu khi chạy

- Windows 10/11;
- BarTender đã được cài đặt và license có hỗ trợ XML Script;
- driver máy in tem đã được cài;
- file `.btw` đã liên kết đúng với file dữ liệu tương ứng.

## Tải bản build từ GitHub

Vào tab **Actions** → workflow **Build Windows app** → chọn lần chạy mới nhất →
tải artifact `KiotViet-Printer-Better-win-x64`.

Workflow cũng tạo bộ cài `KiotViet-Printer-Better-Setup.exe`. Khi đẩy tag bắt đầu bằng
`v`, GitHub tự tạo Release gồm bản portable `.zip` và bộ cài `.exe`.

## Build trên máy lập trình

```powershell
dotnet restore "KiotViet Label Printer Pro V2.csproj"
dotnet build "KiotViet Label Printer Pro V2.csproj" -c Release
```

Project sử dụng .NET 9 WinForms.

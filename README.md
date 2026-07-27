# KiotViet Printer Better

Ứng dụng Windows giúp đọc file Excel xuất từ KiotViet và in bằng mẫu BarTender `.btw`.
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
- Dùng nút `−` / `+` hoặc giữ `Ctrl` và lăn chuột để zoom từ 70% đến 180%.
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

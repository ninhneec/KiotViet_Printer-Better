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
- đổi file dữ liệu `.xls`, `.xlsx` hoặc `.csv`;
- chọn cách xử lý `GENERIC`, `FULL`, `BARCODE` hoặc `GLASSES`;
- kiểm tra mẫu đã đủ điều kiện in hay chưa.

Thao tác **Bỏ khỏi app** chỉ xóa cấu hình. App không xóa file `.btw`, file dữ liệu hoặc
file Excel gốc của người dùng.

## Yêu cầu khi chạy

- Windows 10/11;
- BarTender đã được cài đặt và license có hỗ trợ XML Script;
- driver máy in tem đã được cài;
- file `.btw` đã liên kết đúng với file dữ liệu tương ứng.

## Tải bản build từ GitHub

Vào tab **Actions** → workflow **Build Windows app** → chọn lần chạy mới nhất →
tải artifact `KiotViet-Printer-Better-win-x64`.

## Build trên máy lập trình

```powershell
dotnet restore "KiotViet Label Printer Pro V2.csproj"
dotnet build "KiotViet Label Printer Pro V2.csproj" -c Release
```

Project sử dụng .NET 9 WinForms.

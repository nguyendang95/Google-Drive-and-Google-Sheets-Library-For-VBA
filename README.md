# Google Drive and Google Sheets Library For VBA
Thư viện này được viết bằng ngôn ngữ C# dựa trên thư viện Drive API .NET do Google cung cấp để người dùng có thể sử dụng được trong VBA.
Trang chủ: Get Started  |  API Client Library for .NET  |  Google Developers
Tác giả: Đăng Nguyễn
Email: nguyendang050695@gmail.com

Để sử dụng được thư viện này trong VBA, người dùng cần:
1. Xác định phiên bản Microsoft Office đang sử dụng trên máy tính
2. Mở trình nhắc lệnh Cmd của Windows (chọn Start | gõ “Cmd” trên thanh tìm kiếm, nhấp chuột phải vào Cmd và chọn Run As Administrator.
3. Chạy lệnh dưới đây, tùy thuộc vào phiên bản Microsoft Office đang sử dụng trên máy tính:
+ 32 bit: 
cd C:\Windows\Microsoft.NET\Framework\v4.0.30319
RegAsm %USERPROFILE%\Đường_dẫn\GoogleApis.dll /tlb:%USERPROFILE\Đường_dẫn\GoogleApis.tlb /codebase
+ 64 bit: 
cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
RegAsm %USERPROFILE%\Đường_dẫn\GoogleApis.dll /tlb:%USERPROFILE\Đường_dẫn\GoogleApis.tlb /codebase
Trong đó, Đường_dẫn là đường dẫn thư mục chứa hai tập tin GoogleApis.dll và GoogleApis.dll.
4. Mở cửa sổ soạn thảo code VBA trong ứng dụng Microsoft Office bất kỳ và tham chiếu đến thư viện GoogleApis

Tải xuống thư viện:
https://1drv.ms/u/s!AmwM8BaWQzt0nLwlwuroo0O46x7pkg?e=hq0hPl

Tài liệu tham khảo cách cài đặt và sử dụng:
https://1drv.ms/w/s!AmwM8BaWQzt0nLwfCrYcHvNHXGJ5lg?e=8lzGvM

Trong dự án này, người dùng có thể tìm thấy mã nguồn C# của thư viện để có thể dễ dàng tùy biến, chỉnh sửa, thêm bớt tính năng vào thư viện theo ý mình.

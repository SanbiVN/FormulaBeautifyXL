# FormulaBeautifyXL
Add-in phân tích công thức Excel thành dạng cây để dễ dàng chỉnh sửa
 
[TẢI XUỐNG](https://github.com/SanbiVN/FormulaBeautifyXL/releases/tag/formula_beautify_excel)
 
 
![excel formula beautify](https://user-images.githubusercontent.com/58664571/208216197-2dd3ec3d-9db3-4a6e-b316-7fa552a36d89.gif)


# Cách sử dụng rất đơn giản:


Ví dụ: ta có công thức ở ô A1 là:
=IF(SUM(IF("FOO"="BAR",10,0),10)=20,"FOO","BAR")
Khi nhấp hai lần chuột vào ô A1
Công thức sẽ được phân tích thành như ảnh bên trên, trả về ngay tại ô công thức đó.

Sau khi nhấn Enter hoặc Ctrl+Shift+Enter để thực thi công thức, nếu có cảnh báo như hình bên dưới thì bạn hãy gõ thêm Dấu Cách vào đầu công thức trước dấu "=", thành " =" và nhấn Enter hoặc Ctrl+Shift+Enter.

![alert_over_8192_characters](https://user-images.githubusercontent.com/58664571/208218159-cf8bca0e-c121-4924-9c48-7d92cacbd648.jpg)

### Các lệnh thực thi để đóng mở và cài đặt thông số:
(Chỉ cần gõ =Fx_ các hàm liên quan sẽ được gợi ý)

Thực thi | Gõ Hàm ô bất kỳ | Chú thích
------------- | ------------- | -------------
Mở form đầy đủ cài đặt | =Fx_ShowSettings() | 
Mở form chuyển đổi công thức | =Fx_ShowConverter() | 
Sao chép nhanh ô hiện tại | =Fx_Copy() | hoặc =Fx_Copy(B5:C10)
Sao chép và chuyển đổi ô hiện tại | =Fx_Copy_Reserve() | hoặc =Fx_Copy_Reserve(B5:C10)
Sao chép (thành cú pháp Google Sheet) | =Fx_Copy_forExcelOnline() | hoặc =Fx_Copy_forExcelOnline(B5:C10)
Sao chép (thành cú pháp Google Sheet) | =Fx_Copy_forGoolgeSheet() | hoặc =Fx_Copy_forGoolgeSheet(B5:C10)
Dán công thức (tự động chuyển đổi) | =Fx_Paste | 
Chuyển đổi công thức | =Fx_ChooseDirectories() | 
Khôi phục cài đặt mặc định | =Fx_ResetSettings() | 
Bật tự động kiểm tra cập nhật | =Fx_UpdateCheck_On() | 
Tắt tự động kiểm tra cập nhật | =Fx_UpdateCheck_Off() | 

### Ví dụ sao chép chuyển dấu tách đối số:

![formula beautify copy convert](https://user-images.githubusercontent.com/58664571/210694850-48d5b562-94e2-481b-b17e-09d17345a958.gif)


# Phím tắt:
1. Ctrl+Shift+B: Để phân tích hoặc phục hồi công thức trong ô (cần chọn vùng ô)
2. Ctrl+Alt+C dùng để Cho phép quá trình Phân tích công thức hoạt động hay tạm dừng
(Phím tắt mặc định là Ctrl+Alt+C, bạn có thể sửa thành phím tắt mong muốn của bạn, ^ là Ctrl, + là Shift và % là Alt)


# Cài đặt: 
cài đặt Add-in trong Tab Developer hoặc sao chép vào thư mục khởi động XLStart của ứng dụng Excel.

Đăng góp ý và câu họi tại diễn đàn Giải Pháp Excel: [Giải Pháp Excel](https://www.giaiphapexcel.com/diendan/threads/159912/)

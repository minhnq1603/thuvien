# Kiểm tra Product key
Mở Command Prompt với quyền administrator và di chuyển tới thư mục chứa Microsoft Office

Office 2013:

64-bit: `cd C:\Program Files\Microsoft Office\Office15`

32-bit: `cd C:\Program Files (x86)\Microsoft Office\Office15`

Office 2016/2019/2021:

64-bit: `C:\Program Files\Microsoft Office\Office16`

32-bit: `C:\Program Files (x86)\Microsoft Office\Office16`

Sau đó thực hiện lệnh `cscript ospp.vbs /dstatus` để lấy thông tin product key đã cài đặt. Để ký dòng `Last 5 characters of installed product key` chính là thông tin cần lấy

![Go-bo-product-key-cua-Microsoft-Office-01](https://s3-hcm-r1.longvan.net/thuvien/shared/1223/Go-bo-product-key-cua-Microsoft-Office-01.png)
# Gỡ bỏ product key
Để thực hiện gỡ bỏ product key thì thực hiện lệnh `cscript ospp.vbs /unpkey:{**5 ký tự cuối vừa lấy ở trên**}` và chờ một lát cho đến khi có thông báo `Product key uninstall successful`
![Go-bo-product-key-cua-Microsoft-Office-02](https://s3-hcm-r1.longvan.net/thuvien/shared/1223/Go-bo-product-key-cua-Microsoft-Office-02.png)
Thực hiện lại lệnh `cscript ospp.vbs /dstatus` để kiểm tra lại thông tin product key, nếu xuất hiện thông báo `No installed product keys detected` nghĩa là product key đã được gỡ bỏ hết
![Go-bo-product-key-cua-Microsoft-Office-03](https://s3-hcm-r1.longvan.net/thuvien/shared/1223/Go-bo-product-key-cua-Microsoft-Office-03.png)
Mở một ứng dụng bất kỳ trong bộ Microsoft Office để kiểm tra thông tin giấy phép
![Go-bo-product-key-cua-Microsoft-Office-04](https://s3-hcm-r1.longvan.net/thuvien/shared/1223/Go-bo-product-key-cua-Microsoft-Office-04.png)

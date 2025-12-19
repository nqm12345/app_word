# HƯỚNG DẪN CÀI ĐẶT VÀ SỬ DỤNG CHINHSUAOFFICE

## 📋 YÊU CẦU

- **Windows**: từ **Windows 7 SP1** đến **Windows 11**.
- **Microsoft Office**: **Office 2013, 2016, 2019, 2021, Microsoft 365**.
- Chủ yếu sử dụng với **Microsoft Word**, hỗ trợ mở link `ms-word:` trực tiếp từ website.

---

## 🚀 CÀI ĐẶT VÀ CHẠY ỨNG DỤNG

### 1. Giải nén bộ cài
- Tải file nén chứa bộ cài `ChinhSuaOffice`.
- Chuột phải vào file `.zip` → chọn **Extract All… / Giải nén**.
- Vào thư mục đã giải nén, tìm file `ChinhSuaOffice_Setup.exe`.

_Chèn ảnh: Màn hình giải nén file cài đặt (Hình 1)_

### 2. Cài đặt ứng dụng
1. Chuột phải vào `ChinhSuaOffice_Setup.exe` → chọn **Run as Administrator**.
2. Bấm **Next / Tiếp tục** theo hướng dẫn cho đến khi hoàn tất.
3. Có thể tick chọn **Tạo icon ngoài Desktop** và **Khởi động cùng Windows** (nếu muốn).
4. Sau khi cài xong, ứng dụng sẽ tự chạy nền, không cần mở tay.

_Chèn ảnh: Chuột phải chọn "Run as Administrator" (Hình 2)_  
_Chèn ảnh: Màn hình các bước Next / chọn tùy chọn cài đặt (Hình 3)_

### 3. Ứng dụng sẽ làm gì khi chạy?
- Tự cấu hình một số **Registry của Office** để Word cho phép mở file từ web qua `ms-word:`.
- Thêm các địa chỉ server nội bộ vào **Trusted Locations** (vị trí tin cậy).
- Khởi động một **dịch vụ nền (WebDAV server)** để trung gian tải/lưu file với server.
- Hiển thị **icon ở System Tray** (góc phải dưới màn hình, gần đồng hồ) và chạy ẩn.

_Chèn ảnh: Icon ứng dụng ở System Tray và menu chuột phải (Hình 4)_

---

## 🖱️ CÁCH SỬ DỤNG VỚI WORD

### 1. Mở tài liệu từ website
1. Mở trình duyệt, đăng nhập vào **website nội bộ** có tài liệu.
2. Tìm tài liệu cần chỉnh sửa, nhấn nút **“Chỉnh sửa”**.

_Chèn ảnh: Nút "Chỉnh sửa" trên giao diện website (Hình 5)_

3. Nếu trình duyệt hiển thị hộp thoại hỏi cho phép mở ứng dụng / truy cập mạng cục bộ → chọn **Cho phép / Allow**.

_Chèn ảnh: Popup hỏi quyền truy cập, chọn Cho phép (Hình 6)_

### 2. Chỉnh sửa và lưu trong Word
4. **Microsoft Word** sẽ tự động mở tài liệu từ server.
5. Chỉnh sửa nội dung như bình thường (soạn thảo, chèn bảng, hình ảnh,...).
6. Nhấn **Ctrl+S** để lưu → file sẽ được gửi ngược lên server thông qua ứng dụng ChinhSuaOffice.

_Chèn ảnh: Word đang mở tài liệu từ server (Hình 7)_  
_Chèn ảnh: Nhấn Ctrl+S để lưu tài liệu (Hình 8)_

---

## 🔧 GỠ BỎ ỨNG DỤNG (NẾU KHÔNG DÙNG NỮA)

1. Mở **Control Panel → Programs and Features**.
2. Tìm **ChinhSuaOffice** trong danh sách chương trình.
3. Nhấn **Uninstall** để gỡ cài đặt.

_Chèn ảnh: Màn hình Programs and Features với mục ChinhSuaOffice (Hình 9)_

---


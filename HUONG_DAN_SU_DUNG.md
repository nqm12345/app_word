# 📘 HƯỚNG DẪN CÀI ĐẶT VÀ SỬ DỤNG
## Ứng dụng Chỉnh sửa Office (ChinhSuaWord)

<!-- [CHÈN ẢNH: Logo công ty hoặc banner ứng dụng] -->

---

## 📋 Giới thiệu

**ChinhSuaWord** là ứng dụng cho phép chỉnh sửa file Office (Word, Excel, PowerPoint, Visio) **trực tiếp từ hệ thống web** mà không cần tải về máy rồi upload lại.

### Cách hoạt động:

```
┌─────────────┐      ┌─────────────┐      ┌─────────────┐      ┌─────────────┐
│   Website   │ ──▶  │  App chạy   │ ──▶  │   Office    │ ──▶  │   Server    │
│  Click Edit │      │  trên máy   │      │  Word/Excel │      │  Lưu file   │
└─────────────┘      └─────────────┘      └─────────────┘      └─────────────┘
```

<!-- [CHÈN ẢNH: Sơ đồ luồng hoạt động - Web → App → Office → Lưu] -->

### Tính năng:
- ✅ Mở và chỉnh sửa file Word (.doc, .docx)
- ✅ Mở và chỉnh sửa file Excel (.xls, .xlsx)
- ✅ Mở và chỉnh sửa file PowerPoint (.ppt, .pptx)
- ✅ Mở và chỉnh sửa file Visio (.vsd, .vsdx)
- ✅ Tự động cấu hình Registry tắt Protected View
- ✅ Tự động thêm server vào Trusted Locations
- ✅ Lưu file trực tiếp lên server (Ctrl+S)

---

## 📦 Yêu cầu hệ thống

| Yêu cầu | Chi tiết |
|---------|----------|
| **Hệ điều hành** | Windows 10 / Windows 11 |
| **Microsoft Office** | 2013, 2016, 2019, 2021 hoặc Microsoft 365 |
| **Kết nối mạng** | Có kết nối Internet để truy cập server |
| **Quyền** | Không cần quyền Admin |

---

## 📁 Danh sách file cài đặt

Sau khi giải nén, bạn sẽ thấy các file sau:

| File | Mô tả | Bắt buộc |
|------|-------|----------|
| `ChinhSuaWord.exe` | Ứng dụng chính (159 MB) | ✅ Có |
| `config.json` | File cấu hình server | ✅ Có |
| `app.ico` | Icon ứng dụng | ❌ Không |
| `HUONG_DAN_SU_DUNG.md` | Hướng dẫn này | ❌ Không |

<!-- [CHÈN ẢNH: Screenshot folder chứa các file trong Windows Explorer] -->

---

## 🚀 HƯỚNG DẪN CÀI ĐẶT CHI TIẾT

### Bước 1: Giải nén file

1. Nhận file `ChinhSuaOffice.zip` hoặc `ChinhSuaOffice.rar` từ quản trị viên
2. Click chuột phải vào file → Chọn **"Extract Here"** hoặc **"Giải nén tại đây"**
3. Một folder mới sẽ được tạo chứa các file ứng dụng

**Gợi ý vị trí lưu:**
- `C:\ChinhSuaOffice\`
- `D:\ChinhSuaOffice\`
- Desktop (để dễ truy cập)

<!-- [CHÈN ẢNH: Screenshot click chuột phải → Extract Here] -->

---

### Bước 2: Kiểm tra cấu hình (Quan trọng!)

1. Mở folder vừa giải nén
2. Tìm file `config.json`
3. Click chuột phải → **Open with** → **Notepad**
4. Kiểm tra nội dung:

```json
{
  "port": 1901,
  "companyApiUrl": "https://administrator.lifetex.vn:316",
  "apiEndpoint": "/api/files/download"
}
```

**Giải thích các thông số:**

| Thông số | Ý nghĩa | Có cần sửa? |
|----------|---------|-------------|
| `port` | Cổng ứng dụng chạy | ❌ Không (giữ 1901) |
| `companyApiUrl` | Địa chỉ server công ty | ⚠️ Kiểm tra đúng chưa |
| `apiEndpoint` | Đường dẫn API | ❌ Không |

> ⚠️ **Lưu ý:** Nếu `companyApiUrl` sai, ứng dụng sẽ không hoạt động. Liên hệ quản trị viên để lấy URL đúng.

<!-- [CHÈN ẢNH: Screenshot file config.json mở trong Notepad với các thông số] -->

---

### Bước 3: Chạy ứng dụng lần đầu

1. Quay lại folder chứa file
2. **Double-click** vào file `ChinhSuaWord.exe`
3. Nếu Windows hỏi **"Windows protected your PC"**:
   - Click **"More info"**
   - Click **"Run anyway"**

<!-- [CHÈN ẢNH: Screenshot Windows SmartScreen với nút "Run anyway"] -->

**Khi khởi động lần đầu, ứng dụng sẽ TỰ ĐỘNG:**

| Bước | Mô tả | Thời gian |
|------|-------|-----------|
| 1 | Cấu hình Registry tắt Protected View | 1-2 giây |
| 2 | Thêm server vào Trusted Locations | 1-2 giây |
| 3 | Tạo shortcut trên Desktop | 1 giây |
| 4 | Khởi động WebDAV server | 1-2 giây |

---

### Bước 4: Xác nhận ứng dụng đã sẵn sàng

Khi thấy giao diện như sau, ứng dụng đã sẵn sàng:

```
┌────────────────────────────────────────────────┐
│  Trình chỉnh sửa Word  •  Đang chạy           │
├────────────────────────────────────────────────┤
│  PORT          STATUS          API             │
│  1901          RUNNING         https://...     │
│                (màu xanh)                      │
├────────────────────────────────────────────────┤
│  CONSOLE OUTPUT                                │
│  ✅ Đã cấu hình Registry cho 4 ứng dụng Office │
│  ✅ Đã sẵn sàng trên cổng 1901                 │
│  🌐 API: https://administrator.lifetex.vn:316 │
└────────────────────────────────────────────────┘
```

<!-- [CHÈN ẢNH: Screenshot giao diện app đang chạy với STATUS = RUNNING màu xanh] -->

**Kiểm tra thành công:**
- ✅ Status hiển thị **RUNNING** (màu xanh)
- ✅ Log hiển thị **"Đã sẵn sàng trên cổng 1901"**
- ✅ Không có lỗi màu đỏ

---

### Bước 5: Giữ ứng dụng chạy nền

> ⚠️ **QUAN TRỌNG:** Ứng dụng phải **LUÔN CHẠY** khi bạn muốn chỉnh sửa file từ web!

**Cách để ứng dụng chạy nền:**

1. **Thu nhỏ** (click nút `-`) → App thu xuống taskbar
2. **KHÔNG đóng** (không click nút `X`)
3. Để app chạy suốt ngày làm việc

**Mẹo:** Sau khi cài đặt, shortcut **"Chỉnh sửa Office"** sẽ xuất hiện trên Desktop. Lần sau chỉ cần double-click shortcut này.

<!-- [CHÈN ẢNH: Screenshot shortcut "Chỉnh sửa Office" trên Desktop] -->

---

## 💻 HƯỚNG DẪN SỬ DỤNG CHI TIẾT

### Quy trình chỉnh sửa file

```
Bước 1          Bước 2          Bước 3          Bước 4          Bước 5
┌─────┐        ┌─────┐        ┌─────┐        ┌─────┐        ┌─────┐
│ Web │   ──▶  │Click│   ──▶  │File │   ──▶  │Sửa  │   ──▶  │Ctrl │
│     │        │Edit │        │ mở  │        │file │        │ +S  │
└─────┘        └─────┘        └─────┘        └─────┘        └─────┘
Mở website    Click nút     Office mở     Chỉnh sửa      Lưu lên
              Chỉnh sửa     file lên      nội dung       server
```

---

### Bước 1: Mở website và tìm file

1. Mở trình duyệt (Chrome, Edge, Firefox...)
2. Truy cập hệ thống quản lý văn bản của công ty
3. Đăng nhập tài khoản
4. Tìm đến file cần chỉnh sửa

<!-- [CHÈN ẢNH: Screenshot trang web hiển thị danh sách file] -->

---

### Bước 2: Click nút "Chỉnh sửa"

1. Tìm nút **"Chỉnh sửa"**, **"Edit"** hoặc icon bút chì ✏️
2. Click vào nút đó
3. Trình duyệt sẽ hỏi **"Open with..."** → Chọn **OK** hoặc **Allow**

<!-- [CHÈN ẢNH: Screenshot nút "Chỉnh sửa" được highlight trên web] -->

<!-- [CHÈN ẢNH: Screenshot popup "Open with..." của trình duyệt] -->

---

### Bước 3: File tự động mở trong Office

Sau khi click, Office tương ứng sẽ tự động mở:

| Loại file | Ứng dụng mở |
|-----------|-------------|
| .doc, .docx | Microsoft Word |
| .xls, .xlsx | Microsoft Excel |
| .ppt, .pptx | Microsoft PowerPoint |
| .vsd, .vsdx | Microsoft Visio |

**Thời gian chờ:** 3-10 giây (tùy tốc độ mạng và kích thước file)

<!-- [CHÈN ẢNH: Screenshot file Word đang mở với nội dung từ server] -->

---

### Bước 4: Chỉnh sửa nội dung

1. File đã mở → Chỉnh sửa như bình thường
2. Thêm, xóa, sửa nội dung tùy ý
3. Định dạng văn bản, thêm bảng, hình ảnh...

> 💡 **Mẹo:** Làm việc như với file bình thường trên máy tính!

---

### Bước 5: Lưu file lên server

**Cách 1: Phím tắt (Khuyến nghị)**
- Nhấn **Ctrl + S**

**Cách 2: Menu**
- File → Save

**Sau khi lưu:**
- Thanh tiêu đề không còn dấu `*` (dấu sao biểu thị chưa lưu)
- File đã được cập nhật lên server

<!-- [CHÈN ẢNH: Screenshot nhấn Ctrl+S, thanh tiêu đề không còn dấu *] -->

---

### Bước 6: Đóng file

1. Sau khi lưu xong, đóng file: **File → Close** hoặc click **X**
2. Nếu Office hỏi **"Save changes?"** → Click **Save** để chắc chắn
3. Quay lại web để kiểm tra file đã cập nhật

---

## ⚠️ CÁC LƯU Ý QUAN TRỌNG

### ✅ NÊN làm:

| Nên | Lý do |
|-----|-------|
| Giữ app chạy suốt ngày | Để mở file bất cứ lúc nào |
| Lưu thường xuyên (Ctrl+S) | Tránh mất dữ liệu |
| Đóng file khi xong | Giải phóng tài nguyên |
| Kiểm tra mạng trước khi lưu | Đảm bảo lưu thành công |

### ❌ KHÔNG NÊN làm:

| Không nên | Hậu quả |
|-----------|---------|
| Tắt app khi đang sửa file | Mất kết nối, không lưu được |
| Đổi tên file khi đang mở | Lỗi khi lưu |
| Mở cùng 1 file trên 2 máy | Xung đột dữ liệu |
| Chỉnh sửa offline | Không lưu được lên server |

---

## 🔧 XỬ LÝ SỰ CỐ CHI TIẾT

### Sự cố 1: File mở ở chế độ Protected View

**Triệu chứng:**
- Thanh vàng hiện ở trên cùng: **"PROTECTED VIEW - Be careful..."**
- Không thể chỉnh sửa file

<!-- [CHÈN ẢNH: Screenshot thanh vàng Protected View trong Excel] -->

**Nguyên nhân:** 
- Chạy app lần đầu nhưng Office đã mở sẵn
- Registry chưa được cấu hình

**Cách xử lý:**

**Cách 1: Khởi động lại (Đơn giản)**
1. Đóng TẤT CẢ file Word/Excel/PowerPoint
2. Tắt app ChinhSuaWord
3. Mở lại app ChinhSuaWord
4. Mở lại file từ web

**Cách 2: Cấu hình thủ công (Nếu cách 1 không được)**

Làm theo các bước sau trong Word/Excel/PowerPoint:

```
1. Mở Word (hoặc Excel/PowerPoint)
2. File → Options (Tùy chọn)
3. Trust Center (Trung tâm Tin cậy) → Trust Center Settings
4. Trusted Locations (Vị trí Tin cậy)
5. Click "Add new location..." (Thêm vị trí mới)
6. Nhập: https://administrator.lifetex.vn:316
7. ✅ Tick "Subfolders of this location are also trusted"
8. Click OK → OK
```

<!-- [CHÈN ẢNH: Screenshot Trust Center với Trusted Locations] -->

<!-- [CHÈN ẢNH: Screenshot dialog "Add new location" với URL đã nhập] -->

---

### Sự cố 2: Không mở được file từ web

**Triệu chứng:**
- Click "Chỉnh sửa" nhưng không có gì xảy ra
- Hoặc báo lỗi "Cannot open file"

**Kiểm tra:**

| Kiểm tra | Cách kiểm tra |
|----------|---------------|
| App đang chạy? | Xem taskbar có icon app không |
| Status = RUNNING? | Mở app, xem status màu xanh chưa |
| Có mạng? | Thử mở website khác |
| URL đúng? | Kiểm tra config.json |

**Cách xử lý:**
1. Mở app ChinhSuaWord (nếu chưa mở)
2. Chờ status = RUNNING
3. Thử lại click "Chỉnh sửa" trên web

---

### Sự cố 3: Lưu file bị lỗi

**Triệu chứng:**
- Nhấn Ctrl+S nhưng báo lỗi
- Hoặc hiện thông báo "Upload failed"

**Nguyên nhân có thể:**
- Mất kết nối mạng
- Server đang bảo trì
- Phiên đăng nhập hết hạn

**Cách xử lý:**

1. **Lưu tạm ra máy:**
   - File → Save As → Chọn Desktop
   - Đặt tên khác để không nhầm

2. **Kiểm tra mạng:**
   - Thử mở website công ty
   - Nếu không mở được → Đợi mạng ổn định

3. **Đăng nhập lại:**
   - Mở web, đăng xuất rồi đăng nhập lại
   - Thử mở và lưu file lại

4. **Upload thủ công:**
   - Nếu vẫn lỗi, upload file đã lưu ở bước 1 lên web

---

### Sự cố 4: Port 1901 đã được sử dụng

**Triệu chứng:**
- App báo lỗi: "Port 1901 is already in use"
- Status hiện màu đỏ

**Cách xử lý:**

1. Mở file `config.json` bằng Notepad
2. Đổi `"port": 1901` thành `"port": 1902`
3. Lưu file (Ctrl+S)
4. Khởi động lại app

```json
{
  "port": 1902,  ← Đổi số này
  "companyApiUrl": "https://administrator.lifetex.vn:316",
  "apiEndpoint": "/api/files/download"
}
```

---

## ❓ CÂU HỎI THƯỜNG GẶP (FAQ)

### Q1: App có cần chạy liên tục không?
**A:** Có, app phải chạy khi bạn muốn mở/lưu file từ web. Có thể thu nhỏ xuống taskbar.

### Q2: Có thể cài trên nhiều máy không?
**A:** Có, mỗi máy cần cài riêng.

### Q3: Mất mạng giữa chừng thì sao?
**A:** File vẫn mở được, nhưng không lưu lên server được. Hãy lưu tạm ra máy (Save As).

### Q4: Có cần quyền Admin không?
**A:** Không, app chỉ ghi vào Registry của user hiện tại (HKCU).

### Q5: Có thể sử dụng khi ở nhà không?
**A:** Có, nếu server công ty cho phép truy cập từ internet. Liên hệ IT để biết thêm.

### Q6: App có tự động cập nhật không?
**A:** Không, khi có phiên bản mới sẽ được thông báo và gửi file cài đặt mới.

---

## 📞 HỖ TRỢ KỸ THUẬT

Nếu gặp vấn đề không thể tự xử lý, vui lòng liên hệ:

| Kênh | Thông tin |
|------|-----------|
| **Email** | support@lifetex.vn |
| **Hotline** | 1900-xxxx |
| **Website** | https://lifetex.vn |

**Khi liên hệ, vui lòng cung cấp:**
- Screenshot lỗi (nếu có)
- Nội dung log trong app
- Phiên bản Windows và Office đang dùng

---

## 📝 LỊCH SỬ PHIÊN BẢN

| Phiên bản | Ngày | Thay đổi |
|-----------|------|----------|
| 1.0.0 | 02/12/2024 | Phát hành đầu tiên |

---

**© 2024 LifeTex Company. All rights reserved.**

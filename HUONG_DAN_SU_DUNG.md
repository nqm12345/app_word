# 📘 HƯỚNG DẪN SỬ DỤNG ỨNG DỤNG CHỈNH SỬA OFFICE

**Phiên bản:** 1.0.0  
**Công ty:** LifeTex

---

## � YÊU CẦU HỆ THỐNG

### Windows:
| Phiên bản | Hỗ trợ |
|-----------|--------|
| Windows 11 | ✅ Có |
| Windows 10 | ✅ Có |

### Microsoft Office:
| Phiên bản | Hỗ trợ |
|-----------|--------|
| Microsoft 365 | ✅ Có |
| Office 2021 | ✅ Có |
| Office 2019 | ✅ Có |
| Office 2016 | ✅ Có |
| Office 2013 | ✅ Có |

> ⚠️ **Lưu ý:** Office 2010 trở xuống và Windows 7 không được hỗ trợ.

---

## �📦 BƯỚC 1: GIẢI NÉN FILE

### 1.1. Nhận file từ quản trị viên

Bạn sẽ nhận được file: **`ChinhSuaOffice.rar`** (khoảng 49 MB)

<!-- 
📸 ẢNH 1: Screenshot file ChinhSuaOffice.rar trong Windows Explorer
- Hiện file RAR với icon WinRAR
- Hiện kích thước file ~49 MB
-->

### 1.2. Giải nén file

1. **Click chuột phải** vào file `ChinhSuaOffice.rar`
2. Chọn **"Extract Here"** hoặc **"Giải nén tại đây"**

<!-- 
📸 ẢNH 2: Screenshot menu chuột phải khi click vào file RAR
- Highlight dòng "Extract Here" hoặc "Extract to ChinhSuaOffice\"
- Dùng mũi tên hoặc khoanh đỏ để chỉ rõ
-->

### 1.3. Kết quả sau khi giải nén

Sau khi giải nén xong, bạn sẽ thấy **folder mới** tên `ChinhSuaOffice`:

```
📂 ChinhSuaOffice/
├── ChinhSuaOffice.exe    (155 MB - File chạy chính)
├── config.json           (0.1 KB - File cấu hình)
└── HUONG_DAN_SU_DUNG.md  (File hướng dẫn này)
```

<!-- 
📸 ẢNH 3: Screenshot folder ChinhSuaOffice đã giải nén
- Hiện 3 file bên trong: ChinhSuaOffice.exe, config.json, HUONG_DAN_SU_DUNG.md
- Hiện cột Size để thấy kích thước từng file
-->

### 1.4. Nên đặt folder ở đâu?

Bạn có thể để folder ở:
- **Desktop** (Màn hình nền) - Dễ tìm
- **C:\ChinhSuaOffice** - Gọn gàng
- **D:\ChinhSuaOffice** - Nếu ổ C ít dung lượng

---

## 📄 BƯỚC 2: HIỂU CÁC FILE TRONG FOLDER

### 2.1. File `ChinhSuaOffice.exe` - File chạy chính

| Thông tin | Chi tiết |
|-----------|----------|
| **Tên file** | ChinhSuaOffice.exe |
| **Kích thước** | ~155 MB |
| **Công dụng** | Ứng dụng chính để chỉnh sửa file Office từ web |
| **Cách dùng** | Double-click để chạy |

<!-- 
📸 ẢNH 4: Screenshot file ChinhSuaOffice.exe được highlight
- Khoanh đỏ hoặc mũi tên chỉ vào file exe
- Hiện tooltip hoặc Properties nếu cần
-->

### 2.2. File `config.json` - File cấu hình

| Thông tin | Chi tiết |
|-----------|----------|
| **Tên file** | config.json |
| **Kích thước** | ~0.1 KB |
| **Công dụng** | Chứa cấu hình kết nối đến server công ty |
| **Có cần sửa?** | ❌ Không - Đã cấu hình sẵn |

**Nội dung bên trong file config.json:**

```json
{
  "port": 1901,
  "companyApiUrl": "https://administrator.lifetex.vn:316",
  "apiEndpoint": "/api/files/download"
}
```

<!-- 
📸 ẢNH 5: Screenshot file config.json mở bằng Notepad
- Hiện nội dung JSON với 3 dòng cấu hình
- Có thể thêm chú thích bên cạnh giải thích từng dòng
-->

**Giải thích từng dòng:**

| Tham số | Giá trị | Ý nghĩa |
|---------|---------|---------|
| `port` | 1901 | Cổng mà ứng dụng sử dụng trên máy bạn |
| `companyApiUrl` | https://administrator.lifetex.vn:316 | Địa chỉ server công ty |
| `apiEndpoint` | /api/files/download | Đường dẫn API để tải file |

> ⚠️ **Lưu ý:** KHÔNG chỉnh sửa file này trừ khi được hướng dẫn bởi quản trị viên.

---

## 🚀 BƯỚC 3: CHẠY ỨNG DỤNG

### 3.1. Cách chạy

1. Mở folder **ChinhSuaOffice**
2. **Double-click** vào file **`ChinhSuaOffice.exe`**

<!-- 
📸 ẢNH 6: Screenshot double-click vào file exe
- Mũi tên chỉ vào file ChinhSuaOffice.exe
- Có thể thêm icon chuột đang click
-->

### 3.2. Lần đầu chạy - Windows SmartScreen

Nếu Windows hiện thông báo **"Windows protected your PC"**:

<!-- 
📸 ẢNH 7: Screenshot màn hình Windows SmartScreen
- Hiện đầy đủ popup "Windows protected your PC"
- Khoanh đỏ nút "More info"
-->

**Cách xử lý:**

**Bước 1:** Click **"More info"** (Thông tin thêm)

<!-- 
📸 ẢNH 8: Screenshot sau khi click "More info"
- Hiện nút "Run anyway" đã xuất hiện
- Khoanh đỏ nút "Run anyway"
-->

**Bước 2:** Click **"Run anyway"** (Vẫn chạy)

### 3.3. Sau khi chạy - App ở đâu?

> ⚠️ **QUAN TRỌNG:** Sau khi chạy, ứng dụng sẽ **KHÔNG hiện cửa sổ**!

Ứng dụng chạy **NGẦM** và chỉ hiện **ICON** ở **khay hệ thống** (System Tray).

**Vị trí khay hệ thống:**

<!-- 
📸 ẢNH 9: Screenshot toàn màn hình với mũi tên chỉ vào khay hệ thống
- Khoanh đỏ vùng khay hệ thống (góc phải taskbar, gần đồng hồ)
- Mũi tên lớn chỉ vào vị trí đó
- Ghi chú: "Khay hệ thống (System Tray)"
-->

**Vị trí cụ thể trên taskbar:**

```
┌──────────────────────────────────────────────────────────────┐
│ [Start] [...............................] [^] [🔊] [📅 14:00] │
└──────────────────────────────────────────────────────────────┘
                                             ↑
                                             │
                              Khay hệ thống nằm ở đây
```

### 3.4. Tìm icon ứng dụng

**Cách 1:** Nhìn trực tiếp ở khay hệ thống (có thể thấy ngay icon app)

**Cách 2:** Nếu không thấy, click vào **mũi tên `^`** để xem icon ẩn:

<!-- 
📸 ẢNH 10: Screenshot click vào mũi tên ^ để mở khay icon ẩn
- Khoanh đỏ mũi tên ^
- Mũi tên chỉ hướng click
-->

<!-- 
📸 ẢNH 11: Screenshot popup hiện các icon ẩn
- Khoanh đỏ icon của ứng dụng ChinhSuaOffice
- Ghi chú: "Icon ứng dụng"
-->

### 3.5. Xác nhận app đang chạy

**Di chuột (hover)** lên icon ứng dụng, sẽ hiện tooltip:

<!-- 
📸 ẢNH 12: Screenshot hover lên icon, hiện tooltip
- Hiện rõ tooltip "Trình chỉnh sửa Office - Đang chạy"
- Khoanh đỏ tooltip
-->

```
┌─────────────────────────────────────┐
│ Trình chỉnh sửa Office - Đang chạy  │
└─────────────────────────────────────┘
```

✅ **Nếu thấy dòng này = App đang chạy bình thường!**

---

## 🖱️ BƯỚC 4: MENU CỦA ỨNG DỤNG

### 4.1. Mở menu

**Click chuột PHẢI** vào icon ứng dụng ở khay hệ thống

<!-- 
📸 ẢNH 13: Screenshot click chuột phải vào icon
- Hiện icon app
- Mũi tên + text "Click chuột PHẢI"
-->

### 4.2. Menu hiển thị

<!-- 
📸 ẢNH 14: Screenshot menu context hiện ra
- Hiện đầy đủ menu với 2 mục: "✅ Đang chạy" và "❌ Thoát"
- Có thể thêm chú thích bên cạnh giải thích từng mục
-->

```
┌─────────────────────┐
│ ✅ Đang chạy        │  ← Trạng thái (không click được)
├─────────────────────┤
│ ❌ Thoát            │  ← Click để tắt app
└─────────────────────┘
```

| Mục | Ý nghĩa | Click được không? |
|-----|---------|-------------------|
| **✅ Đang chạy** | Hiển thị trạng thái app | ❌ Không (chỉ để xem) |
| **❌ Thoát** | Tắt ứng dụng hoàn toàn | ✅ Có |

---

## ❌ BƯỚC 5: THOÁT ỨNG DỤNG

### 5.1. Cách thoát

1. **Click chuột phải** vào icon app ở khay hệ thống
2. Click **"❌ Thoát"**

<!-- 
📸 ẢNH 15: Screenshot menu với mũi tên chỉ vào "Thoát"
- Khoanh đỏ hoặc highlight mục "❌ Thoát"
- Mũi tên chỉ vào
-->

### 5.2. Sau khi thoát

- Icon app sẽ **biến mất** khỏi khay hệ thống
- App đã **tắt hoàn toàn**
- Muốn dùng lại → Chạy lại file `ChinhSuaOffice.exe`

### 5.3. Khi nào nên thoát?

| ✅ Thoát khi | ❌ Không thoát khi |
|--------------|-------------------|
| Hết ngày làm việc | Đang cần chỉnh sửa file |
| Tắt máy tính | Đang làm việc với web |
| Không cần dùng nữa | Muốn mở file từ web |

### 5.4. Tắt máy tính

Khi bạn **tắt máy tính** hoặc **restart**, ứng dụng sẽ **tự động tắt** theo. Không cần thoát thủ công.

---

## 💻 CÁCH SỬ DỤNG HÀNG NGÀY

### Quy trình đơn giản:

<!-- 
📸 ẢNH 16: Sơ đồ quy trình 4 bước (có thể dùng tool vẽ sơ đồ)
- Bước 1: Bật máy tính
- Bước 2: Chạy app (double-click exe)
- Bước 3: Vào web làm việc
- Bước 4: Chỉnh sửa file Office
-->

```
Bước 1              Bước 2              Bước 3              Bước 4
┌──────────┐       ┌──────────┐       ┌──────────┐       ┌──────────┐
│ Bật máy  │  ──▶  │ Chạy app │  ──▶  │ Vào web  │  ──▶  │ Chỉnh    │
│ tính     │       │ (1 lần)  │       │ làm việc │       │ sửa file │
└──────────┘       └──────────┘       └──────────┘       └──────────┘
```

### Mỗi ngày chỉ cần:

1. **Bật máy tính**
2. **Double-click `ChinhSuaOffice.exe`** (chỉ 1 lần đầu ngày)
3. **Làm việc bình thường** - App chạy ngầm, không cần quan tâm
4. **Tắt máy khi xong** - App tự tắt theo

---

## 🌐 SỬ DỤNG VỚI WEBSITE

### Khi muốn chỉnh sửa file từ web:

1. Đảm bảo **app đang chạy** (xem icon ở khay hệ thống)
2. Vào **website** công ty
3. Tìm file cần chỉnh sửa
4. Click nút **"Chỉnh sửa"** hoặc **"Edit"**

<!-- 
📸 ẢNH 17: Screenshot website với nút "Chỉnh sửa" được highlight
- Hiện giao diện web
- Khoanh đỏ nút "Chỉnh sửa" / "Edit"
-->

5. **Office tự động mở** file (Word/Excel/PowerPoint...)

<!-- 
📸 ẢNH 18: Screenshot Microsoft Word đang mở file từ server
- Hiện giao diện Word với file đang mở
- Có thể highlight thanh tiêu đề hiện URL
-->

6. Chỉnh sửa nội dung
7. Nhấn **Ctrl + S** để lưu lên server

<!-- 
📸 ẢNH 19: Screenshot nhấn Ctrl+S trong Word
- Có thể hiện bàn phím với Ctrl và S được highlight
- Hoặc hiện menu File > Save
-->

---

## ⚠️ LƯU Ý QUAN TRỌNG

### ✅ NÊN làm:

- Chạy app **MỖI LẦN bật máy**
- Để app **chạy ngầm** suốt ngày làm việc
- **Lưu file (Ctrl+S)** thường xuyên khi chỉnh sửa

### ❌ KHÔNG NÊN:

- **Xóa** file `config.json`
- **Chỉnh sửa** file `config.json` khi không được hướng dẫn
- **Di chuyển** riêng file exe ra khỏi folder (phải giữ cùng folder với config.json)

---

## ❓ CÂU HỎI THƯỜNG GẶP

### Q: App có hiện cửa sổ không?
**A:** Không. App chạy ngầm, chỉ có icon ở khay hệ thống.

### Q: Làm sao biết app đang chạy?
**A:** Xem icon ở khay hệ thống (góc phải taskbar). Di chuột lên sẽ thấy "Đang chạy".

### Q: App tự chạy khi bật máy không?
**A:** Có, app sẽ tự đăng ký khởi động cùng Windows sau lần chạy đầu tiên.

### Q: Quên chạy app thì sao?
**A:** Khi click chỉnh sửa file trên web sẽ không hoạt động. Hãy chạy app rồi thử lại.

### Q: Có cần quyền Admin không?
**A:** Không cần.

---

## 📞 HỖ TRỢ

Nếu gặp vấn đề, liên hệ:
- **Email:** support@lifetex.vn
- **Website:** https://lifetex.vn

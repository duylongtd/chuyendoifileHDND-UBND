# Tách bộ biểu VTVL: HĐND – UBND

Web nhận file bộ biểu vị trí việc làm **HĐND-UBND cấp xã** (`.xlsx`) và trả về **2 file riêng**: một file HĐND, một file UBND.
Nội dung các ô giữ nguyên, không thêm bớt; chỉ chia dòng theo vị trí và dồn lên liền nhau.

## Quy tắc tách

**HĐND gồm đúng 4 chức danh** (nhận diện theo tên ở biểu `6. C1_VTVL`, cột B):

- Chủ tịch HĐND, Phó Chủ tịch HĐND
- Trưởng ban, Phó Trưởng ban (của HĐND)

**Mọi vị trí còn lại là UBND**, kể cả "Chánh Văn phòng HĐND và UBND" hay "Chuyên viên tham mưu, giúp việc HĐND".

Vị trí **để trống tên** thì hệ thống chỉ gợi ý (dựa vào mô tả ở biểu `7. C2_NangLuc`) và tô vàng để người dùng xác nhận trước khi tách.

| Biểu | Chia dòng theo |
| --- | --- |
| `6. C1_VTVL`, `7. C2_NangLuc`, `8. C3_NhuCau`, `10. D2_PhuongAn`, `DU_BAO_5_NAM` | Mã VTVL cột A |
| `9. D1_NhanLuc` | Mã VTVL cột B |
| `2. A2_CongViec` | Mã VTVL cột C |
| `3. B1_KhoiLuong`, `4. B2_CT_TacDong`, `5. B2_ThayDoi` | Mã công việc cột A, đối chiếu sang A2 |
| `1. A1_NhiemVu` | Mã nhiệm vụ: nhiệm vụ chỉ của một khối thì vào khối đó; dùng chung hoặc chưa gắn công việc thì có ở **cả hai file** |
| Các biểu khác (hướng dẫn, danh mục, tham số…) | Không tách, giữ nguyên ở cả hai file |

Sau khi tách:

- Các dòng còn lại **dồn lên liền nhau từ dòng 3**, giữ đúng thứ tự gốc.
- Công thức, định dạng, danh sách chọn, ghi chú giữ nguyên. File kết quả bật chế độ **tính lại công thức khi mở** bằng Excel.
- Dòng thiếu mã hoặc mã không tra được thì giữ ở cả hai file và hiện cảnh báo.
- Tên file: `01.HĐND-UBND Phường X.xlsx` → `01.HĐND Phường X.xlsx` và `01.UBND Phường X.xlsx`
  (tên không có cụm "HĐND-UBND" thì thêm hậu tố ` - HĐND` / ` - UBND`).

> Giá trị công thức lưu sẵn trong file (ví dụ các ô đếm ở `13. K_KiemTra`, `14. T_TienDo`) chỉ cập nhật sau khi mở file bằng Excel.

## Giới hạn

| | |
| --- | --- |
| Số file mỗi lượt | 3 |
| Dung lượng mỗi file | 60 MB |
| Thời gian lưu file trên máy chủ | 24 giờ (tự xóa) |
| Số lần tự thử lại khi lỗi hệ thống | 3 |

Sửa trong [`lib/server/config.ts`](lib/server/config.ts).

## Kiến trúc

```
Trình duyệt ──(tải thẳng)──▶ Vercel Blob (private)          jobs/<id>/input.xlsx
     │                              ▲   │
     │ POST /api/jobs/:id/inspect   │   │ đọc file
     │ POST /api/jobs/:id/split ────┼───┼──▶ Vercel Queues (topic split-jobs)
     │                              │   ▼            │ giao lại khi lỗi
     │ GET  /api/jobs?ids=… (2s) ◀──┤  /api/queues/split  (function, tối đa 300s)
     │                              │        │ tách → ghi hdnd.xlsx, ubnd.xlsx, job.json
     └ GET /api/jobs/:id/download ◀─┘◀───────┘
```

- **Next.js (App Router, TypeScript)**: giao diện + API.
- **Vercel Blob (private)**: file gốc, file kết quả và trạng thái tác vụ (`job.json`). Trình duyệt tải thẳng lên Blob nên không vướng giới hạn 4,5 MB của function. Không cần cơ sở dữ liệu.
- **Vercel Queues**: mỗi file là một thông điệp; lỗi hệ thống thì hàng đợi tự giao lại. Nếu Queues chưa bật hoặc gửi lỗi, hệ thống tự chuyển sang chạy nền trong function (`after()`), nên vẫn hoạt động.
- **Mỗi instance chỉ tách 1 file tại một thời điểm** (các file khác xếp hàng) để không tràn bộ nhớ.
- **Cron** hằng ngày xóa tác vụ quá 24 giờ.
- **Bộ tách** ([`lib/xlsx`](lib/xlsx)) làm việc trực tiếp trên XML bên trong `.xlsx`, không dựng cả workbook trong bộ nhớ: file 37 MB (424 MB sau giải nén, 1,26 triệu công thức) tách xong trong vài giây. Các thành phần không thay đổi được chép nguyên dạng nén.

Xử lý lỗi: sai định dạng, file `.xls`, file đặt mật khẩu, file hỏng, không phải bộ biểu VTVL, không có vị trí HĐND, quá dung lượng, quá số file, "bom nén" (giới hạn 1,5 GB sau giải nén), hết phiên đăng nhập, tác vụ treo (hiện nút "Chạy lại").

## Chạy trên máy

```bash
npm install
cp .env.example .env.local   # rồi điền ACCESS_CODE
npm run dev                  # http://localhost:3600
```

Trên máy, file lưu trong thư mục `.data/` và tác vụ chạy nền trong tiến trình (không cần Blob/Queues).

Tách thử bằng dòng lệnh:

```bash
npm run split -- "duong-dan/file.xlsx" thu-muc-ket-qua
```

## Deploy lên Vercel

1. **Import repo** vào Vercel (Framework: Next.js, không cần chỉnh build).
2. **Storage → Create → Blob**, chọn access **Private**, rồi **Connect** vào project (Vercel tự thêm biến môi trường của Blob).
3. **Settings → Environment Variables**, thêm:
   - `ACCESS_CODE`: mã truy cập chung do bạn tự đặt.
   - `SESSION_SECRET`: chuỗi ngẫu nhiên (ví dụ tạo bằng `openssl rand -hex 32`).
   - `CRON_SECRET`: chuỗi ngẫu nhiên, để chỉ Vercel Cron gọi được đường dẫn dọn file.
4. **Redeploy** để nhận biến môi trường.
5. Vercel Queues được đăng ký sẵn trong [`vercel.json`](vercel.json). Nếu tài khoản chưa dùng được Queues, đặt thêm `QUEUE_DRIVER=inline`.

Khuyến nghị:

- **Bộ nhớ function**: mặc định 2 GB là đủ cho file tới 60 MB. Nếu log báo hết bộ nhớ, vào Settings → Functions chọn mức **Performance (4 GB)**.
- **Region** của function nên cùng region với Blob store để đọc/ghi file nhanh.

Thiếu cấu hình bắt buộc (`ACCESS_CODE`, Blob, `CRON_SECRET`) thì API trả lỗi 503 kèm hướng dẫn, không chạy âm thầm.

## Bảo mật

- Vào trang phải có mã truy cập; phiên đăng nhập là cookie `httpOnly` ký HMAC, hết hạn sau 12 giờ; nhập sai quá 10 lần/10 phút bị khóa tạm.
- File lưu ở Blob **private**, đường dẫn chứa mã ngẫu nhiên (UUID), chỉ tải được qua API đã đăng nhập, tự xóa sau 24 giờ. Người dùng có thể bấm "Xóa" để xóa ngay.
- Không đưa file bộ biểu hay `.env` lên repo (đã chặn trong `.gitignore`).

## Kiểm chứng kết quả

[`scripts/verify_split.py`](scripts/verify_split.py) dùng một thư viện khác (openpyxl) để so 2 file kết quả với file gốc, từng ô:

```bash
pip install openpyxl
python scripts/verify_split.py goc.xlsx hdnd.xlsx ubnd.xlsx H27.120_VT2,H27.120_VT4,H27.120_VT5,H27.120_VT6
```

Script kiểm: zip toàn vẹn; thành phần không tách giữ nguyên từng byte; đủ số dòng và số công thức; từng dòng giữ lại khớp từng ô với dòng gốc và đúng thứ tự; không mất, không thừa dòng.

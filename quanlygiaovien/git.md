# Cài app lên máy khác & cấp quyền cập nhật dữ liệu (hướng dẫn Git)

Tài liệu này dành cho người đang **quản lý/sở hữu** repo GitHub của app
(hiện là tài khoản `luantt94`), muốn cài app này cho một máy khác (đồng
nghiệp, máy ở văn phòng...) và cho phép người dùng máy đó tự cập nhật dữ
liệu lên trang web công khai.

Người **dùng app hàng ngày** (chỉ cập nhật Excel/XML rồi bấm nút) không
cần đọc file này — xem mục "3. Quy trình cập nhật web hàng ngày/hàng tuần"
trong [README.md](README.md) là đủ. File này chỉ cần đọc **1 lần khi cài
app lên 1 máy mới**.

## Vì sao app cần tới Git?

Trang web công khai (`gv.html`, `public_data.json`...) nằm trong thư mục
`docs/` của repo GitHub `luantt94/GTCC`, host bằng GitHub Pages. Mỗi lần
bạn bấm **"Cập nhật && Đăng lên web"**, app tự chạy `git add` → `git
commit` → `git push` dữ liệu mới vào `docs/` (xem `git_publish.py`) —
người dùng không cần gõ lệnh git nào cả. Nhưng máy nào muốn bấm nút đó
thành công thì máy đó cần: (1) có bản sao (clone) của repo đúng cấu trúc
thư mục, và (2) tài khoản GitHub đang đăng nhập trên máy đó có **quyền
đẩy (push)** vào repo.

## A. Cài app lên 1 máy mới

### Bước 1 — Cài phần mềm nền

- **Python 3.10 trở lên**: tải tại [python.org](https://www.python.org/downloads/)
  — khi cài nhớ tick **"Add python.exe to PATH"**.
- **Git for Windows**: tải tại [git-scm.com](https://git-scm.com/download/win)
  — cứ bấm Next mặc định là được, bản cài đặt đã có sẵn **Git Credential
  Manager** (GCM) để đăng nhập GitHub bằng trình duyệt, không cần cấu hình
  thêm.

### Bước 2 — Tải mã nguồn (clone repo)

Mở PowerShell hoặc Git Bash, chạy (có thể đổi `D:\GTCC` thành đường dẫn
khác tuỳ ý — app tự tính đường dẫn theo vị trí file, không bắt buộc phải
đúng y hệt `D:\GTCC`, miễn giữ nguyên cấu trúc `docs/` và `quanlygiaovien/`
nằm cùng cấp như khi clone bình thường):

```
git clone https://github.com/luantt94/GTCC.git D:\GTCC
```

Nếu repo đang để **public** thì clone không cần đăng nhập gì cả. Nếu repo
đã chuyển sang **private**, máy mới cần được mời làm Collaborator (xem
mục B) *trước* khi clone được.

### Bước 3 — Cài thư viện Python

```
cd D:\GTCC\quanlygiaovien
pip install -r requirements.txt
```

### Bước 4 — Chép dữ liệu nội bộ từ máy cũ qua (KHÔNG nằm trong Git)

Những mục sau bị `.gitignore` chặn, **không** tự có khi clone (vì là dữ
liệu nhạy cảm hoặc đặc thù theo máy) — nếu muốn máy mới có đầy đủ dữ liệu
đã làm dở từ máy cũ, tự chép tay qua (USB, Zalo, Google Drive...) vào đúng
thư mục `quanlygiaovien/` trên máy mới:

| Chép gì | Để làm gì | Bắt buộc? |
|---|---|---|
| `data/` (Excel, DSGV/DSXTL XML, giấy tờ xe, hồ sơ GV) | Nguồn dữ liệu gốc | **Có** — không có thì chưa đăng được gì |
| `anh_gv/` | Ảnh đại diện giáo viên đã thêm | Nên có, không thì phải thêm ảnh lại từ đầu |
| `giao_vien_extra.json` | Hạng lái xe dạy đã nhập | Nên có |
| `xe_extra.json` | Link giấy tờ xe đã nhập | Nên có |
| `app_settings.json` | Mật khẩu hồ sơ cá nhân + thông tin in thẻ GV | Nên có |
| `qr_codes/` | Mã QR đã tạo sẵn | Không bắt buộc — nội dung mã QR không đổi, thiếu app tự tạo lại |
| `data_sources.json` | Đường dẫn các file đã chọn (đường dẫn tuyệt đối) | **Không nên chép** — xem lưu ý dưới |

> **Lưu ý về `data_sources.json`:** file này lưu đường dẫn **tuyệt đối**
> theo máy cũ (vd `D:\GTCC\quanlygiaovien\data\DSGV_...xml`). Nếu máy mới
> cài khác đường dẫn, chép file này qua sẽ khiến app hiện đường dẫn sai.
> Cách đơn giản nhất: **đừng chép** `data_sources.json` — chỉ cần chép
> nguyên thư mục `data/` qua, lần đầu mở app sẽ **tự quét lại** `data/` và
> điền sẵn đúng nguồn (xem mục 2 trong README.md).

Bỏ qua Bước 4 cũng mở app được bình thường — chỉ cần vào tab **CẬP NHẬT
WEB** chọn lại nguồn dữ liệu, và vào tab **DANH SÁCH GIÁO VIÊN**/**DANH
SÁCH XE** nhập lại ảnh/hạng/link từ đầu.

### Bước 5 — Mở app

```
python main.py
```

## B. Cho phép máy mới đăng dữ liệu lên web (quyền push GitHub)

Chọn 1 trong 2 cách:

### Cách 1 (khuyên dùng) — Thêm người dùng mới làm Collaborator

1. Người dùng máy mới cần có **tài khoản GitHub riêng** (miễn phí, tự đăng
   ký tại github.com nếu chưa có).
2. Chủ repo (`luantt94`) vào
   `https://github.com/luantt94/GTCC/settings/access` → **Add people** →
   nhập username hoặc email người đó → gửi lời mời.
3. Người được mời mở email hoặc thông báo trên GitHub, bấm **Accept
   invitation**.
4. Trên máy mới, lần đầu bấm **"Cập nhật && Đăng lên web"** (hoặc bất kỳ
   nút nào có đăng web — đổi ảnh, lưu link xe...), Git Credential Manager
   sẽ tự mở trình duyệt yêu cầu đăng nhập GitHub — đăng nhập bằng **tài
   khoản của người đó**. Các lần sau không cần đăng nhập lại.

Ưu điểm: mỗi người một tài khoản, lịch sử commit biết rõ ai sửa; thu hồi
quyền (xoá khỏi Collaborators) được bất cứ lúc nào nếu cần.

### Cách 2 — Dùng chung 1 tài khoản GitHub

Đơn giản hơn (khỏi mời ai), nhưng mọi commit đều đứng tên chung 1 người,
không phân biệt được máy nào sửa. Chỉ nên dùng nếu là nhóm nhỏ, tin tưởng
nhau và không cần theo dõi ai sửa gì:

1. Trên máy mới, khi Git Credential Manager hỏi đăng nhập (lần đầu bấm nút
   đăng web), đăng nhập bằng tài khoản GitHub **chung** (tài khoản chủ repo).

> Dù chọn cách nào, mật khẩu hồ sơ cá nhân/ảnh/dữ liệu đều **không** liên
> quan tới tài khoản GitHub — tài khoản GitHub chỉ dùng để xác nhận máy đó
> được phép ghi vào repo.

## C. Dùng app hàng ngày trên máy mới

Giống hệt máy gốc — xem mục **"3. Quy trình cập nhật web hàng ngày/hàng
tuần"** trong [README.md](README.md).

## D. Lưu ý khi có NHIỀU máy cùng cập nhật

App hiện **không tự `git pull`** trước khi đăng (xem `git_publish.py`).
Nếu 2 máy cùng bấm "Cập nhật && Đăng lên web" gần nhau, máy bấm **sau** có
thể bị GitHub từ chối push (lỗi "rejected - non-fast-forward" hiện trong
hộp thoại lỗi) vì máy đó đang giữ bản cũ hơn máy kia vừa đẩy lên. Cách xử
lý: mở Git Bash tại thư mục gốc (`D:\GTCC`) và chạy:

```
git pull
```

rồi bấm lại nút "Cập nhật && Đăng lên web" trong app. Để tránh tình huống
này, tốt nhất **quy ước rõ 1 máy chính** chuyên cập nhật dữ liệu hàng
ngày; máy khác chỉ dùng để tra cứu (tab TRA CỨU/PHÂN XE/DANH SÁCH XE/LỊCH
HỌC) hoặc cập nhật ảnh/hạng/link riêng lẻ khi thật sự cần.

## E. Sự cố thường gặp liên quan tới Git

| Vấn đề | Cách xử lý |
|---|---|
| Bấm nút đăng web, trình duyệt không tự mở để đăng nhập GitHub | Kiểm tra đã cài Git for Windows (có Git Credential Manager) chưa; thử mở Git Bash gõ `git push` tại `D:\GTCC` để xem lỗi chi tiết |
| Lỗi "Permission denied" / "403" khi đăng | Tài khoản GitHub đang đăng nhập trên máy đó **chưa được** thêm làm Collaborator — làm theo mục B |
| Lỗi "rejected - non-fast-forward" | Máy khác vừa đăng trước — chạy `git pull` ở `D:\GTCC` rồi bấm lại nút trong app (xem mục D) |
| Muốn đổi sang tài khoản GitHub khác trên máy đang dùng | Mở **Windows Credential Manager** (gõ "Credential Manager" trong Start Menu) → mục Windows Credentials → tìm `git:https://github.com` → **Remove** → lần bấm nút đăng web kế tiếp app sẽ hỏi đăng nhập lại |
| Repo chuyển private, máy cũ bỗng không clone/pull được nữa | Phải được thêm làm Collaborator trước (xem mục B) |

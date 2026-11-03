# Quản lý giáo viên / học viên — Trung tâm đào tạo lái xe

App desktop để tra cứu giáo viên/xe/lịch học, và đăng thông tin giáo viên
(dạng mã QR) lên một trang web công khai, cập nhật tự động mỗi khi bạn
bấm nút — không cần biết git hay lập trình.

## 1. Cách mở app

Bấm đúp vào file chạy app, hoặc mở terminal tại thư mục này và chạy:

```
python main.py
```

App có 6 tab:

| Tab | Dùng để |
|---|---|
| **CẬP NHẬT WEB** | Khai báo nguồn dữ liệu, mật khẩu hồ sơ cá nhân, thông tin in trên thẻ GV & đăng số liệu lên trang web (mã QR) |
| **DANH SÁCH GIÁO VIÊN** | Duyệt danh sách GV, xem ngay thẻ (ảnh có mã QR) từng người, đổi ảnh/hạng lái xe dạy, xuất PDF để in |
| TRA CỨU | Tìm giáo viên / mã khóa, xem học viên, lịch học |
| PHÂN XE | Xem danh sách xe và giáo viên phụ trách |
| **DANH SÁCH XE** | Duyệt danh sách xe, xem lịch sử giáo viên/khóa đã dùng, giấy tờ xe cục bộ, nhập link giấy tờ xe cho trang web |
| LỊCH HỌC | Xem trực tiếp nội dung file PDF lịch học thực hành theo khóa |

> **Lưu ý:** 5 tab DANH SÁCH GIÁO VIÊN / TRA CỨU / PHÂN XE / DANH SÁCH XE /
> LỊCH HỌC chỉ đọc dữ liệu nội bộ lúc **mở app** (riêng DANH SÁCH GIÁO VIÊN
> vẫn sửa ảnh/hạng/xuất PDF và đăng lên web được ngay, chỉ danh sách TÊN
> giáo viên là cần mở lại app nếu đổi file DSGV). Sau khi bấm "Cập nhật &&
> Đăng lên web" ở tab đầu tiên, hãy **khởi động lại app** để các tab này
> thấy dữ liệu mới nhất.

## 2. Nguồn dữ liệu

Toàn bộ dữ liệu lấy từ thư mục `data/` (không còn dùng file Excel tổng hợp
cũ). Tab **CẬP NHẬT WEB** có 6 vùng khai báo, tự động quét sẵn từ `data/`
khi mở app lần đầu — bạn chỉ cần sửa khi có thay đổi:

| Vùng | Nội dung | Khi nào cần sửa |
|---|---|---|
| File Excel "BIÊN CHẾ XE KHÓA ..." | Mỗi khóa đào tạo 1 file `.xlsx` hoặc `.xls` (97-2003) | Có khóa mới → bấm **Thêm...**; khóa đã xong → chọn dòng rồi **Xóa mục chọn** |
| File PDF "LICH HOC THUC HANH ..." | Mỗi khóa 1 file lịch học `.pdf` | Tương tự, thêm/xóa theo khóa |
| DSGV_*.xml | Danh sách **toàn bộ giáo viên** (nguồn chính thức, dùng để tạo mã QR) | Khi có bản DSGV mới xuất ra từ hệ thống |
| DSXTL_*.xml | Danh sách xe tập lái (đối chiếu biển số) | Khi có bản DSXTL mới |
| Thư mục giấy tờ xe tập lái | Chứa file `.docx`/`.pdf` giấy tờ từng xe - chỉ dùng trong app (tab DANH SÁCH XE), **không đưa lên trang web** | Khi đổi chỗ lưu thư mục |
| Thư mục hồ sơ cá nhân giáo viên | Chứa file `.pdf` hồ sơ từng giáo viên (ghép theo tên trong tên file) | Khi đổi chỗ lưu thư mục |

### Link giấy tờ xe trên trang web

Trang web **không còn nhúng file PDF/Word giấy tờ xe** nữa. Thay vào đó, ở
tab **DANH SÁCH XE**: chọn biển số trong ô danh sách, dán link (vd trang
tra cứu xe tập lái của Sở GTVT) vào ô "Link giấy tờ xe" rồi bấm **Lưu
link**. Bấm "Cập nhật && Đăng lên web" (tab CẬP NHẬT WEB) để áp dụng —
trang `gv.html` sẽ hiện nút "Xem giấy tờ xe (trang của Sở)" mở link này ở
tab mới, thay cho việc nhúng file như trước.

App tự ghép file PDF/Word vào đúng khóa/xe/giáo viên bằng cách so khớp tên
file (biển số xe, mã khóa, tên giáo viên) — **không cần khai báo tay**.

### Mật khẩu xem hồ sơ cá nhân

Ô "Mật khẩu xem hồ sơ cá nhân giáo viên" dùng chung cho tất cả giáo viên.
**Mật khẩu mặc định: `gtcc2026`** (máy mới chưa từng lưu cài đặt riêng sẽ
tự điền sẵn mật khẩu này — đổi/xóa được bất cứ lúc nào ở ô này):
- Có đặt mật khẩu → trang web hiện mục "Hồ sơ cá nhân", người xem phải
  nhập đúng mật khẩu mới xem được (hồ sơ có CCCD/bằng lái nên cần hạn chế).
- Để **trống** → ẩn hoàn toàn mục hồ sơ cá nhân trên trang web.
- Bấm **Lưu mật khẩu**, sau đó bấm "Cập nhật && Đăng lên web" để áp dụng.

> **⚠️ Đây KHÔNG phải bảo mật thật sự, chỉ là lớp che tượng trưng.**
> Trang web chạy trên GitHub Pages (hosting tĩnh, không có server kiểm
> tra mật khẩu). Đường link PDF hồ sơ của **mọi** giáo viên vẫn nằm sẵn
> trong `docs/public_data.json` — ai mở DevTools (F12) → tab Network,
> hoặc tự tải file này bằng tay, đều xem/tải được hồ sơ **mà không cần
> biết mật khẩu**. Mật khẩu chỉ ngăn người dùng phổ thông bấm xem nhầm
> qua giao diện bình thường, không ngăn được người cố ý/biết kỹ thuật.
> Đây là lựa chọn đã cân nhắc (ưu tiên đơn giản, không cần thêm hạ tầng
> server) — nếu cần chặn thật, phải chuyển hồ sơ sang một dịch vụ có
> backend kiểm tra quyền truy cập, không thể làm được với GitHub Pages.

### Ảnh đại diện, hạng lái xe dạy & thẻ giáo viên — tab DANH SÁCH GIÁO VIÊN

Tab này gồm danh sách toàn bộ giáo viên bên trái; bấm vào 1 người là thấy
ngay **bản xem trước đúng thẻ thật** (có ảnh + mã QR) của người đó bên
phải, cập nhật **ngay lập tức** mỗi khi đổi ảnh hoặc gõ hạng — không cần
xuất file ra mới xem được.

- **Chọn ảnh...**: chọn 1 file ảnh (.png/.jpg) từ máy.
- **Dán ảnh (Ctrl+V)**: dán ảnh đang copy trong clipboard (vd chụp màn
  hình, hoặc copy từ phần mềm quét/máy scan) — không cần lưu file ra trước.
- **Xóa ảnh**: gỡ ảnh khỏi trang web.

Dù chọn file hay dán, ảnh đều được tự resize nhỏ lại rồi **đăng lên web
ngay lập tức** (không cần bấm "Cập nhật && Đăng lên web"), hiện ngay trên
đầu trang `gv.html` của giáo viên đó. Ảnh này **công khai, không cần mật
khẩu** (khác với "Hồ sơ cá nhân" ở mục trên). Cần Internet để dùng 3 nút
ảnh này.

Ô **"Dạy lái xe hạng"** (vd `B`, `B2, C`) hiện ngay trên bản xem trước khi
gõ; bấm **Lưu hạng** để lưu lại thật sự. Ảnh đại diện và hạng lái xe dạy
lưu trong 2 file **riêng** (`anh_gv/` và `giao_vien_extra.json`, theo mã
GV) — **tách biệt khỏi DSGV_*.xml/DSXTL_*.xml**, nên khi sau này đổi sang
file DSGV/DSXTL mới (bản xuất lại từ Sở), 2 thông tin này **không bị mất**,
miễn mã GV của giáo viên đó giữ nguyên.

Khung **"Thông tin in trên thẻ giáo viên"** (tab CẬP NHẬT WEB) dùng để
nhập tên cơ quan chủ quản + tên cơ sở đào tạo in ở đầu mỗi thẻ (bấm **Lưu**,
dùng chung cho mọi thẻ).

Bấm **"Xuất thẻ GV này (PDF)"** để xuất thẻ của 1 giáo viên đang chọn, hoặc
**"Xuất tất cả giáo viên (PDF)"** để xuất toàn bộ (tối đa 8 thẻ/trang A4,
tự ngắt trang). Đây là tính năng **in cục bộ, không đăng lên web** - không
cần Internet. Phần mềm không tự đóng dấu giáp lai lên ảnh thẻ, cần đóng
dấu tay sau khi in như quy định.

Nếu có file **`data/khunganhnen.jpg`** (ảnh nền thẻ đã thiết kế sẵn, đúng
logo/màu/bố cục của đơn vị), phần mềm sẽ dùng ảnh này làm nền và tự động
dán đè: ảnh giáo viên, mã QR, tên cơ quan/cơ sở đào tạo (ghi đè lên băng
xanh có sẵn, lấy đúng nội dung đã lưu ở khung "Thông tin in trên thẻ giáo
viên"), họ tên (luôn 1 dòng, ngay dưới tiêu đề) và hạng lái xe dạy. Không
có file nền thì phần mềm tự vẽ thẻ theo đúng khổ/cỡ chữ quy định ở Mẫu số
01 (85mm x 50mm). File `khunganhnen.jpg` nằm trong `data/` nên **không đưa
lên git** (giống Excel/XML) - máy nào cần ảnh nền riêng thì tự đặt file
vào đúng đường dẫn này.

> Thẻ in cục bộ **không** có ngày sinh/địa chỉ - 2 trường này (đọc từ
> NgaySinh/NoiCT trong DSGV_*.xml) hiện **công khai trên trang web**
> `gv.html` (mục "Thông tin giáo viên"), xem mục 5.

## 3. Quy trình cập nhật web hàng ngày/hàng tuần

**Bước 1 — Cập nhật file Excel/PDF/XML** như bình thường (thêm học viên,
đổi lịch, đổi xe, xuất lại DSGV/DSXTL...), nhớ **Save** lại trong thư mục
`data/`.

**Bước 2 — Mở app**, vào tab **"CẬP NHẬT WEB"**, kiểm tra lại 6 vùng dữ
liệu ở trên (thêm/xóa nếu cần).

**Bước 3 — Bấm "Cập nhật && Đăng lên web"**. App sẽ tự động:
1. Đọc Excel/PDF/XML đã chọn, ghép đúng khóa/xe/giáo viên.
2. Tính số học viên đang học của từng giáo viên.
3. Nếu có **giáo viên mới** (có trong DSGV nhưng chưa từng đăng), tự tạo
   mã QR mới trong thư mục `qr_codes/` (app báo rõ tên file, ví dụ
   `48002017.png` — in ra và đưa cho giáo viên đó, **chỉ cần in 1 lần**).
4. Đăng dữ liệu mới nhất lên trang web (GitHub).

**Bước 4 — Đọc khung log** sau khi chạy xong. Nếu có dòng bắt đầu bằng
"CẢNH BÁO" (ví dụ tên giáo viên trong Excel không khớp với DSGV do viết
tắt, hoặc biển số không có trong DSXTL), hãy sửa lại cho đúng chính tả
trong file Excel rồi đăng lại — nếu không sửa, giáo viên/xe đó sẽ bị thiếu
thông tin trên trang web.

Vài giây đến vài phút sau, quét lại mã QR cũ của bất kỳ giáo viên nào cũng
sẽ thấy số liệu mới — **mã QR không cần tạo lại**.

> Máy chạy app cần có **kết nối Internet** ở Bước 3. Nếu mất mạng, app báo
> lỗi rõ ràng, bạn chỉ cần bấm lại nút khi có mạng.

## 4. Mã QR giáo viên

Mỗi giáo viên có 1 file ảnh QR trong thư mục `qr_codes/`, đặt tên theo
**mã giáo viên chính thức (MaGV trong DSGV)**, ví dụ `48002017.png`, có in
tên giáo viên ngay bên dưới mã QR để dễ phân biệt khi in/dán. In mã này ra
và phát cho giáo viên — quét bằng camera điện thoại sẽ mở trang hiển thị:

- Các khóa giáo viên đang dạy (mã khóa, lớp, thời gian đào tạo).
- Xe tập lái phụ trách (bấm vào để xem SĐT + giấy tờ xe).
- Số lượng học viên (bấm vào để xem danh sách tên - tên nào có ghi chú ở
  cột "Ghi chú" trong Excel thì hiện kèm trong ngoặc, vd "Tên (Bảo lưu)").
  Học viên ghi chú đúng **"Bảo lưu"** (không phân biệt hoa/thường, có/không
  dấu) vẫn hiện trong danh sách nhưng **không tính vào** số lượng học viên.
- Lịch học (bấm vào để xem bảng lịch chi tiết + file PDF gốc).
- Hồ sơ cá nhân (nếu đã đặt mật khẩu) — cần nhập đúng mật khẩu mới xem.

Mã QR này **cố định vĩnh viễn** theo MaGV — in ra 1 lần dùng mãi, các lần
cập nhật sau không đổi link/mã QR của giáo viên cũ, chỉ sinh thêm mã mới
cho giáo viên mới xuất hiện trong DSGV. Nếu DSGV sửa lại tên 1 giáo viên
đã có mã QR (ví dụ sửa chính tả), app tự nhận ra và **vẽ lại ảnh QR với
tên mới** ở lần đăng web kế tiếp — link/mã QR vẫn giữ nguyên, không cần
in lại, không cần xóa file tay.

## 5. Thông tin bảo mật

Trang web là **công khai trên Internet** (ai có link đều xem được). Những
gì **lên web**: họ tên, **số điện thoại, ngày sinh, địa chỉ thường trú**
của giáo viên (lấy từ trường DienThoai/NgaySinh/NoiCT trong DSGV_*.xml -
mục "Thông tin giáo viên" trên `gv.html`, **công khai, không cần mật
khẩu**, theo yêu cầu), mã khóa/lớp, thời gian học, biển số xe + link giấy
tờ xe (nếu đã nhập - xem mục 2, **không** còn nhúng file PDF/Word giấy tờ
xe), **tên + ghi chú** học viên (vd "Bảo lưu" - xem mục 4), lịch học (kèm
file PDF gốc), và hồ sơ cá nhân giáo viên **nếu** có đặt mật khẩu xem (xem
mục 2).

> ⚠️ Số điện thoại/ngày sinh/địa chỉ là thông tin khá nhạy cảm - đây là
> lựa chọn đã cân nhắc và xác nhận với người quản lý app (mặc định trước
> đây KHÔNG public các trường này). Muốn ẩn lại, cần sửa code
> (`publish_data.run_pipeline`, bỏ 3 dòng `dien_thoai`/`ngay_sinh`/
> `dia_chi` khi dựng `teachers_out`) rồi đăng lại.

Những gì **KHÔNG BAO GIỜ** lên web, kể cả khi đã đặt mật khẩu hồ sơ cá
nhân — app chủ động không đọc các trường này từ DSGV/DSXTL:
- Ảnh chân dung (AnhCD) và số CMT/CCCD (SoCMT) trong file DSGV — **kể cả
  của giáo viên**, dù SĐT/ngày sinh/địa chỉ của giáo viên đã public như trên.
- File Excel/XML gốc, và toàn bộ thư mục `data/` (chỉ những file được app
  copy sang `docs/files/` theo đúng giáo viên/xe/khóa mới public).
- Mật khẩu hồ sơ cá nhân ở dạng gốc (chỉ lưu bản băm SHA-256 trên web).
- Ngày sinh/địa chỉ của **học viên** trong Excel - chỉ giáo viên mới có 2
  trường này public, học viên thì không.

## 6. Cài đặt lần đầu (dành cho máy mới)

```
pip install -r requirements.txt
```

Nếu đây là máy mới hoàn toàn chưa từng đăng nhập GitHub, lần đầu bấm
"Cập nhật && Đăng lên web" có thể hiện cửa sổ đăng nhập GitHub — đăng nhập
1 lần, các lần sau không cần nhập lại.

> **Cài app cho một máy khác (đồng nghiệp, máy văn phòng...) từ đầu?** Xem
> hướng dẫn chi tiết hơn — clone repo, chép dữ liệu nội bộ, cấp quyền đăng
> web cho tài khoản GitHub khác — ở [git.md](git.md).

### Bật GitHub Pages (chỉ làm 1 lần cho repo)

1. Vào `https://github.com/luantt94/GTCC/settings/pages`
2. **Source**: "Deploy from a branch" → Branch: `main`, thư mục `/docs` → **Save**.
3. Sau 1–2 phút, trang chạy tại: `https://luantt94.github.io/GTCC/gv.html?id=<mã GV>`

## 7. Xử lý sự cố thường gặp

| Vấn đề | Cách xử lý |
|---|---|
| "Hãy thêm ít nhất 1 file 'BIÊN CHẾ XE KHÓA ...'" | Vào tab CẬP NHẬT WEB, bấm "Thêm..." ở vùng Excel |
| "Hãy chọn file danh sách giáo viên (DSGV_*.xml)" | Chọn file DSGV ở vùng tương ứng |
| Đã lưu ở máy nhưng chưa đăng được lên web | Kiểm tra Internet, bấm lại nút "Cập nhật && Đăng lên web" |
| Cảnh báo "trùng mã khóa" | Hai file Excel đang chọn có cùng 1 mã khóa — kiểm tra lại có chọn nhầm file không |
| Cảnh báo "không khớp được giáo viên" / "biển số không có trong DSXTL" | Tên/biển số viết tắt hoặc sai chính tả trong Excel so với DSGV/DSXTL — sửa lại cho khớp rồi đăng lại |
| Quét QR báo "Không tìm thấy mã giáo viên" | Chưa bật GitHub Pages, hoặc mới đăng xong cần đợi vài phút |
| Tab TRA CỨU/PHÂN XE/LỊCH HỌC/DANH SÁCH XE chưa thấy dữ liệu mới | Khởi động lại app sau khi đăng web xong |
| Tab DANH SÁCH GIÁO VIÊN chưa thấy giáo viên mới | Khởi động lại app sau khi đổi file DSGV_*.xml (ảnh/hạng/xuất thẻ của giáo viên đã có trong danh sách thì không cần khởi động lại) |

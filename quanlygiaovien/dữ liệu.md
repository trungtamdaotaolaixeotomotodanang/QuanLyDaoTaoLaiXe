# Dữ liệu & đồng bộ nhiều máy

Tài liệu này trả lời 3 câu hỏi hay gặp nhất khi **nhiều máy cùng dùng app**:
1. Nhiều máy dùng chung được không?
2. Máy A sửa dữ liệu, máy B có tự thấy không?
3. Máy A sửa **sai**, sau đó sửa lại cho **đúng** - máy B có tự thấy bản
   đúng không, hay vẫn bị kẹt ở bản sai cũ?

> Đây là tài liệu tra cứu nhanh theo **loại dữ liệu**. Cách cài máy mới,
> cấp quyền GitHub, xử lý xung đột git chi tiết → xem [git.md](git.md).
> Cách dùng app hàng ngày → xem [README.md](README.md).

## 1. Trả lời nhanh

**Có, nhiều máy dùng chung được** - nhưng chỉ với điều kiện mỗi máy là một
bản **git clone thật sự** của repo (xem mục 3), không phải chỉ giải nén
file `QuanLyDaoTaoLaiXe_release.zip` ra một thư mục trống. Nếu không có
git, mọi tính năng "đăng lên web"/"đồng bộ" bên dưới sẽ báo lỗi hoặc im
lặng không làm gì cả.

**Có tự cập nhật qua máy khác không** → tùy loại dữ liệu, xem bảng mục 2.
Một số loại (ngày học gán tay, tên GV xử lý tay, file lịch học gán tay,
ảnh đại diện GV) đồng bộ **ngay lập tức** qua GitHub. Một số loại khác
(toàn bộ nội dung `data.json`: học viên, phân xe, hạng lái xe dạy, link
giấy tờ xe) **chỉ tồn tại cục bộ trên từng máy**, không tự qua máy khác -
xem mục 4.

**Sửa sai rồi sửa lại cho đúng** → với dữ liệu có đồng bộ (ngay_hoc/
ten_gv/lich_hoc overrides), máy khác **sẽ thấy đúng bản mới nhất**, kể cả
khi máy đó đã lỡ tải bản sai trước đó trong cùng 1 phiên làm việc (đã kiểm
tra + sửa lỗi ngày 2026-11-11 - xem ghi chú trong mục 2).

## 2. Bảng tổng hợp theo loại dữ liệu

| Dữ liệu | Sửa ở đâu | Đồng bộ qua máy khác? | Bao lâu thấy ở máy khác | Sửa sai → sửa lại đúng, máy khác có cập nhật không? |
|---|---|---|---|---|
| **Ngày bắt đầu/kết thúc gán tay** (`ngay_hoc_overrides.json`) | Tab PHÂN XE → "Xem chi tiết..." → sửa ô ngày → "Lưu thay đổi" | ✅ Tự động (push ngay khi lưu) | Máy khác: mở lại app, hoặc bấm nút **"🔄 Lấy dữ liệu mới nhất"** ở góc trên màn hình chính | ✅ Có - kể cả không khởi động lại app (nút 🔄 tự nạp lại đúng bản mới, không bị kẹt ở bản cũ đã lỡ hiện trong phiên) |
| **Tên GV xử lý tay** (`ten_gv_overrides.json`) - khi Excel ghi tên viết tắt/sai không khớp DSGV | Tự động hỏi khi "Cập nhật && Đăng lên web" gặp tên không khớp | ✅ Tự động (push ngay) | File override qua ngay; **hiệu lực thật sự** (áp vào data.json của máy đó) chỉ có sau khi máy đó tự bấm "Cập nhật && Đăng lên web" ít nhất 1 lần | ✅ Có (ghi đè giá trị mới trong file override) |
| **File lịch học gán tay** (`lich_hoc_overrides.json`) | Tab LỊCH HỌC, khi không tự khớp được PDF | ✅ Tự động (push ngay) | Giống dòng trên (file qua ngay, hiệu lực sau lần "Cập nhật && Đăng lên web" kế tiếp) | ✅ Có |
| **Ảnh đại diện giáo viên** (bản đã resize, public) | Tab DANH SÁCH GIÁO VIÊN → Chọn ảnh/Dán ảnh/Xóa ảnh | ✅ Tự động (push ngay), hiện ngay trên `gv.html` | Gần như ngay lập tức trên trang web | ✅ Có (ảnh mới ghi đè ảnh cũ) |
| **`docs/public_data.json`** (toàn bộ dữ liệu trang web công khai) | Tự tạo mỗi khi bấm "Cập nhật && Đăng lên web", hoặc các thao tác "✅ Tự động" ở trên | ✅ Tự động (đây chính là dữ liệu đồng bộ) | Trang web: vài giây - vài phút (GitHub Pages build). App máy khác: xem mục 1 | ✅ Có |
| **`data.json`** (dữ liệu đầy đủ app dùng: học viên, phân xe, SĐT, ngày sinh, địa chỉ...) | Tự sinh ra mỗi khi bấm "Cập nhật && Đăng lên web" (đọc từ Excel) | ❌ **KHÔNG** - cố ý không đưa lên GitHub vì chứa thông tin nhạy cảm của học viên (ngày sinh, địa chỉ) và repo đang **public** | Không áp dụng | Không áp dụng - xem mục 4 để biết cách đồng bộ thật sự (chia sẻ file Excel gốc) |
| **Hạng lái xe dạy** (`giao_vien_extra.json`) | Tab DANH SÁCH GIÁO VIÊN → ô "Dạy lái xe hạng" → "Lưu hạng" | ❌ Không đồng bộ máy khác (chỉ lưu local) | Chỉ lên **trang web công khai** sau khi máy đó bấm "Cập nhật && Đăng lên web" - máy khác vẫn không thấy giá trị này trong app của họ | — |
| **Link giấy tờ xe** (`xe_extra.json`) | Tab DANH SÁCH XE → ô "Link giấy tờ xe" → "Lưu link" | ❌ Không đồng bộ máy khác (chỉ lưu local) | Giống dòng trên | — |
| **Mật khẩu xem hồ sơ cá nhân** (`app_settings.json`) | Tab CẬP NHẬT WEB | ❌ File gốc không đồng bộ, nhưng **bản băm (hash)** của nó được ghi vào `docs/public_data.json` mỗi lần "Cập nhật && Đăng lên web" | Web: theo lần đăng gần nhất của **bất kỳ máy nào** - máy nào đăng sau cùng, mật khẩu của máy đó thắng | — |
| **Nguồn dữ liệu đã chọn** (`data_sources.json`: đường dẫn Excel/XML/PDF) | Tab CẬP NHẬT WEB | ❌ Không đồng bộ (cố ý - đường dẫn tuyệt đối, khác nhau theo máy) | — | — |

## 3. Điều kiện bắt buộc để nhiều máy dùng chung được

Mọi dòng "✅ Tự động" ở bảng trên đều chạy bằng lệnh `git` ngầm bên dưới
(xem `git_publish.py`). Máy nào muốn dùng được các tính năng đó (và nút
**"🔄 Lấy dữ liệu mới nhất"**) phải thỏa **cả 3 điều kiện**:

1. Thư mục chứa `main.exe`/`main.py` nằm **bên trong một bản `git clone`**
   của repo `QuanLyDaoTaoLaiXe` (có thư mục `.git` ở thư mục gốc, `docs/`
   và `quanlygiaovien/` nằm cùng cấp) - xem [git.md mục A](git.md).
   → **Chỉ giải nén `QuanLyDaoTaoLaiXe_release.zip` ra một thư mục trống,
   không có sẵn git clone, sẽ KHÔNG dùng được các tính năng đồng bộ** - mọi
   nút lưu/đăng sẽ báo lỗi git (not a git repository).
2. Máy đó đã cài **Git for Windows** và đã đăng nhập GitHub ít nhất 1 lần
   (Git Credential Manager tự nhớ, không cần đăng nhập lại mỗi lần).
3. Tài khoản GitHub đang đăng nhập trên máy đó có **quyền push** vào repo
   (Collaborator) - xem [git.md mục B](git.md).

> **Gửi bản cập nhật cho máy đã cài sẵn (không phải cài mới):** chỉ cần
> build lại bằng `pyinstaller main.spec --noconfirm` rồi chạy
> `python build_release.py` (tự đồng bộ `main.exe` + toàn bộ file dữ liệu
> cần thiết vào `release/`), nén `release/` thành zip rồi gửi - máy nhận
> **giải nén đè vào đúng thư mục `release/` đã có sẵn trong git clone của
> họ** (không giải nén ra chỗ khác), để `.git` ở thư mục gốc vẫn còn đó.

## 4. Những gì KHÔNG tự đồng bộ giữa các máy nội bộ - và cách xử lý

`data.json` (học viên, phân xe, xe/giáo viên phụ trách...) đọc trực tiếp
từ file Excel/PDF/XML trong thư mục `data/` - **không đưa lên GitHub**
(xem `.gitignore`, lý do: chứa ngày sinh/địa chỉ học viên, repo đang
public). Nếu 2 máy cần thấy **cùng một dữ liệu đầy đủ** này:

- Máy nào sửa file Excel gốc (thêm học viên, đổi xe, đổi lịch...) cần
  **tự chia sẻ lại file Excel đó** cho máy kia (USB, Zalo, Google Drive...)
  - giống hệt cách làm khi cài máy mới ở [git.md mục A bước 4](git.md).
- Máy nhận file Excel mới chỉ cần thêm/thay file đó ở tab CẬP NHẬT WEB rồi
  bấm "Cập nhật && Đăng lên web" - `data.json` của máy đó sẽ tự cập nhật
  đúng.
- Việc đồng bộ Excel gốc này **không có cơ chế tự động trong app** (khác
  với 4 dòng "✅ Tự động" ở bảng trên) - cần tự làm bằng tay.

`giao_vien_extra.json` (hạng lái xe dạy) và `xe_extra.json` (link giấy tờ
xe) cũng vậy: sửa ở máy nào chỉ có hiệu lực cục bộ ở máy đó + trang web
công khai (sau khi máy đó đăng), **không hiện trong app của máy khác**.
Nếu nhiều người cùng cần nhập 2 loại này, nên thống nhất **chỉ 1 máy** phụ
trách nhập, hoặc tự chép file `.json` qua máy khác bằng tay khi cần.

## 5. Cách chủ động lấy dữ liệu mới nhất

3 cách, dùng được cả 3 cùng lúc:

1. **Tự động lúc mở app** - `main.py` tự chạy đồng bộ trước khi nạp dữ
   liệu, không cần làm gì. Nhược điểm: lỗi (mất mạng, xung đột) sẽ bị bỏ
   qua âm thầm, không báo gì cho người dùng.
2. **Nút "🔄 Lấy dữ liệu mới nhất"** (góc trên màn hình chính, mọi tab) -
   thêm mới, dùng khi đang mở app mà muốn lấy ngay không cần tắt/mở lại.
   Báo rõ kết quả (thành công/lý do thất bại), và **đã kiểm tra: lấy đúng
   bản mới nhất kể cả khi máy này đã lỡ hiện bản cũ/sai trong cùng phiên.**
3. **"Cập nhật && Đăng lên web"** (tab CẬP NHẬT WEB) - ngoài việc đăng dữ
   liệu của chính máy đó, bước đầu tiên cũng tự đồng bộ về mới nhất trước
   khi đăng (xem `git.md` mục D).

## 6. Kiểm tra nhanh khi nghi ngờ dữ liệu chưa mới

| Hiện tượng | Kiểm tra |
|---|---|
| Bấm "🔄 Lấy dữ liệu mới nhất" báo lỗi "not a git repository" | Máy này chưa phải git clone thật - xem mục 3, cài lại theo [git.md](git.md) |
| Báo lỗi "Permission denied"/403 | Tài khoản GitHub trên máy này chưa có quyền push - xem [git.md mục B](git.md) |
| Báo "xung đột dữ liệu, không tự gộp được" | 2 máy sửa trùng đúng 1 dòng - xem [git.md mục D](git.md) |
| Ngày/tên/lịch học gán tay vẫn không đúng sau khi bấm 🔄 | Kiểm tra máy *kia* đã thật sự lưu + đăng thành công chưa (có thông báo "Hoàn tất" không); hoặc Excel đã có ngày/tên đúng rồi thì override không còn hiệu lực (Excel luôn ưu tiên hơn) |
| Học viên/phân xe (trong `data.json`) không khớp giữa 2 máy | Đây là hành vi đúng theo thiết kế hiện tại (mục 4) - cần tự chia sẻ lại file Excel, không phải lỗi |

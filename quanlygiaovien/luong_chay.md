# Luồng hoạt động của app "Quản lý giáo viên / học viên"

Tài liệu này mô tả app chạy như thế nào từ trong ra ngoài: cấu trúc thư
mục, luồng dữ liệu, từng bước trong pipeline đăng web, và luồng người
dùng cuối (giáo viên quét mã QR). Phần cuối đánh giá app đã sẵn sàng dùng
thực tế chưa.

## 1. Cấu trúc tổng thể

```
D:\GTCC\                        <-- git repo gốc (remote: github.com/luantt94/GTCC)
├── docs/                       <-- public, GitHub Pages phục vụ trực tiếp từ đây
│   ├── gv.html                 (trang xem thông tin giáo viên, tĩnh)
│   ├── public_data.json        (dữ liệu công khai, app tự sinh + git push)
│   └── files/                  (PDF/Word lịch học, giấy tờ xe, hồ sơ GV - bản copy công khai)
│
└── quanlygiaovien/             <-- app desktop (PySide6), KHÔNG public
    ├── main.py                 (điểm khởi động, dựng 4 tab)
    ├── tab_cap_nhat_web.py     (UI tab "CẬP NHẬT WEB" - nơi bấm nút đăng)
    ├── publish_data.py         (pipeline: đọc nguồn -> sinh data.json + public_data.json + QR)
    ├── convert_excel_to_json.py(đọc file "BIÊN CHẾ XE KHÓA *.xlsx" -> data.json)
    ├── roster_xml.py           (đọc DSGV_*.xml / DSXTL_*.xml)
    ├── attach_files.py         (ghép PDF/Word vào đúng khóa/xe/giáo viên theo tên file)
    ├── git_publish.py          (git add/commit/push thư mục docs/)
    ├── data_model.py           (tra cứu trong data.json cho 3 tab còn lại)
    ├── tab_tra_cuu.py / tab_phan_xe.py / tab_lich_hoc.py  (3 tab chỉ-đọc)
    ├── data.json               (dữ liệu ĐẦY ĐỦ, nội bộ, .gitignore - không public)
    ├── data_sources.json       (ghi nhớ đường dẫn nguồn đã chọn, đặc thù theo máy)
    ├── app_settings.json       (mật khẩu hồ sơ cá nhân, chỉ máy local biết bản gốc)
    ├── qr_codes/*.png          (ảnh QR từng giáo viên, chỉ lưu local để in)
    └── data/                   (Excel/XML/PDF/Word nguồn - .gitignore, KHÔNG public)
```

Điểm mấu chốt: **git root là `D:\GTCC`, không phải `quanlygiaovien`**.
`publish_data.py` tính `REPO_ROOT = APP_DIR.parent` nên `docs/` luôn nằm
ngoài thư mục app một cấp — đây là lý do GitHub Pages phục vụ được từ
nhánh `main`, thư mục `/docs`.

## 2. Luồng dữ liệu tổng quát (khi bấm "Cập nhật && Đăng lên web")

```
 Excel "BIÊN CHẾ XE KHÓA*.xlsx" ─┐
 PDF "LICH HOC THUC HANH*.pdf"  ─┤
 DSGV_*.xml (danh sách GV)      ─┼──► publish_data.run_pipeline() ──► data.json (nội bộ, full)
 DSXTL_*.xml (danh sách xe)     ─┤          │
 thư mục giấy tờ xe (.docx/.pdf)─┤          ├──► đối chiếu GV theo tên, xe theo biển số
 thư mục hồ sơ GV (.pdf)        ─┘          ├──► copy file PDF/Word liên quan -> docs/files/
                                            ├──► sinh/vẽ lại QR cho GV mới hoặc đổi tên -> qr_codes/
                                            └──► ghi docs/public_data.json (chỉ dữ liệu được phép public)
                                                        │
                                            git_publish.push_public_data()
                                                        │
                                   git add docs/public_data.json docs/files
                                   git commit -m "Cập nhật số liệu giáo viên (...)"
                                   git push origin main
                                                        │
                                          GitHub Pages tự build lại (~1-2 phút)
                                                        │
                        https://luantt94.github.io/GTCC/gv.html?id=<MaGV>
                                                        │
                        Giáo viên quét QR (cố định, in 1 lần) -> xem dữ liệu mới nhất
```

### Chi tiết các bước trong `publish_data.run_pipeline()`

1. **Đọc & validate nguồn** (`load_data_sources`): bắt buộc phải có ít
   nhất 1 file "BIÊN CHẾ XE KHÓA" và file `DSGV_*.xml`; thiếu thì raise
   lỗi ngay, UI hiện `QMessageBox`.
2. **Convert Excel → `data.json`** (`convert_excel_to_json.build_data`):
   mỗi file = 1 khóa, parse theo dòng tiêu đề "BIÊN CHẾ XE KHÓA ..." và
   bảng học viên (cột gộp xe/GV/SĐT). Hai file trùng mã khóa → file đọc
   sau đè file trước, trả về cảnh báo "mã khóa trùng".
3. **Parse DSGV / DSXTL** (`roster_xml`): chỉ lấy `MaGV + Họ tên + SĐT`
   (DSGV) và `biển số` (DSXTL) — chủ động KHÔNG đọc ảnh/CCCD.
4. **Dọn `docs/files/`**: xoá sạch rồi copy lại từ đầu, để không còn sót
   file PDF/Word của khóa/giáo viên không còn hoạt động.
5. **Với mỗi khóa đang "Đang đào tạo"** (`DataModel.status_of` so ngày
   hệ thống với `tu_ngay`/`den_ngay`):
   - Ghép lịch học PDF theo mã lớp (`attach_files.find_lich_hoc_pdf`).
   - Với mỗi nhóm học viên: khớp tên GV với DSGV
     (không khớp được → cảnh báo, **bỏ qua nhóm đó** khỏi public data);
     khớp biển số với DSXTL (không có → cảnh báo nhưng vẫn giữ);
     copy giấy tờ xe tương ứng.
   - Gom vào `teaching_now[ma_gv]`.
6. **Với mỗi giáo viên trong DSGV** (toàn bộ, không chỉ người đang dạy):
   - Ghép hồ sơ cá nhân PDF theo tên (`find_ho_so_gv`).
   - Nếu MaGV chưa có QR → sinh QR mới (`qr_codes/<MaGV>.png`), báo
     "giáo viên mới".
   - Nếu đã có QR nhưng tên trong DSGV đổi (so với `_captions.json`) →
     vẽ lại QR với tên mới, **không đổi link/mã QR**.
7. **Ghi `docs/public_data.json`**: tên GV, khóa đang dạy (mã khóa, lớp,
   thời gian, xe, SĐT, **tên** học viên, link PDF lịch học/giấy tờ xe),
   hồ sơ cá nhân (link file) + **hash SHA-256** của mật khẩu (nếu có đặt).
8. **Trả kết quả** cho UI: số GV, GV mới, GV đổi tên, cảnh báo, số khóa
   trùng → hiển thị trong khung log + popup.
9. **`git_publish.push_public_data()`**: `git add -A` đúng
   `docs/public_data.json` + `docs/files`, nếu có thay đổi mới `commit` +
   `push`. Nếu push lỗi (mất mạng/chưa đăng nhập) → báo rõ, dữ liệu vẫn
   đã lưu local, bấm lại nút khi có mạng là đủ (không mất gì).

## 3. Luồng mở app bình thường (3 tab tra cứu)

```
main.py khởi động
  → nếu chưa có data.json -> tạo file rỗng
  → DataModel(data.json)   # đọc 1 LẦN DUY NHẤT lúc mở app
  → 3 tab TRA CỨU / PHÂN XE / LỊCH HỌC dùng chung model này (chỉ đọc RAM)
```

→ **Hệ quả quan trọng**: sau khi bấm "Cập nhật && Đăng lên web" (ghi
`data.json` mới), 3 tab này **không tự refresh** — phải tắt/mở lại app.
Đây là hành vi đã biết và được ghi rõ trong README, không phải lỗi.

## 4. Luồng người dùng cuối (giáo viên quét mã QR)

```
Giáo viên mở camera, quét mã QR in sẵn
  → mở https://luantt94.github.io/GTCC/gv.html?id=<MaGV>  (link cố định, không đổi)
  → gv.html (trang tĩnh) fetch docs/public_data.json (no-store, luôn lấy bản mới nhất)
  → tìm teachers[id]; không có -> "Mã QR không hợp lệ"
  → render: tên, số khóa đang dạy, tổng học viên
  → mỗi khóa là 1 card, bấm vào từng dòng mới load chi tiết (lazy):
      - Xe tập lái -> SĐT + giấy tờ xe (nhúng iframe PDF.js/Office Viewer)
      - Số học viên -> danh sách tên
      - Lịch học -> bảng + PDF gốc
  → Hồ sơ cá nhân (nếu có đặt mật khẩu): bấm -> prompt() nhập mật khẩu ->
    so sánh SHA-256 ngay trên trình duyệt -> đúng mới hiện file
    (⚠️ chỉ là lớp che, không phải bảo mật thật - xem mục 2 README)
```

Đã kiểm tra thực tế lúc viết tài liệu này: trang
`https://luantt94.github.io/GTCC/gv.html?id=48002001` và
`public_data.json` đang **chạy thật trên Internet**, có 64 giáo viên,
dữ liệu mới nhất cập nhật lúc `2026-10-28T21:24:36` — tức là app đã
được dùng và đăng dữ liệu thành công ít nhất 1 lần trong thực tế.

## 5. App đã sẵn sàng dùng thực tế chưa?

**Đã dùng được, và thực tế đang chạy** — không phải bản demo: trang web
public đang sống, có dữ liệu 64 giáo viên, lịch sử 6 commit cho thấy đã
qua vài vòng sửa lỗi thực tế (crash khi xem lịch học, lỗi hiển thị PDF
trong app nhắn tin, escape HTML...). Với quy mô 1 trung tâm, vài chục
giáo viên, do 1-2 người vận hành, app **dùng được ngay cho công việc
hàng ngày/hàng tuần** như README mục 3 mô tả.

Tuy vậy, đây vẫn là app nội bộ quy mô nhỏ, viết nhanh để giải quyết đúng
nhu cầu hiện tại — còn vài điểm nên biết trước khi coi là "sản phẩm sẵn
sàng vận hành lâu dài, nhiều người dùng":

| Rủi ro | Mức độ | Chi tiết |
|---|---|---|
| **Không có tự động kiểm thử (test)** | Trung bình | Không có unit test nào trong repo. Mỗi lần sửa code (vd đổi cách đọc Excel) chỉ được kiểm tra bằng cách chạy tay. Nếu phía Sở GTVT đổi định dạng export DSGV/DSXTL hoặc đổi template Excel, app có thể lỗi mà không ai phát hiện kịp ngoài khi chạy thật. |
| **Khớp dữ liệu bằng so khớp tên file / tiêu đề chuỗi** | Trung bình | `attach_files.py` và `convert_excel_to_json.py` dựa vào quy ước đặt tên/văn bản tiếng Việt cố định ("BIÊN CHẾ XE KHÓA...", biển số, tên GV viết tắt). Đã có cơ chế cảnh báo khi không khớp (tốt), nhưng vẫn cần người đọc kỹ khung log mỗi lần đăng — dễ bị bỏ qua nếu làm vội. |
| **3 tab tra cứu không tự refresh** | Thấp | Đã biết, đã ghi rõ trong README, nhưng dễ gây thông tin sai nếu nhân viên quên khởi động lại app sau khi đăng web. |
| **Bảo mật hồ sơ cá nhân chỉ là "che tượng trưng"** | Trung bình-Cao (tuỳ mức nhạy cảm hồ sơ) | Chính README đã cảnh báo rất rõ: link PDF hồ sơ (có thể có thông tin định danh) nằm sẵn trong `public_data.json`, ai mở DevTools cũng xem được, không cần mật khẩu. Nếu hồ sơ GV có dữ liệu cần bảo vệ thật (không chỉ "che mắt" người dùng thường), cần đánh giá lại có nên public kiểu này không. |
| **Không có cơ chế chống ghi đè đồng thời** | Thấp (vì hiện chỉ 1 người bấm nút) | Nếu sau này 2 máy/2 người cùng bấm "Đăng lên web" gần nhau, `git push` có thể bị reject do không fast-forward — code hiện tại không tự `pull`/`rebase`, chỉ báo lỗi chung, người dùng phải biết cách xử lý git xung đột (ngoài tầm UI của app). |
| **Không có backup dữ liệu nguồn ngoài máy local** | Trung bình | Thư mục `data/` (Excel/XML/PDF/Word gốc) bị `.gitignore` hoàn toàn — chỉ tồn tại trên máy chạy app. Nếu máy hỏng/mất, toàn bộ dữ liệu nguồn (trừ phần đã lên web public) sẽ mất, không có bản sao ở git hay nơi khác. |
| **Dữ liệu học viên (tên) hiển thị công khai trên Internet** | Cần xác nhận pháp lý | Đây là quyết định sản phẩm có chủ đích (ghi trong README mục 5), nhưng nên xác nhận lại với yêu cầu thực tế của trung tâm/quy định hiện hành về bảo vệ dữ liệu cá nhân trước khi dùng rộng, vì tên học viên đang public không cần đăng nhập. |

**Kết luận**: app đủ ổn để tiếp tục dùng cho quy trình hiện tại (quy mô
nhỏ, người vận hành hiểu rõ quy trình, đã quen đọc cảnh báo trong log).
Trước khi mở rộng (nhiều người dùng hơn, dữ liệu nhạy cảm hơn, hoặc coi
đây là hệ thống "phải luôn chạy đúng"), nên ưu tiên xử lý theo thứ tự:
1. Xác nhận lại mức độ chấp nhận được của việc public tên học viên +
   cách "bảo mật" hồ sơ cá nhân (mục đã nêu rõ trong README, không phải
   phát hiện mới — nhưng cần người có thẩm quyền xác nhận chấp nhận rủi
   ro này).
2. Có bản backup định kỳ (thủ công cũng được) cho thư mục `data/` ở một
   nơi khác ngoài máy chạy app.
3. Thêm vài test tự động cho phần parse Excel/XML (phần dễ vỡ nhất khi
   nguồn dữ liệu đổi định dạng).

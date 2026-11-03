"""Pipeline cập nhật dữ liệu công khai cho trang web (GitHub Pages):

1. Đọc các nguồn trong data_sources.json (xem load_data_sources()):
   - File Excel "BIÊN CHẾ XE KHÓA xxxx.xlsx/.xls" (mỗi khóa 1 file)
   - File PDF "LICH HOC THUC HANH xxxx.pdf" (mỗi khóa 1 file)
   - File DSGV_*.xml (danh sách toàn bộ giáo viên, nguồn ID chính thức)
   - File DSXTL_*.xml (danh sách xe tập lái, dùng để đối chiếu biển số)
   - Thư mục hồ sơ cá nhân giáo viên (.pdf)
2. Convert Excel -> data.json (dữ liệu đầy đủ, chỉ dùng nội bộ, KHÔNG commit lên git)
3. Build docs/public_data.json: tên/SĐT/ngày sinh/địa chỉ giáo viên (theo
   yêu cầu - lấy từ DSGV_*.xml) + chi tiết từng khóa đang đào tạo (khóa/lớp,
   thời gian, xe, SĐT, danh sách tên học viên) + file lịch học PDF + link
   giấy tờ xe (nhập tay, xem xe_extra.json - KHÔNG còn nhúng file PDF/Word
   giấy tờ xe lên trang web nữa)
4. Sinh mã QR cho MỌI giáo viên trong DSGV (mã QR = MaGV - mã chính thức, cố
   định, không bao giờ trùng lặp) vào qr_codes/<MaGV>.png, có ghi tên bên dưới.
   Link/mã QR không bao giờ đổi; nếu DSGV sửa lại tên 1 giáo viên đã có QR,
   app tự nhận ra (so với qr_codes/_captions.json) và vẽ lại ảnh với tên mới,
   không cần xóa tay.

Lưu ý: public_data.json có TÊN + kết quả thi/ngày tốt nghiệp (nếu có) của
HỌC VIÊN (theo yêu cầu) nhưng KHÔNG có ngày sinh / địa chỉ / CMT / ảnh của
HỌC VIÊN - những trường đó không bao giờ được đọc ra khỏi data.json nội bộ
(chỉ lấy tên + điện thoại + ghi chú + kết quả thi + ngày tốt nghiệp). SĐT/ngày
sinh/địa chỉ CỦA GIÁO VIÊN thì CÓ public (theo yêu cầu) - riêng ảnh chân
dung (AnhCD) và số CMT/CCCD (SoCMT) trong DSGV thì không bao giờ đọc, kể
cả của giáo viên.

Không đụng tới git ở đây - phần git do git_publish.py xử lý.
"""
import hashlib
import json
import shutil
import unicodedata
from pathlib import Path

import attach_files
import convert_excel_to_json
import git_publish
import roster_xml
from data_model import DataModel
from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent  # docs/ phải nằm ở root repo để GitHub Pages phục vụ được
DATA_DIR = APP_DIR / "data"
DOCS_DIR = REPO_ROOT / "docs"
PUBLIC_DATA_FILE = DOCS_DIR / "public_data.json"
FILES_OUT_DIR = DOCS_DIR / "files"
QR_DIR = APP_DIR / "qr_codes"
QR_CAPTIONS_FILE = QR_DIR / "_captions.json"
APP_SETTINGS_FILE = APP_DIR / "app_settings.json"
DATA_SOURCES_FILE = APP_DIR / "data_sources.json"
PHOTOS_DIR = APP_DIR / "anh_gv"  # ảnh đại diện giáo viên - lưu local, bản resize sẽ copy sang docs/files
GV_EXTRA_FILE = APP_DIR / "giao_vien_extra.json"  # hạng lái xe dạy - tách khỏi DSGV, lưu local theo MaGV
XE_EXTRA_FILE = APP_DIR / "xe_extra.json"  # link giấy tờ xe (nhập tay) - lưu local theo biển số chuẩn hoá

BASE_URL = "https://luantt94.github.io/GTCC/gv.html"

DEFAULT_XE_FOLDER_NAME = "Xe tập lái"
DEFAULT_HO_SO_FOLDER_NAME = "HO SO GV (2026_GTCC)"

# Mật khẩu xem hồ sơ cá nhân mặc định (máy mới chưa từng lưu cài đặt riêng).
DEFAULT_HO_SO_PASSWORD = "gtcc2026"

ANH_GV_SUBDIR = "anh_gv"


# ---------------- cài đặt chung (mật khẩu hồ sơ cá nhân) ----------------
def load_app_settings():
    if APP_SETTINGS_FILE.exists():
        return json.loads(APP_SETTINGS_FILE.read_text(encoding="utf-8"))
    # Máy mới chưa từng lưu cài đặt - mặc định đặt sẵn mật khẩu hồ sơ cá
    # nhân (người dùng vẫn đổi/xóa được ở tab CẬP NHẬT WEB như bình thường).
    return {"ho_so_password": DEFAULT_HO_SO_PASSWORD}


def save_app_settings(settings):
    APP_SETTINGS_FILE.write_text(
        json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ---------------- cột ẩn/hiện đã chọn trong dialog "Danh sách chi tiết" ----------------
# `key` phân biệt giáo viên/xe (vd "giao_vien", "xe"). Trả None nếu người
# dùng chưa từng tự chọn cho key này - khi đó dialog tự tính mặc định (ẩn
# cột không có dữ liệu) thay vì coi là "ẩn tất cả".
def get_detail_dialog_hidden_columns(key):
    settings = load_app_settings()
    return settings.get("detail_dialog_hidden_columns", {}).get(key)


def set_detail_dialog_hidden_columns(key, hidden_columns):
    settings = load_app_settings()
    hidden_map = settings.setdefault("detail_dialog_hidden_columns", {})
    hidden_map[key] = list(hidden_columns)
    save_app_settings(settings)


# ---------------- thông tin riêng từng giáo viên (hạng lái xe dạy) ----------------
# Lưu local theo MaGV, tách khỏi DSGV/DSXTL (những file này lấy từ nguồn của
# Sở và có thể bị thay thế bất cứ lúc nào) nên không bị mất khi đổi file XML,
# miễn MaGV của giáo viên đó không đổi.
def load_giao_vien_extra():
    if GV_EXTRA_FILE.exists():
        return json.loads(GV_EXTRA_FILE.read_text(encoding="utf-8"))
    return {}


def save_giao_vien_extra(data):
    GV_EXTRA_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def get_hang_lai_xe(ma_gv):
    return load_giao_vien_extra().get(ma_gv, {}).get("hang", "")


def set_hang_lai_xe(ma_gv, hang):
    data = load_giao_vien_extra()
    rec = data.setdefault(ma_gv, {})
    rec["hang"] = (hang or "").strip()
    save_giao_vien_extra(data)


# ---------------- thông tin riêng từng xe (link giấy tờ xe) ----------------
# Trang web KHÔNG còn nhúng file PDF/Word giấy tờ xe nữa - thay bằng 1 link
# do người dùng tự nhập (vd trang tra cứu của Sở), lưu local theo biển số đã
# chuẩn hoá (attach_files.normalize_plate) - tách khỏi DSXTL.
def load_xe_extra():
    if XE_EXTRA_FILE.exists():
        return json.loads(XE_EXTRA_FILE.read_text(encoding="utf-8"))
    return {}


def save_xe_extra(data):
    XE_EXTRA_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def get_xe_link(plate_norm):
    return load_xe_extra().get(plate_norm, {}).get("link", "")


def set_xe_link(plate_norm, link):
    data = load_xe_extra()
    rec = data.setdefault(plate_norm, {})
    rec["link"] = (link or "").strip()
    save_xe_extra(data)


# ---------------- nguồn dữ liệu (data_sources.json) ----------------
def _latest_by_mtime(paths):
    """Chọn file được sửa đổi gần đây nhất (đáng tin hơn sort theo tên khi
    định dạng tên file xuất ra có thể thay đổi)."""
    paths = list(paths)
    return max(paths, key=lambda p: p.stat().st_mtime) if paths else None


def _scan_default_sources():
    """Tự động quét thư mục data/ để gợi ý sẵn danh sách nguồn lần đầu mở app."""
    bien_che = sorted(
        str(p)
        for p in list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.xls"))
        if "BIÊN CHẾ XE KHÓA" in p.name.upper()
    )
    lich_hoc = sorted(str(p) for p in DATA_DIR.glob("*.pdf") if "LICH HOC" in p.name.upper())
    dsgv = _latest_by_mtime(DATA_DIR.glob("DSGV_*.xml"))
    dsxtl = _latest_by_mtime(DATA_DIR.glob("DSXTL_*.xml"))
    ket_qua_thi = _latest_by_mtime(
        p
        for p in list(DATA_DIR.glob("*.xlsx")) + list(DATA_DIR.glob("*.xls"))
        if "THI ĐẠT TỐT NGHIỆP" in p.name.upper()
    )
    xe_folder = DATA_DIR / DEFAULT_XE_FOLDER_NAME
    ho_so_folder = DATA_DIR / DEFAULT_HO_SO_FOLDER_NAME
    return {
        "bien_che_files": bien_che,
        "lich_hoc_files": lich_hoc,
        "dsgv_file": str(dsgv) if dsgv else "",
        "dsxtl_file": str(dsxtl) if dsxtl else "",
        "ket_qua_thi_file": str(ket_qua_thi) if ket_qua_thi else "",
        "xe_folder": str(xe_folder) if xe_folder.is_dir() else "",
        "ho_so_folder": str(ho_so_folder) if ho_so_folder.is_dir() else "",
    }


def load_data_sources():
    """Danh sách nguồn dữ liệu đã chọn (lưu lại giữa các lần mở app).
    Lần đầu (chưa có file cấu hình) sẽ tự quét thư mục data/ để gợi ý sẵn."""
    if DATA_SOURCES_FILE.exists():
        saved = json.loads(DATA_SOURCES_FILE.read_text(encoding="utf-8"))
        saved["bien_che_files"] = [p for p in saved.get("bien_che_files", []) if Path(p).exists()]
        saved["lich_hoc_files"] = [p for p in saved.get("lich_hoc_files", []) if Path(p).exists()]
        return saved
    return _scan_default_sources()


def save_data_sources(sources):
    DATA_SOURCES_FILE.write_text(
        json.dumps(sources, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ---------------- QR ----------------
def _load_caption_font(size):
    import os

    from PIL import ImageFont

    windir = os.environ.get("WINDIR", "C:\\Windows")
    for name in ("segoeui.ttf", "arial.ttf"):
        path = os.path.join(windir, "Fonts", name)
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def _paste_photo_center(qr_img, photo_path, ratio=0.26):
    """Dán ảnh đại diện giáo viên (cắt tròn, viền trắng) vào chính giữa mã QR.
    Mã QR phải dùng error_correction mức H (30%) để vẫn đọc được khi bị che
    1 phần ở giữa (xem _generate_qr)."""
    from PIL import Image, ImageDraw

    photo = Image.open(photo_path).convert("RGB")
    side = min(photo.size)
    left = (photo.width - side) // 2
    top = (photo.height - side) // 2
    size = int(qr_img.width * ratio)
    photo = photo.crop((left, top, left + side, top + side)).resize(
        (size, size), Image.LANCZOS
    )

    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, size, size), fill=255)

    border = max(size // 12, 4)
    badge_size = size + border * 2
    badge = Image.new("RGBA", (badge_size, badge_size), (0, 0, 0, 0))
    ImageDraw.Draw(badge).ellipse((0, 0, badge_size, badge_size), fill="white")
    badge.paste(photo, (border, border), mask)

    canvas = qr_img.convert("RGBA")
    x = (canvas.width - badge_size) // 2
    y = (canvas.height - badge_size) // 2
    canvas.paste(badge, (x, y), badge)
    return canvas.convert("RGB")


def _generate_qr(ma_gv, ten):
    """Tạo mã QR (link cố định theo MaGV) kèm tên giáo viên in bên dưới, để
    dễ phân biệt các mã QR đã in. Nếu giáo viên đã có ảnh đại diện, dán ảnh
    (cắt tròn) vào chính giữa mã QR cho đẹp - dùng error_correction mức H để
    mã QR vẫn đọc được dù bị che giữa."""
    import qrcode
    from qrcode.constants import ERROR_CORRECT_H
    from PIL import Image, ImageDraw

    QR_DIR.mkdir(exist_ok=True)
    url = f"{BASE_URL}?id={ma_gv}"
    qr = qrcode.QRCode(error_correction=ERROR_CORRECT_H, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    photo_path = teacher_photo_path(ma_gv)
    if photo_path:
        qr_img = _paste_photo_center(qr_img, photo_path)

    font = _load_caption_font(20)
    caption_h = 34
    canvas = Image.new("RGB", (qr_img.width, qr_img.height + caption_h), "white")
    canvas.paste(qr_img, (0, 0))
    draw = ImageDraw.Draw(canvas)
    bbox = draw.textbbox((0, 0), ten, font=font)
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    x = max((canvas.width - text_w) // 2, 0)
    y = qr_img.height + max((caption_h - text_h) // 2, 0)
    draw.text((x, y), ten, fill="black", font=font)
    canvas.save(QR_DIR / f"{ma_gv}.png")


def _load_qr_captions():
    if QR_CAPTIONS_FILE.exists():
        return json.loads(QR_CAPTIONS_FILE.read_text(encoding="utf-8"))
    return {}


def _save_qr_captions(captions):
    QR_CAPTIONS_FILE.write_text(
        json.dumps(captions, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _copy_public_file(src_path, rel_subdir, dest_stem):
    """Copy 1 file gốc vào docs/files/<rel_subdir>/<dest_stem><ext>.
    Trả về đường dẫn tương đối (vd "files/xe/43A38936.docx") để nhúng vào
    public_data.json, hoặc None nếu src_path rỗng."""
    if not src_path:
        return None
    dest_dir = FILES_OUT_DIR / rel_subdir
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest_path = dest_dir / f"{dest_stem}{src_path.suffix.lower()}"
    shutil.copyfile(src_path, dest_path)
    return f"files/{rel_subdir}/{dest_path.name}"


def _is_bao_luu(ghi_chu):
    """So khớp ghi chú học viên với "Bảo lưu", không phân biệt hoa/thường
    hay có dấu/không dấu (vd "bao luu", "BẢO LƯU", "Bảo Lưu" đều khớp)."""
    text = unicodedata.normalize("NFD", ghi_chu or "")
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = " ".join(text.split()).strip().upper()
    return text == "BAO LUU"


def run_pipeline(sources=None):
    """Chạy toàn bộ pipeline theo data_sources.json (hoặc sources truyền vào).
    Trả về dict kết quả để hiển thị cho người dùng."""
    if sources is None:
        sources = load_data_sources()

    bien_che_files = sources.get("bien_che_files") or []
    if not bien_che_files:
        raise FileNotFoundError(
            "Chưa chọn file Excel 'BIÊN CHẾ XE KHÓA ...' nào."
        )
    for p in bien_che_files:
        if not Path(p).exists():
            raise FileNotFoundError(f"Không tìm thấy file Excel: {p}")

    dsgv_file = sources.get("dsgv_file") or ""
    if not dsgv_file or not Path(dsgv_file).exists():
        raise FileNotFoundError(
            "Chưa chọn file danh sách giáo viên (DSGV_*.xml). File này là "
            "nguồn mã QR chính thức cho giáo viên, bắt buộc phải có."
        )

    ket_qua_thi_file = sources.get("ket_qua_thi_file") or ""
    ket_qua_thi = (
        convert_excel_to_json.load_ket_qua_thi(ket_qua_thi_file)
        if ket_qua_thi_file and Path(ket_qua_thi_file).exists()
        else {}
    )

    data_json_path = APP_DIR / "data.json"
    data, overwritten = convert_excel_to_json.build_data(bien_che_files, ket_qua_thi=ket_qua_thi)
    data_json_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    model = DataModel(data_json_path)
    dsgv = roster_xml.parse_dsgv(dsgv_file)
    dsgv_file_name = Path(dsgv_file).name

    dsxtl_file = sources.get("dsxtl_file") or ""
    dsxtl = roster_xml.parse_dsxtl(dsxtl_file) if dsxtl_file and Path(dsxtl_file).exists() else {}
    dsxtl_plates = {attach_files.normalize_plate(p) for p in dsxtl}
    dsxtl_file_name = Path(dsxtl_file).name if dsxtl_file else ""

    lich_hoc_files = sources.get("lich_hoc_files") or []
    ho_so_folder = sources.get("ho_so_folder") or ""
    xe_extra = load_xe_extra()

    name_to_magv = {attach_files.normalize_name(rec["ten"]): magv for magv, rec in dsgv.items()}

    warnings = []

    # Dọn sạch file public cũ (PDF/Word) trước khi gắn lại cho đúng danh sách
    # giáo viên/khóa đang active hiện tại (tránh file cũ của khóa/giáo viên
    # không còn dạy nằm lại trên web public).
    shutil.rmtree(FILES_OUT_DIR, ignore_errors=True)

    lich_hoc_pdf_cache = {}  # ma_khoa -> rel path (hoặc None), tránh copy trùng
    teaching_now = {}  # ma_gv -> [khoa...]
    teacher_phone = {}  # ma_gv -> SĐT (lấy từ cột SĐT trong Excel biên chế xe,
    # khóa nào có nhập trước thì lấy khóa đó) - hiện ở mục "Thông tin giáo
    # viên", không còn hiện riêng trong từng khóa (mục "Xe tập lái") nữa.

    for c in model.all_courses():
        if model.status_of(c) != "Đang đào tạo":
            continue
        ma_khoa = c["ma_khoa"]
        lop = c.get("lop", "")
        if ma_khoa not in lich_hoc_pdf_cache:
            src = attach_files.find_lich_hoc_pdf(lop, lich_hoc_files)
            lich_hoc_pdf_cache[ma_khoa] = _copy_public_file(src, "lich_hoc", ma_khoa or lop)
        lich_hoc_pdf = lich_hoc_pdf_cache[ma_khoa]

        for g in c.get("groups", []):
            name = g.get("giao_vien")
            if not name:
                continue
            ma_gv = name_to_magv.get(attach_files.normalize_name(name))
            if not ma_gv:
                # Không khớp hệt - thử nhận diện tên viết tắt (vd "Trần N. Đ.
                # Phương" -> "Trần Ngô Đông Phương") trước khi báo không khớp.
                candidates = attach_files.find_teacher_candidates(name, dsgv)
                if len(candidates) == 1:
                    ma_gv, ten_day_du = candidates[0]
                    warnings.append(
                        f"Tên viết tắt: đã tự nhận '{name}' trong file Excel "
                        f"'{c.get('source_file', '')}' (khóa {ma_khoa}) là '{ten_day_du}' "
                        f"trong file '{dsgv_file_name}' - nên sửa lại Excel cho đúng chính "
                        "tả để khỏi phải đoán."
                    )
                elif candidates:
                    goi_y = "; ".join(f"'{ten}'" for _, ten in candidates)
                    warnings.append(
                        f"Không khớp được giáo viên '{name}' trong file Excel "
                        f"'{c.get('source_file', '')}' (khóa {ma_khoa}) với file "
                        f"'{dsgv_file_name}' - tên viết tắt này khớp MƠ HỒ với nhiều giáo "
                        f"viên ({goi_y}), hãy sửa lại Excel cho rõ."
                    )
                    continue
                else:
                    warnings.append(
                        f"Không khớp được giáo viên '{name}' trong file Excel "
                        f"'{c.get('source_file', '')}' (khóa {ma_khoa}) với danh sách "
                        f"giáo viên trong file '{dsgv_file_name}' - kiểm tra lại chính tả tên."
                    )
                    continue

            hoc_vien = g.get("hoc_vien", [])
            xe = g.get("xe", "")
            if dsxtl_plates and xe and attach_files.normalize_plate(xe) not in dsxtl_plates:
                warnings.append(
                    f"Biển số '{xe}' trong file Excel '{c.get('source_file', '')}' "
                    f"(GV {name}, khóa {ma_khoa}) không có trong file '{dsxtl_file_name}'."
                )
            giay_to_xe_link = xe_extra.get(attach_files.normalize_plate(xe), {}).get("link") or None

            sdt = g.get("sdt", "")
            if sdt and ma_gv not in teacher_phone:
                teacher_phone[ma_gv] = sdt

            hoc_vien_display = []
            so_hoc_vien_dem = 0
            for hv in hoc_vien:
                ghi_chu = (hv.get("ghi_chu") or "").strip()
                ten_hv = f"{hv['ho_ten']} ({ghi_chu})" if ghi_chu else hv["ho_ten"]
                hoc_vien_display.append(
                    {
                        "ten": ten_hv,
                        # Chỉ công khai kết quả thi (Đạt/chưa) + ngày tốt nghiệp -
                        # KHÔNG đưa CMND/ngày sinh/địa chỉ học viên lên web public.
                        "ket_qua_thi": hv.get("ket_qua_thi", ""),
                        "ngay_tot_nghiep": hv.get("ngay_tot_nghiep", ""),
                    }
                )
                if not _is_bao_luu(ghi_chu):
                    so_hoc_vien_dem += 1

            teaching_now.setdefault(ma_gv, []).append(
                {
                    "ma_khoa": ma_khoa,
                    "lop": lop,
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                    "xe": xe,
                    "so_hoc_vien": so_hoc_vien_dem,
                    "hoc_vien": hoc_vien_display,
                    "lich_hoc_pdf": lich_hoc_pdf,
                    "giay_to_xe_link": giay_to_xe_link,
                }
            )

    DOCS_DIR.mkdir(exist_ok=True)
    QR_DIR.mkdir(exist_ok=True)
    existing_qr = {f.stem for f in QR_DIR.glob("*.png")}
    qr_captions = _load_qr_captions()
    giao_vien_moi = []
    giao_vien_doi_ten = []

    teachers_out = {}
    for ma_gv, rec in dsgv.items():
        ho_so_src = attach_files.find_ho_so_gv(rec["ten"], ho_so_folder)
        ho_so_file = _copy_public_file(ho_so_src, "ho_so_gv", ma_gv)
        anh_src = PHOTOS_DIR / f"{ma_gv}.jpg"
        anh_file = _copy_public_file(anh_src, ANH_GV_SUBDIR, ma_gv) if anh_src.exists() else None
        teachers_out[ma_gv] = {
            "ten": rec["ten"],
            "dien_thoai": teacher_phone.get(ma_gv) or rec.get("dien_thoai", ""),
            "ngay_sinh": rec.get("ngay_sinh", ""),
            "dia_chi": rec.get("dia_chi", ""),
            "hang_lai_xe": get_hang_lai_xe(ma_gv),
            "khoa_dang_day": teaching_now.get(ma_gv, []),
            "ho_so_file": ho_so_file,
            "anh_file": anh_file,
        }
        if ma_gv not in existing_qr:
            _generate_qr(ma_gv, rec["ten"])
            qr_captions[ma_gv] = rec["ten"]
            giao_vien_moi.append((rec["ten"], ma_gv))
        else:
            ten_cu = qr_captions.get(ma_gv)
            if ten_cu is None:
                # QR có từ trước khi có cơ chế theo dõi tên - coi như đang khớp,
                # không cần tạo lại.
                qr_captions[ma_gv] = rec["ten"]
            elif ten_cu != rec["ten"]:
                _generate_qr(ma_gv, rec["ten"])
                giao_vien_doi_ten.append((ten_cu, rec["ten"], ma_gv))
                qr_captions[ma_gv] = rec["ten"]

    _save_qr_captions(qr_captions)

    app_settings = load_app_settings()
    ho_so_password = (app_settings.get("ho_so_password") or "").strip()
    ho_so_password_hash = (
        hashlib.sha256(ho_so_password.encode("utf-8")).hexdigest()
        if ho_so_password
        else None
    )

    out = {
        "generated_at": model.generated_at,
        "teachers": teachers_out,
        "ho_so_password_hash": ho_so_password_hash,
    }
    PUBLIC_DATA_FILE.write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    return {
        "so_giao_vien": len(teachers_out),
        "giao_vien_moi": giao_vien_moi,
        "giao_vien_doi_ten": giao_vien_doi_ten,
        "generated_at": model.generated_at,
        "ma_khoa_trung": overwritten,
        "so_file_excel": len(bien_che_files),
        "canh_bao": warnings,
    }


# ---------------- ảnh đại diện giáo viên ----------------
# Thêm/xóa ảnh không cần chạy lại toàn bộ pipeline - sửa trực tiếp
# docs/public_data.json + docs/files/anh_gv rồi đăng (git push) ngay. Ảnh
# gốc (đã resize) còn giữ ở PHOTOS_DIR để lần "Cập nhật && Đăng lên web"
# sau vẫn tự copy lại đúng ảnh này (run_pipeline xóa sạch docs/files/ mỗi
# lần chạy, xem shutil.rmtree ở trên).
def _load_public_data():
    if PUBLIC_DATA_FILE.exists():
        return json.loads(PUBLIC_DATA_FILE.read_text(encoding="utf-8"))
    return {"generated_at": "", "teachers": {}, "ho_so_password_hash": None}


def _save_public_data(data):
    DOCS_DIR.mkdir(exist_ok=True)
    PUBLIC_DATA_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def _save_resized_photo(src_path, dest_path, max_dim=1000):
    from PIL import Image

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    img = Image.open(src_path).convert("RGB")
    if max(img.size) > max_dim:
        img.thumbnail((max_dim, max_dim), Image.LANCZOS)
    img.save(dest_path, "JPEG", quality=85)


def teacher_photo_path(ma_gv):
    """Đường dẫn ảnh local hiện tại của 1 giáo viên (None nếu chưa có ảnh)."""
    path = PHOTOS_DIR / f"{ma_gv}.jpg"
    return path if path.exists() else None


def add_teacher_photo(ma_gv, ten, src_path):
    """Thêm/đổi ảnh đại diện 1 giáo viên và đăng ngay lên web.
    Trả về (ok: bool, message: str)."""
    PHOTOS_DIR.mkdir(exist_ok=True)
    local_path = PHOTOS_DIR / f"{ma_gv}.jpg"
    _save_resized_photo(src_path, local_path)

    rel_path = _copy_public_file(local_path, ANH_GV_SUBDIR, ma_gv)
    _generate_qr(ma_gv, ten)  # vẽ lại QR, dán ảnh mới vào giữa

    data = _load_public_data()
    teachers = data.setdefault("teachers", {})
    rec = teachers.setdefault(ma_gv, {"ten": ten, "khoa_dang_day": [], "ho_so_file": None})
    rec["anh_file"] = rel_path
    _save_public_data(data)

    return git_publish.push_public_data(f"Thêm ảnh giáo viên {ten} ({ma_gv})")


def remove_teacher_photo(ma_gv, ten):
    """Xóa ảnh đại diện 1 giáo viên (local + public) và đăng ngay lên web.
    Trả về (ok: bool, message: str)."""
    local_path = PHOTOS_DIR / f"{ma_gv}.jpg"
    if local_path.exists():
        local_path.unlink()
    public_path = FILES_OUT_DIR / ANH_GV_SUBDIR / f"{ma_gv}.jpg"
    if public_path.exists():
        public_path.unlink()
    if (QR_DIR / f"{ma_gv}.png").exists():
        _generate_qr(ma_gv, ten)  # vẽ lại QR không còn ảnh giữa

    data = _load_public_data()
    teachers = data.get("teachers", {})
    if ma_gv in teachers:
        teachers[ma_gv]["anh_file"] = None
        _save_public_data(data)

    return git_publish.push_public_data(f"Xóa ảnh giáo viên {ten} ({ma_gv})")


# ---------------- thẻ (phù hiệu) giáo viên dạy lái xe ----------------
# Chỉ in ấn cục bộ (không đăng lên web) - ảnh + hạng lái xe dạy lấy từ
# anh_gv/ và giao_vien_extra.json, KHÔNG đọc trực tiếp từ DSGV/DSXTL.
def build_card_data(ma_gv, rec, hang_override=None):
    """Dựng teacher_card.CardData cho 1 giáo viên (rec như roster_xml.parse_dsgv
    trả về). `hang_override` dùng để xem trước hạng chưa lưu (xem
    render_teacher_card_preview) - None thì lấy hạng đã lưu trong
    giao_vien_extra.json."""
    import teacher_card

    settings = load_app_settings()
    hang = hang_override
    if hang is None:
        hang = load_giao_vien_extra().get(ma_gv, {}).get("hang", "")

    return teacher_card.CardData(
        ma_gv=ma_gv,
        ten=rec.get("ten", ""),
        hang=hang,
        photo_path=teacher_photo_path(ma_gv),
        org_line1=settings.get("the_org_line1", ""),
        org_line2=settings.get("the_org_line2", ""),
        url=f"{BASE_URL}?id={ma_gv}",
        ngay_sinh=rec.get("ngay_sinh", ""),
        dia_chi=rec.get("dia_chi", ""),
    )


def render_teacher_card_preview(ma_gv, rec, hang_override=None):
    """Vẽ thẻ (ảnh PIL) của 1 giáo viên để xem trước trong app - không xuất
    file. Dùng chung logic/toạ độ với export_teacher_cards (teacher_card.py)."""
    import teacher_card

    return teacher_card.generate_card_image(build_card_data(ma_gv, rec, hang_override))


def export_teacher_cards(teachers, output_path):
    """teachers: list (ma_gv, rec) với rec là dict như roster_xml.parse_dsgv
    trả về (ít nhất có "ten", có thể kèm "ngay_sinh"/"dia_chi"). Xuất file
    PDF "Phù hiệu giáo viên dạy lái xe" (Mẫu số 01), tối đa 8 thẻ/trang.
    Trả về số trang đã xuất."""
    import teacher_card

    items = [build_card_data(ma_gv, rec) for ma_gv, rec in teachers]
    return teacher_card.generate_cards_pdf(items, output_path)

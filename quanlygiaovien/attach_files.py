"""Tìm & gắn file PDF/Word gốc (lịch học, giấy tờ xe, hồ sơ giáo viên) vào
gói dữ liệu public. Tất cả các hàm chỉ KHỚP theo TÊN FILE (không đổi tên
file gốc) nên không cần sửa tay khi có file mới, miễn tên file theo đúng
quy ước đã dùng từ trước tới nay:

- Lịch học: "lop" (vd "BK0726") là 1 chuỗi con (so khớp không phân biệt
  hoa/thường, bỏ dấu cách/ký tự đặc biệt) trong tên file PDF
  (vd "LICH HOC THUC HANH BK0726.signed.pdf").
- Giấy tờ xe: biển số (vd "43A-389.36") sau khi bỏ hết ký tự không phải
  chữ/số -> "43A38936" -> là 1 chuỗi con trong tên file.
- Hồ sơ giáo viên: tên giáo viên (bỏ số thứ tự đầu file, chuẩn hoá khoảng
  trắng, so khớp không phân biệt hoa/thường).

Không tìm thấy file nào khớp thì trả về None, không báo lỗi.
"""
import re
import unicodedata
from pathlib import Path

_NON_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")
_LEADING_NUMBERING_RE = re.compile(r"^\s*\d+\s*[\.\-_]\s*")


def normalize_plate(text):
    return _NON_ALNUM_RE.sub("", text or "").upper()


def normalize_name(text):
    text = unicodedata.normalize("NFC", text or "")
    text = _LEADING_NUMBERING_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.upper()


def _name_words(text):
    """Tách tên thành list từ, bỏ dấu chấm/phẩy cuối mỗi từ - dấu hiệu viết
    tắt (vd "Tr." "Đ," -> "Tr" "Đ")."""
    return [w.rstrip(".,") for w in normalize_name(text).split(" ") if w.rstrip(".,")]


def is_abbreviated_match(short_name, full_name):
    """True nếu `short_name` có thể là bản viết tắt của `full_name`: cùng
    số từ, và mỗi từ của short_name bằng hệt hoặc là phần đầu (viết tắt)
    của từ tương ứng trong full_name - áp dụng cho cả họ lẫn tên đệm (vd
    "Trần N. Đ. Phương" khớp "Trần Ngô Đông Phương", "H. Bá Đ, Khoa" khớp
    "Huỳnh Bá Đăng Khoa")."""
    short_words = _name_words(short_name)
    full_words = _name_words(full_name)
    if not short_words or len(short_words) != len(full_words):
        return False
    return all(full_w.startswith(short_w) for short_w, full_w in zip(short_words, full_words))


def _is_prefixed_exact_match(name_with_prefix, full_name):
    """True nếu `name_with_prefix` = 1 vài từ tiền tố (vd "Thầy", "Cô") +
    đúng hệt `full_name` - xử lý trường hợp Excel lỡ nhập kèm tiền tố xưng
    danh trước tên (vd "Thầy Lê Ngọc Hòa" khớp "Lê Ngọc Hòa")."""
    name_words = _name_words(name_with_prefix)
    full_words = _name_words(full_name)
    if not full_words or len(name_words) <= len(full_words):
        return False
    return name_words[-len(full_words):] == full_words


def find_teacher_candidates(name, dsgv):
    """Tìm giáo viên trong `dsgv` (dict ma_gv -> {"ten": ...}) khớp với
    `name` - cùng hệt (bỏ tiền tố xưng danh), hoặc tên viết tắt (xem
    is_abbreviated_match). Trả về list [(ma_gv, ten_day_du), ...] - rỗng
    nếu không ai khớp, nhiều hơn 1 phần tử nếu khớp mơ hồ với nhiều giáo
    viên (không tự quyết định được, cần người dùng tự sửa lại cho rõ)."""
    return [
        (ma_gv, rec["ten"])
        for ma_gv, rec in dsgv.items()
        if is_abbreviated_match(name, rec["ten"]) or _is_prefixed_exact_match(name, rec["ten"])
    ]


def find_lich_hoc_pdf(lop, candidate_files):
    """Tìm file lịch học PDF theo mã lớp (vd "BK0726") trong 1 danh sách
    file PDF đã chọn sẵn (không tự quét thư mục)."""
    code = normalize_plate(lop)  # chỉ dùng để bỏ dấu cách/ký tự đặc biệt
    if not code:
        return None
    for f in candidate_files:
        f = Path(f)
        if code in normalize_plate(f.stem):
            return f
    return None


def find_xe_doc(xe_bien_so, folder):
    """Tìm giấy tờ xe (.docx/.pdf) theo biển số, trong 1 thư mục cho trước."""
    plate = normalize_plate(xe_bien_so)
    folder = Path(folder) if folder else None
    if not plate or not folder or not folder.is_dir():
        return None
    candidates = []
    for f in sorted(folder.iterdir()):
        if f.suffix.lower() not in (".docx", ".pdf"):
            continue
        if plate in normalize_plate(f.stem):
            candidates.append(f)
    return candidates[0] if candidates else None


def find_ho_so_gv(ten_giao_vien, folder):
    """Tìm hồ sơ cá nhân (PDF) theo tên giáo viên, trong 1 thư mục cho trước."""
    target = normalize_name(ten_giao_vien)
    folder = Path(folder) if folder else None
    if not target or not folder or not folder.is_dir():
        return None
    for f in sorted(folder.iterdir()):
        if f.suffix.lower() != ".pdf":
            continue
        if normalize_name(f.stem) == target:
            return f
    # fallback: tên nằm trong tên file (phòng khi file có hậu tố thêm)
    for f in sorted(folder.iterdir()):
        if f.suffix.lower() != ".pdf":
            continue
        if target in normalize_name(f.stem):
            return f
    return None

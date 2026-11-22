"""Vài hàm chuẩn hoá chuỗi dùng chung (biển số/tên) - bản RÚT GỌN của
quanlygiaovien/attach_files.py, chỉ giữ đúng phần tab CẤP XĂNG/ĐỊNH MỨC
XĂNG/TỔNG HỢP CẤP XĂNG cần (không có phần tìm file PDF/Word theo tên - app
này không đọc Excel nguồn, tên giáo viên lấy thẳng từ docs/cap_xang_data.json)."""
import re
import unicodedata

_NON_ALNUM_RE = re.compile(r"[^A-Za-z0-9]")
_LEADING_NUMBERING_RE = re.compile(r"^\s*\d+\s*[\.\-_]\s*")


def normalize_plate(text):
    return _NON_ALNUM_RE.sub("", text or "").upper()


def normalize_name(text):
    text = unicodedata.normalize("NFC", text or "")
    text = _LEADING_NUMBERING_RE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text.upper()


def given_name_sort_key(full_name):
    """Khóa sắp xếp theo TÊN (từ cuối cùng), không theo HỌ - vd "Nguyễn Văn
    An" sắp xếp theo "An". Dùng cho mọi nơi hiển thị danh sách/combobox tên
    giáo viên trong app."""
    parts = (full_name or "").split()
    return (parts[-1], full_name) if parts else (full_name, full_name)

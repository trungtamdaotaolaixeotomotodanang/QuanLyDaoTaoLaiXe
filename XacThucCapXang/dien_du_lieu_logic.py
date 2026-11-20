"""Logic cho tab "Điền dữ liệu" (xem dien_du_lieu_panel.py) - tự động điền 2
cột Giáo viên quản lý/Khóa cho sổ cái hóa đơn GỐC (vd BKeHDMST.xlsx - chưa ai
gõ tay 3 cột Tự nhập) bằng cách suy luận TRỰC TIẾP từ dữ liệu khóa/giáo viên/
xe trong docs/cap_xang_data.json: khớp BIỂN SỐ của dòng với biển số 1 nhóm
(giáo viên, xe) trong 1 khóa, VÀ ngày hóa đơn phải nằm trong khoảng tu_ngay-
den_ngay của khóa đó.

KHÁC với quanlycapxang/merge_tu_nhap_logic.py (suy luận bằng cách khớp "Mã
tra cứu" với 1 sổ KHÁC đã có người gõ tay sẵn - cần có sổ đó) - ở đây suy
luận THẲNG từ phân công, không cần sổ nào đã gõ sẵn trước.

Thực tế nhiều khóa dùng CHUNG 1 xe với khoảng ngày CHỒNG LẤN nhau (khóa sau
khai giảng trước khi khóa trước bế giảng, học viên dùng xe xen kẽ) - nên 1
(biển số, ngày) có thể khớp NHIỀU (giáo viên, khóa). Coi mỗi cột (giáo viên/
khóa) CHẮC CHẮN RIÊNG: nếu mọi candidate cùng 1 giáo viên thì điền Giáo viên
dù Khóa còn mơ hồ (và ngược lại) - tận dụng được nhiều hơn so với chỉ điền
khi khớp đúng 1 CẶP. Xem autofill_for_row()."""
from datetime import datetime

import attach_files

_DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y")

# Số HĐ gốc/Cột 24 ở cuối BKeHDMST.xlsx đều rỗng, không dùng tới.
DROP_LAST_N_COLUMNS = 2
TEACHER_COL_NAME = "Giáo viên quản lý"
KHOA_COL_NAME = "Khóa"
SO_PHIEU_COL_NAME = "Số phiếu"
NEW_COLUMNS = (TEACHER_COL_NAME, KHOA_COL_NAME, SO_PHIEU_COL_NAME)

# "ngày" generic đứng SAU để không đoán nhầm "Ngày-giờ CTỪ" (có trong
# BKeHDMST.xlsx, đứng TRƯỚC "Ngày HĐ" trong bảng) thành cột ngày hóa đơn.
_NGAY_KEYWORD_GROUPS = [["ngày hđ", "ngày cấp"], ["ngày"]]
_PLATE_KEYWORD_GROUPS = [["bsx", "biển số", "biển kiểm soát", "số xe"]]


def _parse_course_date(text):
    text = (text or "").strip()
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_cell_date(value):
    if isinstance(value, datetime):
        return value
    return _parse_course_date(str(value)) if value not in (None, "") else None


def _find_first(headers, keyword_groups):
    """Dò theo TỪNG NHÓM từ khóa theo thứ tự ưu tiên (nhóm trước cụ thể hơn
    nhóm sau) - quét hết các cột với nhóm cụ thể trước khi rơi xuống nhóm
    chung chung hơn, tránh đoán nhầm (xem _NGAY_KEYWORD_GROUPS)."""
    for keywords in keyword_groups:
        for i, h in enumerate(headers):
            low = (h or "").strip().lower()
            if any(kw in low for kw in keywords):
                return i
    return None


def detect_columns(headers):
    """-> {"ngay": chỉ mục cột ngày hóa đơn, "bien_so": chỉ mục cột biển số}
    (None nếu không đoán được - người dùng tự chọn lại qua combo)."""
    return {
        "ngay": _find_first(headers, _NGAY_KEYWORD_GROUPS),
        "bien_so": _find_first(headers, _PLATE_KEYWORD_GROUPS),
    }


def build_output_headers(source_headers):
    """Bỏ DROP_LAST_N_COLUMNS cột cuối của sổ gốc (rỗng, không dùng) + thêm 3
    cột mới NEW_COLUMNS vào cuối."""
    kept = (
        source_headers[:-DROP_LAST_N_COLUMNS] if len(source_headers) > DROP_LAST_N_COLUMNS
        else list(source_headers)
    )
    return list(kept) + list(NEW_COLUMNS)


def candidates_for_row(model, ngay_value, xe_value):
    """-> list các cặp khớp {"giao_vien", "khoa_lop", "ma_khoa", "tu_ngay",
    "den_ngay"} - MỌI (khóa, giáo viên) có xe khớp `xe_value` VÀ `ngay_value`
    nằm trong khoảng tu_ngay-den_ngay của khóa đó. [] nếu biển số/ngày thiếu/
    không đọc được, hoặc không khớp khóa nào."""
    ngay = parse_cell_date(ngay_value)
    xe_norm = attach_files.normalize_plate(str(xe_value)) if xe_value not in (None, "") else ""
    if not ngay or not xe_norm:
        return []
    out, seen = [], set()
    for c in model.courses:
        tu_ngay = _parse_course_date(c.get("tu_ngay", ""))
        den_ngay = _parse_course_date(c.get("den_ngay", ""))
        if not tu_ngay or not den_ngay or not (tu_ngay <= ngay <= den_ngay):
            continue
        for g in c.get("groups", []):
            if attach_files.normalize_plate(g.get("xe", "")) != xe_norm:
                continue
            key = (attach_files.normalize_name(g.get("giao_vien", "")), c.get("lop", ""))
            if key in seen:
                continue
            seen.add(key)
            out.append({
                "giao_vien": g.get("giao_vien", ""), "khoa_lop": c.get("lop", ""),
                "ma_khoa": c.get("ma_khoa", ""),
                "tu_ngay": c.get("tu_ngay", ""), "den_ngay": c.get("den_ngay", ""),
            })
    return out


def autofill_for_row(candidates):
    """-> (giao_vien_or_None, khoa_lop_or_None) - điền CHẮC CHẮN RIÊNG từng
    cột: cột nào mọi candidate đều cùng 1 giá trị thì điền cột đó, cột kia
    vẫn có thể None nếu còn mơ hồ giữa các candidate (xem docstring module)."""
    if not candidates:
        return None, None
    gvs = {c["giao_vien"] for c in candidates}
    khoas = {c["khoa_lop"] for c in candidates}
    giao_vien = next(iter(gvs)) if len(gvs) == 1 else None
    khoa_lop = next(iter(khoas)) if len(khoas) == 1 else None
    return giao_vien, khoa_lop

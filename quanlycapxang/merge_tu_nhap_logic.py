"""Logic ghép 3 cột "Tự nhập" (GV Quản lý / Khóa / Số phiếu) từ 1 sổ NGUỒN đã
có người gõ tay (vd "Re_ Cấp nhiên liệu cho xe tập lái...") sang 1 sổ ĐÍCH
chưa có 3 cột đó (vd BKeHDMST.xlsx - sổ cái hóa đơn gốc xuất ra từ phần mềm
Petrolimex, chỉ có cột "Mã tra cứu" chứ chưa ai gõ GV/Khóa/Số phiếu) - khớp
theo "Mã tra cứu" (mã duy nhất Petrolimex cấp cho từng hóa đơn, có ở cả 2 sổ).

CHỈ điền vào dòng "chắc chắn đúng" - xem compute_fill(): Mã tra cứu phải khớp
DUY NHẤT 1 dòng nguồn (không mơ hồ/trùng dữ liệu khác nhau), VÀ Số HĐ + Tổng
tiền (nếu dò được cột này ở cả 2 sổ) của 2 dòng phải khớp nhau - phòng khi Mã
tra cứu trùng ngẫu nhiên hoặc sổ nguồn có lỗi gõ. Khớp key nhưng 1 trong 3 ô
GV/Khóa/Số phiếu đang rỗng ở sổ nguồn thì cũng để trống ở sổ đích luôn, KHÔNG
suy diễn. Mọi trường hợp còn lại (không khớp, mơ hồ, lệch Số HĐ/Tổng tiền) ->
để trống, người dùng tự nhập sau - đúng yêu cầu "chưa chắc đúng thì để
trống"."""
import re

KEY_KEYWORDS = ["mã tra cứu", "ma tra cuu"]
SO_HD_KEYWORDS = ["số hđ", "so hd"]
TONG_TIEN_KEYWORDS = ["tổng tiền", "tong tien"]

NEW_COLUMNS = ("giao_vien", "khoa", "so_phieu")
NEW_HEADERS = {
    "giao_vien": "GV Quản lý*\n(Tự nhập)",
    "khoa": "Khóa*\n(Tự nhập)",
    "so_phieu": "Số phiếu*\n(Tự nhập)",
}


def _norm_header(h):
    return re.sub(r"\s+", " ", (h or "")).strip().lower()


def _find_first(headers, keywords):
    for i, h in enumerate(headers):
        if any(kw in _norm_header(h) for kw in keywords):
            return i
    return None


def detect_key_column(headers):
    return _find_first(headers, KEY_KEYWORDS)


def detect_check_columns(headers):
    """Các cột dùng để kiểm tra chéo cho chắc (không bắt buộc dò được) ->
    dict {"so_hd": idx|None, "tong_tien": idx|None}."""
    return {
        "so_hd": _find_first(headers, SO_HD_KEYWORDS),
        "tong_tien": _find_first(headers, TONG_TIEN_KEYWORDS),
    }


def _norm_key(value):
    if value in (None, ""):
        return None
    return str(value).strip()


def _norm_number(value):
    if value in (None, ""):
        return None
    try:
        return round(float(str(value).replace(",", "").strip()), 2)
    except ValueError:
        return None


def _values_equal(a, b):
    if a in (None, "") and b in (None, ""):
        return True
    na, nb = _norm_number(a), _norm_number(b)
    if na is not None or nb is not None:
        return na == nb
    return str(a).strip() == str(b).strip()


def build_source_map(source_rows, key_col, new_cols, check_cols):
    """new_cols/check_cols: dict tên -> chỉ mục cột ở sổ NGUỒN (None nếu
    không có/không dò được). -> dict Mã tra cứu (đã chuẩn hóa) -> entry (xem
    bên dưới), hoặc -> None nếu Mã tra cứu đó xuất hiện NHIỀU lần ở sổ nguồn
    với dữ liệu khác nhau (mơ hồ, không đủ chắc để tự điền)."""
    result = {}
    compare_keys = [*new_cols.keys(), *(f"check_{k}" for k in check_cols)]
    for row in source_rows:
        key = _norm_key(row[key_col]) if key_col is not None else None
        if key is None:
            continue
        entry = {name: (row[idx] if idx is not None else None) for name, idx in new_cols.items()}
        entry.update({f"check_{name}": (row[idx] if idx is not None else None) for name, idx in check_cols.items()})
        if key not in result:
            result[key] = entry
        elif result[key] is not None and not all(
            _values_equal(result[key].get(k), entry.get(k)) for k in compare_keys
        ):
            result[key] = None  # nhiều dòng nguồn khác nhau cùng 1 Mã tra cứu -> mơ hồ
    return result


def compute_fill(target_rows, target_row_numbers, key_col, check_cols, source_map):
    """check_cols: chỉ mục cột Số HĐ/Tổng tiền ở sổ ĐÍCH (cùng tên khóa với
    check_cols truyền cho build_source_map, để so đúng cặp). -> list dict
    {"row": số dòng thật, "values": {"giao_vien", "khoa", "so_phieu"} (None =
    để trống), "reason": mô tả ngắn lý do điền/để trống - hiện ở bảng xem
    trước cho người dùng tự kiểm tra lại}."""
    results = []
    for row, rn in zip(target_rows, target_row_numbers):
        key = _norm_key(row[key_col]) if key_col is not None else None
        values = {c: None for c in NEW_COLUMNS}
        if key is None:
            reason = "Không có Mã tra cứu"
        elif key not in source_map:
            reason = "Không tìm thấy Mã tra cứu này trong sổ nguồn"
        elif source_map[key] is None:
            reason = "Mã tra cứu này khớp nhiều dòng khác nhau ở sổ nguồn - bỏ qua cho an toàn"
        else:
            entry = source_map[key]
            mismatch_col = None
            for name, idx in check_cols.items():
                if idx is None:
                    continue
                if not _values_equal(row[idx], entry.get(f"check_{name}")):
                    mismatch_col = name
                    break
            if mismatch_col:
                reason = f"Khớp Mã tra cứu nhưng lệch {mismatch_col.replace('_', ' ')} - bỏ qua cho an toàn"
            else:
                for c in NEW_COLUMNS:
                    v = entry.get(c)
                    values[c] = v if v not in (None, "") else None
                reason = "Đã điền từ sổ nguồn" if any(v is not None for v in values.values()) else (
                    "Khớp Mã tra cứu nhưng sổ nguồn cũng để trống 3 cột này"
                )
        results.append({"row": rn, "values": values, "reason": reason})
    return results

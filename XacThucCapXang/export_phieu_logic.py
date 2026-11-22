"""Từ sổ cấp xăng đã gõ tay (đã xác thực qua validate1_logic.validate_rows,
xem export_phieu_panel.py) -> JSON đúng schema quanlycapxang/cap_xang_store.py
("phiếu cấp xăng": giao_vien, ma_khoa, lop, xe, ngay_cap, so_phieu, mat_hang,
so_luong, so_luong_text, so_phieu_count, ghi_chu, phan_loai_so) để quanlycapxang
NHẬP vào (nút "⬇ Nhập phiếu từ JSON...").

Chỉ xuất những dòng KHÔNG có lỗi mức "loi" ở validate1_logic (cảnh báo vẫn
xuất, vì cảnh báo không có nghĩa là sai - vd ngày cấp lệch vài hôm so với
thời gian khóa). Tên giáo viên/biển số lấy từ phân công KHỚP trong
docs/cap_xang_data.json (không dùng nguyên văn gõ tay) để nhất quán với dữ
liệu đang có, tránh lệch khoảng trắng/viết hoa."""
import json
from datetime import datetime

import attach_files
import validate1_logic

MAT_HANG_KEYWORDS = ["mặt hàng", "hàng hóa"]
SO_LUONG_KEYWORDS = ["số lượng"]
GHI_CHU_KEYWORDS = ["ghi chú"]
PHAN_LOAI_KEYWORDS = ["phân loại"]


def detect_extra_columns(headers):
    """-> dict {"mat_hang", "so_luong", "ghi_chu", "phan_loai_so"} -> chỉ mục
    cột (None nếu không đoán được) - cùng cách dò theo từ khóa với
    validate1_logic.detect_columns, cho các cột KHÔNG cần đối chiếu (chỉ cần
    chép nguyên giá trị vào JSON)."""
    idx = range(len(headers))
    return {
        "mat_hang": validate1_logic.find_first(headers, idx, MAT_HANG_KEYWORDS),
        "so_luong": validate1_logic.find_first(headers, idx, SO_LUONG_KEYWORDS),
        "ghi_chu": validate1_logic.find_first(headers, idx, GHI_CHU_KEYWORDS),
        "phan_loai_so": validate1_logic.find_first(headers, idx, PHAN_LOAI_KEYWORDS),
    }


def _row_has_loi(issues_by_row, row_number):
    return any(it["severity"] == "loi" for it in issues_by_row.get(row_number, []))


def _matched_group(model, khoa_courses, gv_raw, xe_raw):
    """Tìm group (trong 1 trong các khoa_courses) khớp (giáo viên, biển số) -
    trả về (course, group) CHUẨN theo docs/cap_xang_data.json, hoặc None nếu
    không khớp (không nên xảy ra với dòng đã qua validate_rows không có lỗi,
    nhưng vẫn kiểm tra lại cho chắc, không giả định)."""
    name_norm = attach_files.normalize_name(gv_raw)
    plate_norm = attach_files.normalize_plate(xe_raw)
    for c in khoa_courses:
        for g in c.get("groups", []):
            if (
                attach_files.normalize_name(g.get("giao_vien", "")) == name_norm
                and attach_files.normalize_plate(g.get("xe", "")) == plate_norm
            ):
                return c, g
    return None, None


def build_records(headers, rows, row_numbers, model, columns, extra_columns, issues):
    """columns: dict từ validate1_logic.detect_columns() (giao_vien/khoa/
    so_phieu/bien_so/ngay). extra_columns: dict từ detect_extra_columns()
    (mat_hang/so_luong/ghi_chu/phan_loai_so). issues: list từ
    validate1_logic.validate_rows() (cho CHÍNH các rows/row_numbers này).

    -> (records, skipped) - skipped: list {"row", "reason"} cho các dòng
    KHÔNG xuất được, kèm lý do cụ thể để người dùng biết, không âm thầm mất
    dữ liệu."""
    issues_by_row = {}
    for it in issues:
        issues_by_row.setdefault(it["row"], []).append(it)

    gv_col = columns.get("giao_vien")
    khoa_col = columns.get("khoa")
    sp_col = columns.get("so_phieu")
    xe_col = columns.get("bien_so")
    ngay_col = columns.get("ngay")
    mat_hang_col = extra_columns.get("mat_hang")
    so_luong_col = extra_columns.get("so_luong")
    ghi_chu_col = extra_columns.get("ghi_chu")
    phan_loai_col = extra_columns.get("phan_loai_so")

    records = []
    skipped = []
    for i, row in enumerate(rows):
        rn = row_numbers[i]
        if _row_has_loi(issues_by_row, rn):
            skipped.append({"row": rn, "reason": "Còn lỗi chưa sửa (xem tab Xác thực)."})
            continue

        gv_raw = str(row[gv_col]).strip() if gv_col is not None and row[gv_col] not in (None, "") else ""
        khoa_raw = str(row[khoa_col]).strip() if khoa_col is not None and row[khoa_col] not in (None, "") else ""
        xe_raw = str(row[xe_col]).strip() if xe_col is not None and row[xe_col] not in (None, "") else ""
        if not gv_raw or not khoa_raw or not xe_raw:
            skipped.append({"row": rn, "reason": "Thiếu Giáo viên/Khóa/Biển số - không đối chiếu được."})
            continue

        khoa_courses = validate1_logic.course_for_khoa_value(model, khoa_raw)
        if not khoa_courses:
            skipped.append({"row": rn, "reason": f"Khóa/lớp '{khoa_raw}' không khớp dữ liệu hiện có."})
            continue

        course, group = _matched_group(model, khoa_courses, gv_raw, xe_raw)
        if course is None:
            skipped.append({
                "row": rn,
                "reason": f"Bộ (GV, Khóa, Biển số) '{gv_raw}' + '{khoa_raw}' + '{xe_raw}' không khớp phân công nào.",
            })
            continue

        ngay_raw = row[ngay_col] if ngay_col is not None else None
        ngay_cap = validate1_logic.parse_cell_date(ngay_raw)
        if ngay_cap is None:
            skipped.append({"row": rn, "reason": "Không đọc được ngày cấp."})
            continue

        so_phieu_raw = row[sp_col] if sp_col is not None else None
        so_luong_raw = row[so_luong_col] if so_luong_col is not None else None
        try:
            so_luong = float(so_luong_raw)
        except (TypeError, ValueError):
            skipped.append({"row": rn, "reason": f"Cột Số lượng không phải số: '{so_luong_raw}'."})
            continue

        records.append({
            "giao_vien": group["giao_vien"],
            "ma_khoa": course["ma_khoa"],
            "lop": course.get("lop", ""),
            "xe": group["xe"],
            "ngay_cap": ngay_cap.strftime("%d/%m/%Y"),
            "so_phieu": str(so_phieu_raw).strip() if so_phieu_raw not in (None, "") else "",
            "mat_hang": str(row[mat_hang_col]).strip() if mat_hang_col is not None and row[mat_hang_col] not in (None, "") else "Xăng",
            "so_luong": so_luong,
            "so_luong_text": str(so_luong_raw).strip() if so_luong_raw not in (None, "") else str(so_luong),
            "so_phieu_count": 1,
            "ghi_chu": str(row[ghi_chu_col]).strip() if ghi_chu_col is not None and row[ghi_chu_col] not in (None, "") else "",
            "phan_loai_so": str(row[phan_loai_col]).strip() if phan_loai_col is not None and row[phan_loai_col] not in (None, "") else "",
        })

    return records, skipped


def write_json(records, path, source_file, source_sheet):
    data = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_file": source_file,
        "source_sheet": source_sheet,
        "records": records,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

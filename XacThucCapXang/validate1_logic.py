"""Logic đối chiếu dùng chung cho validate1.py (sheet PETROLIMEX) và
validate2.py (sheet PETROLIMEX + NLDACAP) - đối chiếu các cột do người dùng
tự gõ tay (giáo viên quản lý/khóa-lớp/biển số/ngày cấp) với dữ liệu khóa/
giáo viên/xe trong docs/cap_xang_data.json (qua CapXangModel) - KHÁC với tab
"XÁC THỰC EXCEL" (excel_validate.py) ở chỗ:

- excel_validate.cross_reference_issues() đối chiếu TỪNG cột riêng lẻ (giáo
  viên lạ? khóa lạ? biển số lạ?) - không biết 3 ô đó có KHỚP VỚI NHAU không
  (vd giáo viên đúng, khóa đúng, nhưng 2 cái đó không phải cùng 1 phân công).
- Ở đây đối chiếu CẢ BỘ (khóa+giáo viên+biển số) có đúng là 1 phân công đang
  có trong dữ liệu không - giống cách CapXangModel.orphan_reason() làm cho
  phiếu cấp xăng nhập tay trong app, chỉ khác là cột "Khóa"/"Khoá học" trong
  sổ gốc ghi MÃ LỚP (vd "BK0626"), không phải mã khóa đầy đủ (vd
  "48002K262013") - nên tra theo courses_for_lop() trước khi so khớp.

2 sheet PETROLIMEX/NLDACAP không đặt tên cột giống nhau, và không phải cột
nào cần đối chiếu cũng ghi "(Tự nhập)" trong tiêu đề (vd "Khoá học" ở
NLDACAP) - nên detect_columns() dò theo TỪ KHÓA trên TOÀN BỘ tiêu đề trước,
chỉ dùng "cột có chữ Tự nhập"/"3 cột cuối" làm phương án dự phòng cuối cùng
khi không dò được theo từ khóa."""
import re
from datetime import datetime

import attach_files

MANUAL_KEYWORD = "tự nhập"
TEACHER_KEYWORDS = ["giáo viên", "gv"]
KHOA_KEYWORDS = ["khóa", "khoá"]
SO_PHIEU_KEYWORDS = ["số phiếu", "so phieu"]
PLATE_KEYWORDS = ["bsx", "biển số", "biển kiểm soát", "số xe"]
NGAY_KEYWORDS = ["ngày cấp", "ngày hđ", "ngày"]

_SO_PHIEU_RE = re.compile(r"^\s*\d+\s*(?:[,;]\s*\d+\s*)*$")
_COURSE_DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y")


def _norm_header(h):
    return re.sub(r"\s+", " ", (h or "")).strip().lower()


def _find_first(headers, indices, keywords):
    for i in indices:
        low = _norm_header(headers[i])
        if any(kw in low for kw in keywords):
            return i
    return None


def detect_columns(headers):
    """-> dict {"giao_vien", "khoa", "so_phieu", "bien_so", "ngay"} -> chỉ
    mục cột (None nếu không đoán được). Dò theo từ khóa trên TOÀN BỘ tiêu đề
    trước (đúng cho cả 2 sheet dù tên/thứ tự cột khác nhau); với khóa còn lại
    chưa dò được, mới thử trong các cột có chữ "Tự nhập" (hoặc 3 cột cuối nếu
    sheet không có cột nào ghi vậy) - theo đúng cấu trúc sổ Petrolimex gốc."""
    all_idx = range(len(headers))
    result = {
        "giao_vien": _find_first(headers, all_idx, TEACHER_KEYWORDS),
        "khoa": _find_first(headers, all_idx, KHOA_KEYWORDS),
        "so_phieu": _find_first(headers, all_idx, SO_PHIEU_KEYWORDS),
        "bien_so": _find_first(headers, all_idx, PLATE_KEYWORDS),
        "ngay": _find_first(headers, all_idx, NGAY_KEYWORDS),
    }
    if all(v is not None for v in result.values()):
        return result

    manual = [i for i, h in enumerate(headers) if MANUAL_KEYWORD in _norm_header(h)]
    if len(manual) < 3 and len(headers) >= 3:
        manual = list(range(len(headers) - 3, len(headers)))
    remaining_slots = [i for i in manual if i not in result.values()]
    remaining_keys = [k for k in ("giao_vien", "khoa", "so_phieu") if result[k] is None]
    for key, idx in zip(remaining_keys, remaining_slots):
        result[key] = idx
    return result


def _parse_course_date(text):
    text = (text or "").strip()
    for fmt in _COURSE_DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _parse_cell_date(value):
    """value có thể là datetime (Excel tự định dạng ngày) hoặc chuỗi gõ tay
    dd/m/yyyy - khoan xét lỗi định dạng ở đây (đã có excel_validate.
    basic_issues lo phần đó), chỉ cần lấy ra ngày để so khoảng với khóa."""
    if isinstance(value, datetime):
        return value
    return _parse_course_date(str(value)) if value not in (None, "") else None


def _course_for_khoa_value(model, khoa_value):
    """value trong cột "Khóa" của sổ gốc có thể là mã lớp (thường gặp, vd
    "BK0626") hoặc mã khóa đầy đủ (vd "48002K262013") - thử cả 2 cách tra,
    ưu tiên mã khóa đầy đủ trước vì nó xác định đúng 1 khóa duy nhất. -> list
    khóa (course dict) khớp, hoặc [] nếu không khớp khóa nào."""
    if khoa_value in model._courses_by_code:
        return [model._courses_by_code[khoa_value]]
    return model.courses_for_lop(khoa_value)


def validate_rows(headers, rows, row_numbers, model, columns):
    """columns: dict từ detect_columns() (hoặc do người dùng tự chọn lại) -
    thiếu key nào/key đó None thì bỏ qua đúng phần đối chiếu liên quan tới cột
    đó, không báo lỗi. -> list issue {"row", "column", "severity":
    "loi"|"canh_bao", "problem", "fix"} - chỉ xét các cột gõ tay (giáo viên/
    khóa/số phiếu/ngày cấp) + đối chiếu chéo với cột biển số, KHÔNG rà các lỗi
    định dạng chung (đã có ở excel_validate.basic_issues, không lặp lại ở
    đây)."""
    issues = []
    gv_col = columns.get("giao_vien")
    khoa_col = columns.get("khoa")
    sp_col = columns.get("so_phieu")
    xe_col = columns.get("bien_so")
    ngay_col = columns.get("ngay")

    for i, row in enumerate(rows):
        rn = row_numbers[i]
        gv_raw = row[gv_col] if gv_col is not None else None
        khoa_raw = row[khoa_col] if khoa_col is not None else None
        sp_raw = row[sp_col] if sp_col is not None else None
        xe_raw = row[xe_col] if xe_col is not None else None
        ngay_raw = row[ngay_col] if ngay_col is not None else None

        gv = str(gv_raw).strip() if gv_raw not in (None, "") else ""
        khoa = str(khoa_raw).strip() if khoa_raw not in (None, "") else ""
        xe = str(xe_raw).strip() if xe_raw not in (None, "") else ""

        gv_known = False
        if gv_col is not None:
            if not gv:
                issues.append({
                    "row": rn, "column": headers[gv_col], "severity": "loi",
                    "problem": "Ô \"GV Quản lý\" bị bỏ trống.",
                    "fix": "Nhập tên giáo viên quản lý xe cho dòng này.",
                })
            elif not model.is_known_teacher(gv):
                goi_y = model.suggest_teachers(gv)
                fix = "Không có trong danh sách giáo viên hiện tại (docs/cap_xang_data.json)"
                fix += f" - gần giống: {', '.join(goi_y)}." if goi_y else "."
                issues.append({
                    "row": rn, "column": headers[gv_col], "severity": "loi",
                    "problem": f"Giáo viên lạ: '{gv}'", "fix": fix,
                })
            else:
                gv_known = True

        khoa_courses = []
        if khoa_col is not None:
            if not khoa:
                issues.append({
                    "row": rn, "column": headers[khoa_col], "severity": "loi",
                    "problem": "Ô \"Khóa\" bị bỏ trống.",
                    "fix": "Nhập mã lớp/mã khóa cho dòng này.",
                })
            else:
                khoa_courses = _course_for_khoa_value(model, khoa)
                if not khoa_courses:
                    issues.append({
                        "row": rn, "column": headers[khoa_col], "severity": "loi",
                        "problem": f"Khóa/lớp lạ: '{khoa}'",
                        "fix": "Không khớp mã khóa lẫn mã lớp nào trong dữ liệu khóa hiện tại.",
                    })

        if sp_col is not None:
            if sp_raw in (None, ""):
                issues.append({
                    "row": rn, "column": headers[sp_col], "severity": "canh_bao",
                    "problem": "Ô \"Số phiếu\" bị bỏ trống.",
                    "fix": "Nhập số phiếu cấp nhiên liệu tương ứng (1 hoặc nhiều số, cách nhau dấu phẩy).",
                })
            elif isinstance(sp_raw, datetime):
                issues.append({
                    "row": rn, "column": headers[sp_col], "severity": "canh_bao",
                    "problem": f"Ô bị Excel tự đổi thành ngày tháng: '{sp_raw.strftime('%d/%m/%Y')}'",
                    "fix": (
                        "Có thể gõ số phiếu dạng N/Q (vd \"1/9\") và Excel hiểu nhầm thành ngày - "
                        "định dạng lại ô này thành Văn bản (Text) trước khi nhập."
                    ),
                })
            elif not _SO_PHIEU_RE.match(str(sp_raw)):
                # Để CẢNH BÁO (không phải lỗi): 1 số sổ dùng ký hiệu "N/Q<quyển>"
                # (vd "1/Q9" = số phiếu 1, quyển 9) thay vì chỉ số thuần, không
                # hẳn là gõ sai - không đủ chắc để báo lỗi cứng.
                issues.append({
                    "row": rn, "column": headers[sp_col], "severity": "canh_bao",
                    "problem": f"Số phiếu có ký tự lạ, không chỉ gồm số/dấu phẩy: '{sp_raw}'",
                    "fix": "Kiểm tra lại - nếu không phải ký hiệu quyển (vd \"1/Q9\") thì có thể đã gõ nhầm.",
                })

        # Đối chiếu chéo: GV + Khóa + biển số có cùng khớp 1 phân công không -
        # chỉ xét khi cả giáo viên lẫn khóa đều đã xác định đúng ở trên (nếu
        # 1 trong 2 đã sai thì báo lỗi ở đó là đủ, không cần báo thêm lỗi lệch
        # bộ 3 gây rối).
        if gv_known and khoa_courses and xe_col is not None:
            name_norm = attach_files.normalize_name(gv)
            plate_norm = attach_files.normalize_plate(xe) if xe else ""
            all_groups = [g for c in khoa_courses for g in c.get("groups", [])]
            matched = any(
                attach_files.normalize_name(g.get("giao_vien", "")) == name_norm
                and attach_files.normalize_plate(g.get("xe", "")) == plate_norm
                for g in all_groups
            )
            if not matched:
                gv_cua_xe = sorted({
                    g.get("giao_vien", "") for g in all_groups
                    if attach_files.normalize_plate(g.get("xe", "")) == plate_norm
                })
                xe_cua_gv = sorted({
                    g.get("xe", "") for g in all_groups
                    if attach_files.normalize_name(g.get("giao_vien", "")) == name_norm
                })
                if gv_cua_xe:
                    fix = (
                        f"Nếu đúng là xe {xe} thì GV Quản lý phải là '{gv_cua_xe[0]}'"
                        + (f" (hoặc {', '.join(gv_cua_xe[1:])})" if len(gv_cua_xe) > 1 else "")
                        + f"; nếu đúng là '{gv}' thì biển số phải là "
                        + (', '.join(xe_cua_gv) if xe_cua_gv else "1 xe khác trong khóa này")
                        + "."
                    )
                elif xe_cua_gv:
                    fix = f"'{gv}' trong khóa '{khoa}' được phân công xe: {', '.join(xe_cua_gv)}, không phải {xe}."
                else:
                    teachers_txt = ", ".join(sorted({g.get("giao_vien", "") for g in all_groups})) or "(không có)"
                    fix = f"Khóa '{khoa}' chỉ phân công cho: {teachers_txt}."
                issues.append({
                    "row": rn, "column": f"{headers[gv_col]} / {headers[khoa_col]}", "severity": "loi",
                    "problem": f"Bộ (GV, Khóa, Biển số) không khớp phân công nào: '{gv}' + '{khoa}' + '{xe}'",
                    "fix": fix,
                })

        # Ngày cấp có nằm trong khoảng tu_ngay-den_ngay của khóa không - chỉ
        # xét khi khóa đã xác định đúng (1 course), còn khóa khớp NHIỀU course
        # (hiện tại courses_for_lop() luôn ra đúng 1, nhưng phòng khi dữ liệu
        # sau này có lớp học lại) thì bỏ qua, tránh báo sai vì không biết course
        # nào. Chỉ CẢNH BÁO (không phải lỗi) - ngày cấp lố vài hôm trước/sau
        # ngày khai/bế giảng vẫn có thể hợp lệ trong thực tế.
        if ngay_col is not None and ngay_raw not in (None, "") and len(khoa_courses) == 1:
            cell_date = _parse_cell_date(ngay_raw)
            tu_ngay = _parse_course_date(khoa_courses[0].get("tu_ngay", ""))
            den_ngay = _parse_course_date(khoa_courses[0].get("den_ngay", ""))
            if cell_date and tu_ngay and den_ngay and not (tu_ngay <= cell_date <= den_ngay):
                issues.append({
                    "row": rn, "column": headers[ngay_col], "severity": "canh_bao",
                    "problem": (
                        f"Ngày cấp '{ngay_raw}' nằm ngoài thời gian khóa '{khoa}' "
                        f"({khoa_courses[0].get('tu_ngay', '')} - {khoa_courses[0].get('den_ngay', '')})"
                    ),
                    "fix": "Kiểm tra lại ngày cấp hoặc mã khóa/lớp - có thể gõ nhầm 1 trong 2.",
                })

    return issues

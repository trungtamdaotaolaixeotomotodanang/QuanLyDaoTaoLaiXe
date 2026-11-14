"""Đọc + rà lỗi 1 sheet Excel BẤT KỲ (không cần biết trước cấu trúc) - dùng
cho tab XÁC THỰC EXCEL:

- list_sheets()/guess_header_row()/read_sheet(): đọc sheet, tự đoán dòng tiêu
  đề (các file cấp xăng thật trong data/ có file để tiêu đề ở dòng 1, có file
  ở dòng 2, 3, 4 - do có tiêu đề lớn/dòng trống phía trên).
- basic_issues(): lỗi/cảnh báo KHÔNG cần biết nội dung cột là gì - trống bất
  thường, trùng dòng, lệch kiểu dữ liệu so với phần còn lại của cột, ngày
  sai/khó đọc, khoảng trắng thừa.
- cross_reference_issues(): đối chiếu cột người dùng chỉ định (giáo viên/
  khóa-lớp/biển số) với dữ liệu khóa/giáo viên/xe đang có trong app
  (CapXangModel) - CẦN người dùng chọn đúng cột vì file thô không có tên
  cột cố định (xem data/*.xlsx: có file ghi "SỐ XE", có file ghi "Biển số*").

Mỗi issue là 1 dict: {"row": số dòng thật trong Excel, "column": tên cột
(None nếu lỗi toàn dòng), "severity": "loi"|"canh_bao", "problem": mô tả,
"fix": gợi ý cách sửa}."""
import re
from datetime import date, datetime

from openpyxl import load_workbook

import attach_files

_DATE_FORMATS = ("%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d.%m.%Y")
# "18/92026" (thiếu 1 dấu /, phần sau ghép liền tháng+năm) - vd gặp thật khi
# nhập liệu tay trong các sổ cấp nhiên liệu.
_MISSING_SLASH_RE = re.compile(r"^(\d{1,2})/(\d{1,2})(\d{4})$")
_DOUBLE_SPACE_RE = re.compile(r"\s{2,}")

TEACHER_KEYWORDS = ["giáo viên", "gv", "họ và tên", "họ tên"]
# CỐ Ý không gồm "khóa"/"khoá" đứng 1 mình - 1 số sổ dùng cột tên y hệt vậy
# cho MỤC KHÁC (vd cột "Khoá" = loại giao dịch "SH"/"ĐT" trong sổ Petrolimex,
# không phải mã khóa học) -> đoán nhầm cột, đối chiếu sai hàng loạt.
COURSE_KEYWORDS = ["khoá học", "khóa học", "mã khóa", "mã khoá", "lớp"]
PLATE_KEYWORDS = ["biển số", "số xe", "biển kiểm soát", "xe"]


def list_sheets(path):
    wb = load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def guess_header_row(path, sheet, max_scan=12):
    """Dòng tiêu đề = dòng có nhiều ô không rỗng nhất trong vài dòng đầu, MÀ
    dòng ngay sau đó cũng có kha khá ô không rỗng (tránh nhận nhầm 1 dòng
    tiêu đề lớn/ghi chú đơn độc như "BẢNG KÊ..." làm dòng tiêu đề cột)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        rows = [row for row in ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)]
    finally:
        wb.close()
    if not rows:
        return 1

    def _non_empty(row):
        return sum(1 for v in row if v not in (None, ""))

    best_row, best_score = 1, -1
    for i, row in enumerate(rows, start=1):
        n = _non_empty(row)
        if n < 2:
            continue
        next_n = _non_empty(rows[i]) if i < len(rows) else 0
        score = n + (2 if next_n >= n * 0.5 else -3)
        if score > best_score:
            best_score, best_row = score, i
    return best_row


def read_sheet(path, sheet, header_row):
    """-> (headers, rows, row_numbers). headers: tên cột (ô trống -> "Cột N"),
    đã khử trùng tên (thêm " (2)" nếu lặp). rows: list các dòng dữ liệu (list
    giá trị thô, giữ nguyên kiểu openpyxl trả về - số/ngày/chuỗi). Bỏ qua
    dòng hoàn toàn trống."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        raw_header = next(
            ws.iter_rows(min_row=header_row, max_row=header_row, values_only=True), ()
        )
        headers, seen = [], {}
        for i, h in enumerate(raw_header):
            name = re.sub(r"\s+", " ", str(h)).strip() if h not in (None, "") else ""
            name = name or f"Cột {i + 1}"
            if name in seen:
                seen[name] += 1
                name = f"{name} ({seen[name]})"
            else:
                seen[name] = 1
            headers.append(name)
        n_col = len(headers)

        rows, row_numbers = [], []
        for i, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True), start=header_row + 1):
            if all(v in (None, "") for v in row):
                continue
            row = list(row[:n_col])
            row += [None] * (n_col - len(row))
            rows.append(row)
            row_numbers.append(i)
        return headers, rows, row_numbers
    finally:
        wb.close()


def guess_column(headers, keywords):
    """Chỉ mục cột đầu tiên có tên chứa 1 trong các từ khóa - None nếu không
    đoán được (để người dùng tự chọn)."""
    for i, h in enumerate(headers):
        low = h.lower()
        if any(kw in low for kw in keywords):
            return i
    return None


def _parse_date_lenient(value):
    """-> (date_or_None, note_or_None). Nhận cả datetime (Excel đã định dạng
    ngày sẵn) lẫn chuỗi gõ tay; thử sửa 2 lỗi gõ tay hay gặp (thiếu dấu /,
    đảo ngày/tháng) trước khi báo không đọc được."""
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    text = str(value).strip()
    if not text:
        return None, None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date(), None
        except ValueError:
            continue
    m = _MISSING_SLASH_RE.match(text)
    if m:
        day, month, year = m.groups()
        try:
            return datetime.strptime(f"{day}/{month}/{year}", "%d/%m/%Y").date(), (
                f"Thiếu dấu / - có thể bạn muốn gõ {day}/{month}/{year}"
            )
        except ValueError:
            pass
    parts = re.split(r"[/\-.]", text)
    if len(parts) == 3 and all(p.isdigit() for p in parts):
        d1, d2, y = parts
        if not (1 <= int(d2) <= 12) and 1 <= int(d1) <= 12:
            try:
                return datetime.strptime(f"{d2}/{d1}/{y}", "%d/%m/%Y").date(), (
                    f"'{text}' không phải ngày hợp lệ (tháng {d2} sai) - có thể bạn đã đảo "
                    f"ngày/tháng, ý là {d2}/{d1}/{y}"
                )
            except ValueError:
                pass
    return None, None


def _meaningful_count(row):
    """Số ô "có ý nghĩa" trong 1 dòng - khác rỗng VÀ khác 0. Dòng có <2 ô như
    vậy thường là dòng mẫu/trống của template (vd 1 cột Stt/công thức còn
    lại là số 0 do công thức tính trên vùng chưa nhập), không phải dữ liệu
    thật - loại các dòng này TRƯỚC khi rà lỗi, nếu không bảng mẫu có hàng
    nghìn dòng trống sẽ tạo ra hàng nghìn cảnh báo vô nghĩa (đã gặp thật khi
    test với file TONGHOP/PETROLIMEX trong data/ - sheet có sẵn hàng nghìn
    dòng mẫu bỏ trống cho tương lai)."""
    return sum(1 for v in row if v not in (None, "", 0, "0"))


def filter_substantive_rows(rows, row_numbers, min_meaningful=2):
    """-> (rows thật, row_numbers tương ứng, số dòng đã bỏ qua vì là dòng
    mẫu/trống - xem _meaningful_count). Gọi 1 lần rồi dùng kết quả cho cả
    basic_issues() lẫn cross_reference_issues(), để không rà lỗi trên các
    dòng không có dữ liệu thật."""
    kept_rows, kept_nums, skipped = [], [], 0
    for row, rn in zip(rows, row_numbers):
        if _meaningful_count(row) >= min_meaningful:
            kept_rows.append(row)
            kept_nums.append(rn)
        else:
            skipped += 1
    return kept_rows, kept_nums, skipped


def _is_numberish(value):
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip().replace(",", "")
    try:
        float(text)
        return True
    except ValueError:
        return False


def _is_dateish(value):
    return _parse_date_lenient(value)[0] is not None


def basic_issues(headers, rows, row_numbers):
    issues = []
    n_rows = len(rows)
    if n_rows == 0:
        return issues

    # 1. dòng trùng lặp hoàn toàn
    seen_at = {}
    for i, row in enumerate(rows):
        key = tuple(row)
        seen_at.setdefault(key, []).append(row_numbers[i])
    for key, line_nums in seen_at.items():
        if len(line_nums) > 1:
            for ln in line_nums:
                others = ", ".join(str(x) for x in line_nums if x != ln)
                issues.append({
                    "row": ln, "column": None, "severity": "canh_bao",
                    "problem": "Dòng này trùng hoàn toàn với dòng khác",
                    "fix": f"Trùng với dòng {others} - kiểm tra có bị nhập/copy 2 lần không.",
                })

    # 2-4. theo từng cột: tỉ lệ lấp đầy, kiểu dữ liệu đa số, ngày-like
    for col_idx, header in enumerate(headers):
        values = [row[col_idx] for row in rows]
        non_empty = [(i, v) for i, v in enumerate(values) if v not in (None, "")]
        if not non_empty:
            continue
        fill_ratio = len(non_empty) / n_rows
        header_says_date = any(kw in header.lower() for kw in ("ngày", "ngay", "date"))
        date_like = sum(1 for _, v in non_empty if _is_dateish(v))
        number_like = sum(1 for _, v in non_empty if _is_numberish(v))
        total_nonempty = len(non_empty)
        col_is_date = header_says_date or date_like / total_nonempty >= 0.7
        col_is_number = (not col_is_date) and number_like / total_nonempty >= 0.7

        # 2. ô trống bất thường (cột này hầu như luôn có dữ liệu) - ngưỡng
        # cao (0.95) để tránh báo nhầm các cột vốn chỉ thỉnh thoảng bỏ trống
        # một cách hợp lệ (vd ghi chú, mã số thuế không bắt buộc).
        if fill_ratio >= 0.95:
            for i, row in enumerate(rows):
                if row[col_idx] in (None, ""):
                    issues.append({
                        "row": row_numbers[i], "column": header, "severity": "canh_bao",
                        "problem": f"Ô trống, trong khi cột \"{header}\" {fill_ratio:.0%} dòng khác đều có dữ liệu",
                        "fix": "Kiểm tra lại - có thể bị sót khi nhập liệu.",
                    })

        if col_is_date:
            for i, v in non_empty:
                parsed, note = _parse_date_lenient(v)
                if parsed is None:
                    issues.append({
                        "row": row_numbers[i], "column": header, "severity": "loi",
                        "problem": f"Không đọc được ngày: '{v}'",
                        "fix": "Gõ lại theo dạng dd/mm/yyyy.",
                    })
                elif note:
                    issues.append({
                        "row": row_numbers[i], "column": header, "severity": "canh_bao",
                        "problem": f"Ngày '{v}' có vẻ gõ sai định dạng",
                        "fix": note,
                    })
        elif col_is_number:
            for i, v in non_empty:
                if not _is_numberish(v):
                    issues.append({
                        "row": row_numbers[i], "column": header, "severity": "loi",
                        "problem": f"Cột \"{header}\" toàn số, riêng ô này là chữ: '{v}'",
                        "fix": "Kiểm tra lại - có thể gõ nhầm hoặc dính ký tự thừa.",
                    })
        else:
            for i, v in non_empty:
                if isinstance(v, str) and (v != v.strip() or _DOUBLE_SPACE_RE.search(v)):
                    issues.append({
                        "row": row_numbers[i], "column": header, "severity": "canh_bao",
                        "problem": f"Thừa khoảng trắng: '{v}'",
                        "fix": "Xóa khoảng trắng đầu/cuối hoặc khoảng trắng đôi trong ô.",
                    })

    return issues


def cross_reference_issues(headers, rows, row_numbers, model, mapping):
    """mapping: dict - các khóa "giao_vien"/"khoa_lop"/"bien_so" trỏ tới chỉ
    mục cột (hoặc None/không có nếu người dùng không chọn cột đó)."""
    issues = []
    gv_col = mapping.get("giao_vien")
    khoa_col = mapping.get("khoa_lop")
    xe_col = mapping.get("bien_so")

    for i, row in enumerate(rows):
        if gv_col is not None:
            v = row[gv_col]
            if v not in (None, "") and not model.is_known_teacher(str(v)):
                goi_y = model.suggest_teachers(str(v))
                fix = "Không có trong danh sách giáo viên hiện tại của app"
                fix += f" - gần giống: {', '.join(goi_y)}." if goi_y else "."
                issues.append({
                    "row": row_numbers[i], "column": headers[gv_col], "severity": "loi",
                    "problem": f"Giáo viên lạ: '{v}'", "fix": fix,
                })
        if khoa_col is not None:
            v = row[khoa_col]
            if v not in (None, "") and not model.is_known_course_or_lop(str(v)):
                cung_lop = model.courses_for_lop(str(v))
                if cung_lop:
                    ds = ", ".join(c["ma_khoa"] for c in cung_lop)
                    fix = f"'{v}' là mã lớp, không phải mã khóa - mã khóa tương ứng đang có: {ds}."
                else:
                    fix = "Không khớp mã khóa lẫn mã lớp nào trong dữ liệu khóa hiện tại của app."
                issues.append({
                    "row": row_numbers[i], "column": headers[khoa_col], "severity": "loi",
                    "problem": f"Khóa/lớp lạ: '{v}'", "fix": fix,
                })
        if xe_col is not None:
            v = row[xe_col]
            if v not in (None, "") and not model.is_known_plate(str(v)):
                issues.append({
                    "row": row_numbers[i], "column": headers[xe_col], "severity": "loi",
                    "problem": f"Biển số lạ: '{v}'",
                    "fix": "Không khớp biển số xe nào trong dữ liệu khóa hiện tại của app - kiểm tra lại số xe.",
                })
    return issues

"""Đọc các file "BIÊN CHẾ XE KHÓA xxxx.xlsx/.xls" (mỗi khóa 1 file, xuất từ
hệ thống quản lý, hỗ trợ cả định dạng Excel 97-2003) sang data.json dùng
chung cho toàn bộ app.

Mỗi file có 3 sheet:
- "DS": chỉ có số thứ tự, không dùng.
- "Sổ phân công GV": có cột "Thời gian" (Từ ngày... đến ngày...) theo khóa -
  dùng để lấy tu_ngay/den_ngay (sheet chính không còn ghi ngày bắt đầu/kết
  thúc trong tiêu đề như trước).
- Sheet trùng tên mã lớp (vd "0626"): có dòng tiêu đề "BIÊN CHẾ XE KHÓA
  <ma_khoa> (<lop>)" rồi tới bảng học viên, cột E "xe/Gv/Sđt" gộp 3 dòng
  biển số / tên giáo viên / số điện thoại cho nhóm học viên đi theo xe đó.

Không còn sheet lịch học thực hành theo buổi (lịch học giờ nằm ở file PDF
riêng, xem attach_files.py) nên trường "lich_hoc" luôn để trống.
"""
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import openpyxl

import attach_files
import xls_compat

TITLE_RE = re.compile(r"BI[ÊE]N CH[ẾE] XE KH[ÓO]A\s+(.+?)\s*\(([^)]+)\)", re.UNICODE | re.IGNORECASE)
PERIOD_RE = re.compile(r"T[ừu]\s*([0-9/]+).*?đ[ếe]n\s*([0-9/]+)", re.UNICODE | re.IGNORECASE | re.DOTALL)


def clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    # loại artifact kiểu "...TÊNB18:B20" lẫn vào nội dung do lỗi nhập liệu nguồn
    text = re.sub(r"[A-F]{1,2}\d+:[A-F]{1,2}\d+$", "", text).strip()
    return text


_CA_SUFFIX_RE = re.compile(r"[\s.\-]*\(?\s*ca[\s.\-]*(\d+)\s*\)?\s*$", re.IGNORECASE)


def split_plate_ca(value):
    """Tách hậu tố ca (1 xe chạy nhiều ca trong cùng khóa, mỗi ca Excel ghi
    riêng 1 dòng xe/GV/SĐT) ra khỏi biển số - cùng 1 xe thì phải gộp lại 1
    biển số, không tách thành nhiều xe khác nhau, nhưng vẫn giữ lại "ca mấy"
    riêng để hiện ghi chú. Chấp nhận mọi kiểu người dùng hay ghi: có ngoặc
    tròn hay không ('(Ca 1)'/'Ca 1'), cách bởi dấu cách/gạch ngang/chấm hay
    ghi liền không cách ('43B25295-Ca1', '43B25295Ca2'...). Trả về
    (biển_số, ca) - ca rỗng nếu không có hậu tố."""
    text = clean_text(value)
    m = _CA_SUFFIX_RE.search(text)
    if not m:
        return text, ""
    ca = f"Ca {int(m.group(1))}"
    return text[: m.start()].strip(), ca


def clean_plate(value):
    """Chuẩn hoá biển số: bỏ hậu tố '(Ca N)' (xem split_plate_ca) rồi bỏ dấu
    '-' và '.' (Excel hay nhập kiểu '43B-252.95') để đồng bộ với định dạng
    biển số bên DSXTL_*.xml (không có 2 ký tự này, vd '43B25295')."""
    text, _ca = split_plate_ca(value)
    return text.replace("-", "").replace(".", "")


_PLATE_PROVINCE_RE = re.compile(r"^(43|92)[A-Za-z0-9]")


def _looks_like_plate(text):
    """Biển số xe ở Đà Nẵng bắt đầu bằng 43, Quảng Nam bắt đầu bằng 92 - dùng
    để nhận ra dòng nào thực sự là biển số khi file nguồn nhập đảo vị trí
    xe/giáo viên trong cột "xe/Gv/Sđt" (xem _parse_groups)."""
    digits_only = (text or "").replace("-", "").replace(".", "").replace(" ", "")
    return bool(_PLATE_PROVINCE_RE.match(digits_only))


_THAY_PREFIX_RE = re.compile(r"^(thầy|thay)\s+", re.IGNORECASE | re.UNICODE)


def _strip_thay_prefix(name):
    """Bỏ chữ 'Thầy' nhập thừa trước tên giáo viên (vd 'Thầy Nguyễn Văn
    Cường' -> 'Nguyễn Văn Cường') để khớp đúng với tên chính thức bên
    DSGV."""
    return _THAY_PREFIX_RE.sub("", name or "").strip()


def _find_course_sheet(wb):
    """Tìm sheet có dòng tiêu đề 'BIÊN CHẾ XE KHÓA ...'. Trả về (ws, title_text)."""
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 10), max_col=1):
            v = row[0].value
            if v and str(v).strip().upper().startswith("BIÊN CHẾ XE KHÓA"):
                return ws, str(v).strip()
    return None, None


def _parse_groups(ws, lop="", ket_qua_thi=None):
    """Đọc bảng học viên bắt đầu từ dòng có 'Stt' tới hết sheet.
    Trả về list groups: [{giao_vien, xe, ca, sdt, cell_ref, hoc_vien: [...]}].

    cell_ref của mỗi group là vị trí ô "xe/Gv/Sđt" gộp (dạng "<sheet>!E<dòng>")
    - dùng để báo lỗi/cảnh báo kèm vị trí chính xác trong Excel (xem
    publish_data.run_pipeline).

    ket_qua_thi (tùy chọn): dict đối chiếu kết quả thi tốt nghiệp - xem
    load_ket_qua_thi(). Nếu có, mỗi học viên được tra theo (tên, ngày sinh,
    lớp); khớp được thì coi là đã ĐẠT (file nguồn chỉ liệt kê học viên đạt)
    và lấy luôn ngày tốt nghiệp, không khớp thì để trống (chưa có kết quả,
    KHÔNG phải trượt)."""
    ket_qua_thi = ket_qua_thi or {}
    groups = []
    current_group = None
    started = False
    sheet_title = getattr(ws, "title", "")

    for row_num, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=6), start=1):
        a_val = row[0].value
        a_str = str(a_val).strip() if a_val is not None else ""

        if not started:
            if a_str == "Stt":
                started = True
            continue

        if isinstance(a_val, (int, float)):
            stt = int(a_val)
            ho_ten = clean_text(row[1].value)
            ngay_sinh = clean_text(row[2].value)
            dia_chi = clean_text(row[3].value)
            xe_gv_sdt = row[4].value
            ghi_chu = clean_text(row[5].value)

            if xe_gv_sdt and str(xe_gv_sdt).strip():
                lines = [clean_text(x) for x in str(xe_gv_sdt).split("\n")]
                lines = [x for x in lines if x]
                # Dòng xe và giáo viên bị đảo vị trí ở một số file nguồn -
                # nhận diện dòng nào thực sự là biển số qua số đầu (xem
                # _looks_like_plate) để gán đúng xe/giao_vien dù bị đảo. Hậu
                # tố "(Ca N)" có thể nằm ở dòng xe HOẶC dòng tên giáo viên
                # (tùy file), nên tách ở cả 2 dòng rồi lấy dòng nào có ca.
                line0_text, line0_ca = split_plate_ca(lines[0]) if len(lines) > 0 else ("", "")
                line1 = lines[1] if len(lines) > 1 else ""
                line1_text, line1_ca = split_plate_ca(line1)
                if not _looks_like_plate(line0_text) and _looks_like_plate(line1_text):
                    xe_text, giao_vien_text = line1_text, line0_text
                    ca = line1_ca or line0_ca
                else:
                    xe_text, giao_vien_text = line0_text, line1_text
                    ca = line0_ca or line1_ca
                xe = xe_text.replace("-", "").replace(".", "")
                giao_vien = _strip_thay_prefix(giao_vien_text)
                sdt = lines[2] if len(lines) > 2 else ""
                current_group = {
                    "giao_vien": giao_vien,
                    "xe": xe,
                    "ca": ca,  # "Ca 1"/"Ca 2" nếu 1 xe chạy nhiều ca trong khóa, rỗng nếu không
                    "sdt": sdt,
                    "cell_ref": f"{sheet_title}!E{row_num}",
                    "hoc_vien": [],
                }
                groups.append(current_group)

            if current_group is not None:
                key = (attach_files.normalize_name(ho_ten), ngay_sinh, lop)
                ngay_tot_nghiep = ket_qua_thi.get(key, "")
                current_group["hoc_vien"].append(
                    {
                        "stt": stt,
                        "ho_ten": ho_ten,
                        "ngay_sinh": ngay_sinh,
                        "dia_chi": dia_chi,
                        "ghi_chu": ghi_chu,
                        "ket_qua_thi": "Đạt" if ngay_tot_nghiep else "",
                        "ngay_tot_nghiep": ngay_tot_nghiep,
                    }
                )

    return groups


def _find_period(wb, ma_khoa):
    """Lấy (tu_ngay, den_ngay) từ sheet 'Sổ phân công GV' theo mã khóa."""
    if "Sổ phân công GV" not in wb.sheetnames:
        return "", ""
    ws = wb["Sổ phân công GV"]
    header_seen = False
    for row in ws.iter_rows(values_only=True):
        if not header_seen:
            if row and str(row[0] or "").strip() == "Số TT":
                header_seen = True
            continue
        if len(row) < 4:
            continue
        khoa = str(row[3] or "").strip()
        if khoa == ma_khoa:
            thoi_gian = str(row[2] or "")
            m = PERIOD_RE.search(thoi_gian)
            if m:
                return m.group(1).strip(), m.group(2).strip()
    return "", ""


def _load_workbook(path):
    """Mở workbook, hỗ trợ cả .xlsx (openpyxl) và .xls 97-2003 (xlrd)."""
    if Path(path).suffix.lower() == ".xls":
        return xls_compat.load_workbook(path)
    return openpyxl.load_workbook(path, data_only=True)


def load_course_from_file(path, ket_qua_thi=None):
    """Đọc 1 file 'BIÊN CHẾ XE KHÓA xxxx.xlsx/.xls', trả về 1 course dict."""
    wb = _load_workbook(path)
    ws, title = _find_course_sheet(wb)
    if ws is None:
        raise ValueError(
            f"Không tìm thấy dòng tiêu đề 'BIÊN CHẾ XE KHÓA ...' trong file: {path}"
        )

    m = TITLE_RE.search(title)
    if m:
        ma_khoa, lop = m.group(1).strip(), m.group(2).strip()
    else:
        ma_khoa, lop = title, ""

    groups = _parse_groups(ws, lop, ket_qua_thi)
    so_hoc_vien = sum(len(g["hoc_vien"]) for g in groups)
    tu_ngay, den_ngay = _find_period(wb, ma_khoa)

    return {
        "ma_khoa": ma_khoa,
        "lop": lop,
        "tu_ngay": tu_ngay,
        "den_ngay": den_ngay,
        "hv_per_xe": None,
        "so_hoc_vien": so_hoc_vien,
        "groups": groups,
        "lich_hoc": [],  # lịch học giờ chỉ còn ở file PDF, xem attach_files.py
        "source_file": Path(path).name,  # để báo cảnh báo rõ nguồn khi dữ liệu không khớp
        "source_file_path": str(path),  # đường dẫn đầy đủ - để mở file từ hộp thoại cảnh báo
    }


def load_ket_qua_thi(path, sheet_name="TỔNGHỢP"):
    """Đọc file "SỐ LƯỢNG THI ĐẠT TỐT NGHIỆP ...xlsx/.xls" (sheet tổng hợp -
    MỖI DÒNG LÀ 1 HỌC VIÊN ĐÃ ĐẠT, có cột Họ tên/Ngày sinh/Khóa/Ngày TN) để
    đối chiếu kết quả thi cho học viên bên 'BIÊN CHẾ XE KHÓA ...'. Trả về
    dict (tên đã chuẩn hoá, ngày sinh, lớp) -> ngày tốt nghiệp. Không có
    trong dict này nghĩa là CHƯA CÓ kết quả (chưa thi/chưa đạt), không suy
    ra là trượt."""
    wb = _load_workbook(path)
    sheet = sheet_name if sheet_name in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    result = {}
    for row in ws.iter_rows(min_row=2, max_col=8):
        ho_ten = clean_text(row[1].value)
        ngay_sinh = clean_text(row[2].value)
        khoa = clean_text(row[5].value)
        ngay_tn = clean_text(row[7].value)
        if not ho_ten or not khoa:
            continue
        lop = khoa.split(",")[0].strip()
        key = (attach_files.normalize_name(ho_ten), ngay_sinh, lop)
        result[key] = ngay_tn
    return result


def build_data(excel_paths, ket_qua_thi=None):
    """Đọc & gộp nhiều file 'BIÊN CHẾ XE KHÓA xxxx.xlsx' (mỗi file 1 khóa).
    Nếu 2 file trùng mã khóa, file đọc sau ghi đè file đọc trước (trả thêm
    danh sách mã bị ghi đè). ket_qua_thi (tùy chọn): xem load_ket_qua_thi()."""
    courses_by_code = {}
    overwritten = []
    for path in excel_paths:
        c = load_course_from_file(path, ket_qua_thi)
        if c["ma_khoa"] in courses_by_code:
            overwritten.append(c["ma_khoa"])
        courses_by_code[c["ma_khoa"]] = c

    courses = [courses_by_code[k] for k in sorted(courses_by_code)]
    data = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_files": [Path(p).name for p in excel_paths],
        "courses": courses,
    }
    return data, overwritten


def main():
    paths = [Path(p) for p in sys.argv[1:]]
    if not paths:
        print("Dùng: python convert_excel_to_json.py <file1.xlsx> [file2.xlsx ...]")
        sys.exit(1)
    output_path = Path("data.json")
    data, overwritten = build_data(paths)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: exported {len(data['courses'])} khoa -> {output_path}")
    if overwritten:
        print(f"CANH BAO: trung ma khoa: {overwritten}")


if __name__ == "__main__":
    main()

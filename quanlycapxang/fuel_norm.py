"""Đọc file "BẢNG ĐỊNH MỨC NL.xlsx" (định mức nhiên liệu) - 1 sheet, mỗi dòng
gồm: STT, SỐ XE (biển số), HỌ VÀ TÊN (giáo viên), ĐỊNH MỨC (lít xăng/học
viên cho xe đó). 2 cột "SỐ HỌC VIÊN" và "TỔNG ĐM" trong file gốc để trống -
app tự tính (số học viên đã lọc sẵn "Bảo lưu" ngay từ docs/cap_xang_data.json,
xem quanlygiaovien/publish_data.py) rồi điền vào tab "ĐỊNH MỨC XĂNG", không
sửa vào file gốc.

Cùng 1 xe có thể xuất hiện nhiều dòng (xe đổi giáo viên/định mức qua từng
đợt) - dòng nằm CÀNG DƯỚI trong file coi là càng MỚI, nên khi tra theo xe
(không phân biệt giáo viên) sẽ lấy định mức của dòng dưới cùng khớp biển số
đó. Nếu dòng còn khớp cả đúng giáo viên hiện đang dạy xe đó thì ưu tiên dùng
định mức của đúng cặp (xe, giáo viên) đó trước.

File Excel này KHÔNG liên quan tới dữ liệu đào tạo (BIÊN CHẾ XE/DSGV/DSXTL)
của app quản lý lái xe - là input RIÊNG, chỉ phục vụ việc tính xăng, nên vẫn
giữ import Excel trực tiếp ở đây (không có trong docs/cap_xang_data.json)."""
import json
import re

import openpyxl

import attach_files
from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent
SHARED_DATA_DIR = REPO_ROOT / "shared_data"
# Nằm trong shared_data/ (đồng bộ qua git, xem main.py nút "Đẩy phiếu/định
# mức lên") để máy khác cũng dùng được định mức vừa nhập Excel, không cần
# tự import lại file Excel gốc trên từng máy.
CACHE_FILE = SHARED_DATA_DIR / "dinh_muc_xang.json"

# Loại bỏ artifact kiểu "...TÊNB18:B20" lẫn vào nội dung do lỗi nhập liệu
# nguồn - xem quanlygiaovien/convert_excel_to_json.py (clean_text gốc).
_ARTIFACT_RE = re.compile(r"[A-F]{1,2}\d+:[A-F]{1,2}\d+$")


def clean_text(value):
    if value is None:
        return ""
    text = str(value).strip()
    text = _ARTIFACT_RE.sub("", text).strip()
    return text


def parse_dinh_muc_file(path):
    """Đọc sheet đầu tiên của file định mức nhiên liệu, trả về list dict
    [{"stt", "so_xe", "ho_ten", "dinh_muc"}, ...] theo đúng thứ tự trong file."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    all_rows = list(ws.iter_rows(values_only=True))

    header_idx = next(
        (i for i, row in enumerate(all_rows)
         if row and clean_text(row[0]).upper() == "STT"),
        None,
    )
    if header_idx is None:
        raise ValueError(
            "Không tìm thấy dòng tiêu đề 'STT' trong file định mức nhiên liệu."
        )

    rows = []
    for row in all_rows[header_idx + 1:]:
        if not row or row[0] in (None, ""):
            continue
        try:
            stt = int(row[0])
        except (TypeError, ValueError):
            continue
        so_xe = clean_text(row[1]) if len(row) > 1 else ""
        ho_ten = clean_text(row[2]) if len(row) > 2 else ""
        dinh_muc_raw = row[3] if len(row) > 3 else None
        if not so_xe or not ho_ten or dinh_muc_raw in (None, ""):
            continue
        try:
            dinh_muc = float(dinh_muc_raw)
        except (TypeError, ValueError):
            continue
        rows.append({"stt": stt, "so_xe": so_xe, "ho_ten": ho_ten, "dinh_muc": dinh_muc})
    return rows


def build_lookup(rows):
    """Trả về (by_plate, by_plate_teacher) - dict tra định mức theo biển số
    chuẩn hoá (dòng dưới cùng của file thắng nếu trùng) và theo cặp (biển
    số, tên giáo viên chuẩn hoá)."""
    by_plate = {}
    by_plate_teacher = {}
    for r in rows:
        plate = attach_files.normalize_plate(r["so_xe"])
        name = attach_files.normalize_name(r["ho_ten"])
        if plate:
            by_plate[plate] = r["dinh_muc"]
            if name:
                by_plate_teacher[(plate, name)] = r["dinh_muc"]
    return by_plate, by_plate_teacher


def load_and_cache(path):
    """Đọc file định mức, lưu cache cục bộ (dinh_muc_xang.json) để nạp lại
    nhanh ở lần mở app sau, không cần chọn lại file. Trả về dict đã lưu."""
    rows = parse_dinh_muc_file(path)
    data = {"source_file": str(path), "rows": rows}
    SHARED_DATA_DIR.mkdir(exist_ok=True)
    CACHE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def load_cache():
    if not CACHE_FILE.exists():
        return {"source_file": "", "rows": []}
    return json.loads(CACHE_FILE.read_text(encoding="utf-8"))

"""Lưu trữ định mức SỬA TAY theo từng (mã khóa, xe, giáo viên) - shared_data/
dinh_muc_override.json - KHÁC với fuel_norm.py (định mức GỐC đọc từ file
"BẢNG ĐỊNH MỨC NL.xlsx", chỉ mang tính tham khảo). CapXangModel._allowance_
breakdown() ưu tiên dùng override ở đây (nếu có, cho TỪNG giáo viên/ca) thay
cho định mức tính từ Excel - xem cap_xang_model.py.

Có GIÁO VIÊN trong key (không chỉ mã khóa + xe) vì 1 xe có thể dùng chung
bởi NHIỀU giáo viên/ca trong cùng 1 khóa (vd ca 1 thầy A, ca 2 thầy B) - sửa
tay định mức của thầy A không được ảnh hưởng tới phần của thầy B; đồng thời
đã có mã khóa trong key nên định mức sửa tay cũng tự nhiên khác nhau được
giữa các khóa khác nhau, dù cùng 1 xe/giáo viên đó.

Nằm trong shared_data/ (đồng bộ qua git_sync.py, nút "⬆ Đẩy phiếu/định mức
lên") giống phieu_cap_xang.json/dinh_muc_xang.json, để máy khác cũng thấy
đúng định mức đã sửa tay."""
import json
from datetime import datetime

import attach_files
from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent
SHARED_DATA_DIR = REPO_ROOT / "shared_data"
STORE_FILE = SHARED_DATA_DIR / "dinh_muc_override.json"


def _key(ma_khoa, xe, giao_vien):
    return (
        f"{ma_khoa}||{attach_files.normalize_plate(xe)}||"
        f"{attach_files.normalize_name(giao_vien)}"
    )


def load():
    if not STORE_FILE.exists():
        return {"overrides": {}}
    data = json.loads(STORE_FILE.read_text(encoding="utf-8"))
    data.setdefault("overrides", {})
    return data


def _save(data):
    SHARED_DATA_DIR.mkdir(exist_ok=True)
    STORE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def get_override(ma_khoa, xe, giao_vien):
    """-> {"tong_lit", "ghi_chu", "updated_at"} hoặc None nếu (khóa, xe,
    giáo viên) này chưa bị sửa tay."""
    return load()["overrides"].get(_key(ma_khoa, xe, giao_vien))


def set_override(ma_khoa, xe, giao_vien, tong_lit, ghi_chu=""):
    data = load()
    data["overrides"][_key(ma_khoa, xe, giao_vien)] = {
        "ma_khoa": ma_khoa, "xe": xe, "giao_vien": giao_vien,
        "tong_lit": float(tong_lit), "ghi_chu": ghi_chu.strip(),
        "updated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }
    _save(data)


def clear_override(ma_khoa, xe, giao_vien):
    data = load()
    data["overrides"].pop(_key(ma_khoa, xe, giao_vien), None)
    _save(data)

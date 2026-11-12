"""Lưu trữ phiếu cấp xăng (phieu_cap_xang.json) - sổ ghi mỗi lần THỰC TẾ cấp
xăng cho giáo viên theo khóa học/xe, dùng cho tab CẤP XĂNG. Khác với
fuel_norm.py (định mức - mức ĐƯỢC PHÉP cấp) và dinh_muc_xang.json (cache từ
file Excel định mức) - file này là sổ nhập tay, lưu local, không import từ
Excel."""
import json

from runtime_paths import app_dir

APP_DIR = app_dir()
STORE_FILE = APP_DIR / "phieu_cap_xang.json"


def load():
    if not STORE_FILE.exists():
        return {"next_id": 1, "records": []}
    data = json.loads(STORE_FILE.read_text(encoding="utf-8"))
    data.setdefault("next_id", 1)
    data.setdefault("records", [])
    return data


def _save(data):
    STORE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def add_record(fields):
    """fields: dict gồm giao_vien, ma_khoa, lop, xe, ngay_cap (dd/mm/yyyy),
    so_phieu, mat_hang, so_luong (float), ghi_chu, phan_loai_so. Gán id tự
    tăng, lưu lại, trả về record vừa thêm (kèm "id")."""
    data = load()
    record = {"id": data["next_id"], **fields}
    data["records"].append(record)
    data["next_id"] += 1
    _save(data)
    return record


def update_record(record_id, fields):
    """Cập nhật 1 phiếu đã có (sửa trực tiếp trên bảng - vd sửa số lượng/ghi
    chú 1 phiếu cũ)."""
    data = load()
    for r in data["records"]:
        if r.get("id") == record_id:
            r.update(fields)
            break
    _save(data)


def delete_record(record_id):
    data = load()
    data["records"] = [r for r in data["records"] if r.get("id") != record_id]
    _save(data)


def distinct_values(field):
    """Các giá trị khác nhau đã từng nhập cho 1 trường (vd "mat_hang",
    "phan_loai_so") - theo thứ tự xuất hiện gần đây nhất trước - dùng gợi ý
    trong combobox, không cần đặt sẵn danh mục cố định."""
    data = load()
    seen = []
    for r in reversed(data["records"]):
        v = (r.get(field) or "").strip()
        if v and v not in seen:
            seen.append(v)
    return seen

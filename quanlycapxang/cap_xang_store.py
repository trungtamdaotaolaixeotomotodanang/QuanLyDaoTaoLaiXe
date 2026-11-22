"""Lưu trữ phiếu cấp xăng (shared_data/phieu_cap_xang.json) - sổ ghi mỗi lần
THỰC TẾ cấp xăng cho giáo viên theo khóa học/xe, dùng cho tab CẤP XĂNG. Khác
với fuel_norm.py (định mức - mức ĐƯỢC PHÉP cấp) và dinh_muc_xang.json (cache
từ file Excel định mức) - file này là sổ nhập tay.

Nằm trong shared_data/ (cố định theo REPO_ROOT, không theo APP_DIR - xem
quanlygiaovien/publish_data.py để biết lý do) để đồng bộ được qua git giữa
nhiều máy - xem main.py nút "⬆ Đẩy phiếu/định mức lên" (chủ động, không tự
push mỗi lần gõ 1 ô, vì ghi phiếu có thể rất nhiều lần/ngày).

id mỗi phiếu là UUID (không phải số tự tăng) - để 2 máy cùng thêm phiếu mới
trước khi đồng bộ với nhau không bao giờ bị trùng id, tránh ghi đè nhầm lên
nhau lúc gộp qua git."""
import json
import uuid

from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent
SHARED_DATA_DIR = REPO_ROOT / "shared_data"
STORE_FILE = SHARED_DATA_DIR / "phieu_cap_xang.json"
# Ghi đè lại mỗi lần xóa phiếu (delete_records) - lưu đúng trạng thái NGAY
# TRƯỚC lần xóa gần nhất, để nút "Hoàn tác" ở tab TỔNG HỢP CẤP XĂNG khôi phục
# lại được nếu lỡ xóa nhầm (vd xóa hàng loạt phiếu "mồ côi") - xem restore_backup().
BACKUP_FILE = SHARED_DATA_DIR / "phieu_cap_xang.json.backup"


def load():
    if not STORE_FILE.exists():
        return {"records": []}
    data = json.loads(STORE_FILE.read_text(encoding="utf-8"))
    data.setdefault("records", [])
    return data


def _save(data):
    SHARED_DATA_DIR.mkdir(exist_ok=True)
    STORE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def add_record(fields):
    """fields: dict gồm giao_vien, ma_khoa, lop, xe, ngay_cap (dd/mm/yyyy),
    so_phieu, mat_hang, so_luong (float), ghi_chu, phan_loai_so. Gán id UUID,
    lưu lại, trả về record vừa thêm (kèm "id")."""
    data = load()
    record = {"id": uuid.uuid4().hex, **fields}
    data["records"].append(record)
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


def delete_records(record_ids):
    """Xóa nhiều phiếu cùng lúc (xem tab_tong_hop_cap_xang.py _on_delete_selected).
    Tự backup toàn bộ dữ liệu TRƯỚC khi xóa vào BACKUP_FILE (ghi đè bản cũ) -
    xem restore_backup()."""
    ids = set(record_ids)
    data = load()
    BACKUP_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    data["records"] = [r for r in data["records"] if r.get("id") not in ids]
    _save(data)


def delete_record(record_id):
    delete_records([record_id])


def restore_backup():
    """Khôi phục lại đúng trạng thái ngay TRƯỚC lần delete_records() gần nhất
    - dùng cho nút "Hoàn tác". Trả về False nếu chưa có backup nào (chưa xóa
    phiếu nào trong phiên làm việc hiện tại)."""
    if not BACKUP_FILE.exists():
        return False
    STORE_FILE.write_text(BACKUP_FILE.read_text(encoding="utf-8"), encoding="utf-8")
    return True


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

"""Nhập phiếu cấp xăng từ file JSON xuất bởi XacThucCapXang (xem
export_phieu_logic.py) - dùng trong lúc quanlycapxang còn đang phát triển,
nhân viên vẫn ghi sổ cấp xăng bằng tay (Excel) ở nơi khác, tool đó đối chiếu
+ xuất JSON, còn ở đây chỉ cần ĐỌC + GHI vào cap_xang_store, không cần đối
chiếu lại (đã đối chiếu rồi lúc xuất).

Chống nhập trùng (vd xuất/nhập JSON cùng 1 sổ nhiều lần, hoặc sổ đã từng
được nhập tay 1 lần trước đó) bằng khóa tự nhiên (xe, ngay_cap, so_phieu) -
KHÔNG gồm ma_khoa: số phiếu đánh theo 1 sổ/quyển DUY NHẤT của cả trung tâm
(không theo từng khóa), nên (biển số, ngày cấp, số phiếu) đã xác định đúng 1
lần cấp xăng thực tế. Để ma_khoa trong khóa trùng dễ gây nhập lại 2 lần
CÙNG 1 phiếu: phiếu nhập tay trước đó có thể bị để trống ma_khoa (chưa khớp
được khóa lúc đó - xem "phiếu mồ côi"), còn bản xuất JSON mới (đối chiếu lại
bằng validate1_logic) có thể đã khớp ra đúng ma_khoa - 2 bản NHƯ VẬY là CÙNG
1 phiếu thật, không phải 2 phiếu khác nhau. Không dùng "id" vì JSON nguồn
không có id khớp với cap_xang_store (mỗi lần add_record tự gán uuid mới)."""
import json

import cap_xang_store


def _dedup_key(record):
    return (
        record.get("xe", ""),
        record.get("ngay_cap", ""),
        record.get("so_phieu", ""),
    )


def import_from_file(path):
    """-> (added: int, skipped_trung: int, errors: list[str])."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    new_records = data.get("records", [])
    existing = cap_xang_store.load()["records"]
    existing_keys = {_dedup_key(r) for r in existing}

    added = 0
    skipped_trung = 0
    errors = []
    for r in new_records:
        key = _dedup_key(r)
        if key in existing_keys:
            skipped_trung += 1
            continue
        fields = {k: v for k, v in r.items() if k != "id"}
        try:
            cap_xang_store.add_record(fields)
        except Exception as exc:
            errors.append(f"Phiếu {r.get('so_phieu', '?')} (khóa {r.get('ma_khoa', '?')}): {exc}")
            continue
        existing_keys.add(key)
        added += 1

    return added, skipped_trung, errors

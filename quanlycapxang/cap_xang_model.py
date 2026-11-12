"""Đọc docs/cap_xang_data.json (xuất bởi quanlygiaovien/publish_data.py,
đồng bộ qua git - KHÔNG đọc Excel/XML trực tiếp nữa, xem dữ liệu.md ở
quanlygiaovien/) + dinh_muc_xang.json (cache định mức nhiên liệu, xem
fuel_norm.py) - cung cấp các hàm tra cứu cho 3 tab CẤP XĂNG/ĐỊNH MỨC XĂNG/
TỔNG HỢP CẤP XĂNG."""
import json
from datetime import datetime
from pathlib import Path

import attach_files
import fuel_norm


def _parse_date(text):
    text = (text or "").strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


class CapXangModel:
    def __init__(self, json_path):
        self.json_path = Path(json_path)
        self.reload()

    def reload(self):
        """Đọc lại docs/cap_xang_data.json + cache định mức từ đĩa - gọi lúc
        mở app và sau khi bấm "Lấy dữ liệu mới nhất" (xem main.py)."""
        if self.json_path.exists():
            raw = json.loads(self.json_path.read_text(encoding="utf-8"))
        else:
            raw = {"generated_at": "", "courses": []}
        self.generated_at = raw.get("generated_at", "")
        self.courses = raw.get("courses", [])
        self.all_teacher_names = set(raw.get("teachers", []))

        dinh_muc_cache = fuel_norm.load_cache()
        self.dinh_muc_source_file = dinh_muc_cache.get("source_file", "")
        self.dinh_muc_by_plate, self.dinh_muc_by_plate_teacher = fuel_norm.build_lookup(
            dinh_muc_cache.get("rows", [])
        )

    def status_of(self, course):
        today = datetime.now()
        start = _parse_date(course.get("tu_ngay", ""))
        end = _parse_date(course.get("den_ngay", ""))
        if end and today > end:
            return "Đã đào tạo"
        if start and today < start:
            return "Chưa đào tạo"
        return "Đang đào tạo"

    def all_teachers(self):
        """Danh sách tên giáo viên duy nhất, đã sắp xếp theo tên (không
        theo họ) - gồm TOÀN BỘ giáo viên trong DSGV (kể cả giáo viên hiện
        không dạy khóa nào), union thêm tên xuất hiện trong khóa (phòng khi
        "teachers" thiếu/cũ hơn "courses")."""
        names = set(self.all_teacher_names)
        for c in self.courses:
            for g in c.get("groups", []):
                if g.get("giao_vien"):
                    names.add(g["giao_vien"])
        return sorted(names, key=attach_files.given_name_sort_key)

    def course_codes(self):
        return sorted({c["ma_khoa"] for c in self.courses})

    def courses_of_teacher(self, name):
        """Tất cả khóa (group) giáo viên `name` từng được phân công dạy - MỌI
        trạng thái (đang/đã/chưa đào tạo) - dùng cho tab CẤP XĂNG để chọn
        khóa học khi lập phiếu cấp xăng (so khớp tên chính xác)."""
        name_norm = attach_files.normalize_name(name)
        rows = []
        for c in self.courses:
            for g in c.get("groups", []):
                if attach_files.normalize_name(g.get("giao_vien", "")) != name_norm:
                    continue
                rows.append({
                    "ma_khoa": c["ma_khoa"],
                    "lop": c.get("lop", ""),
                    "xe": g.get("xe", ""),
                    "ca": g.get("ca", ""),
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                    "status": self.status_of(c),
                })
        return rows

    def lookup_dinh_muc(self, xe, giao_vien):
        """Tra định mức (lít/học viên) theo xe - ưu tiên khớp đúng cặp (xe,
        giáo viên) hiện đang dạy xe đó trước, không có thì lấy định mức mới
        nhất ghi cho riêng xe đó. Trả về (dinh_muc_hoac_None, "xe+gv"|"xe"|None)."""
        plate = attach_files.normalize_plate(xe or "")
        name = attach_files.normalize_name(giao_vien or "")
        val = self.dinh_muc_by_plate_teacher.get((plate, name))
        if val is not None:
            return val, "xe+gv"
        val = self.dinh_muc_by_plate.get(plate)
        if val is not None:
            return val, "xe"
        return None, None

    def fuel_rows_for_teacher(self, name=None):
        """Trả về list các khóa ĐANG DẠY (trạng thái "Đang đào tạo") của giáo
        viên `name` - hoặc TẤT CẢ giáo viên nếu `name` rỗng/None - kèm định
        mức xăng/học viên của xe, số học viên tính xăng (đã lọc sẵn "Bảo
        lưu" từ docs/cap_xang_data.json) và tổng số lít xăng của khóa đó
        (None nếu không tra được định mức cho xe)."""
        name_norm = attach_files.normalize_name(name) if name else None
        rows = []
        for c in self.courses:
            if self.status_of(c) != "Đang đào tạo":
                continue
            for g in c.get("groups", []):
                giao_vien = g.get("giao_vien", "")
                if name_norm is not None and attach_files.normalize_name(giao_vien) != name_norm:
                    continue
                so_hv = g.get("so_hoc_vien", 0)
                dinh_muc, matched_by = self.lookup_dinh_muc(g.get("xe", ""), giao_vien)
                tong = dinh_muc * so_hv if dinh_muc is not None else None
                rows.append({
                    "giao_vien": giao_vien,
                    "ma_khoa": c["ma_khoa"],
                    "lop": c.get("lop", ""),
                    "xe": g.get("xe", ""),
                    "ca": g.get("ca", ""),
                    "so_hoc_vien": so_hv,
                    "dinh_muc": dinh_muc,
                    "matched_by": matched_by,
                    "tong": tong,
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                })
        return rows

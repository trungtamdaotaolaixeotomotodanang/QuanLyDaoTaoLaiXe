"""Đọc docs/cap_xang_data.json (xuất bởi quanlygiaovien/publish_data.py,
đồng bộ qua git - KHÔNG đọc Excel/XML trực tiếp nữa, xem dữ liệu.md ở
quanlygiaovien/) + dinh_muc_xang.json (cache định mức nhiên liệu, xem
fuel_norm.py) - cung cấp các hàm tra cứu cho 3 tab CẤP XĂNG/ĐỊNH MỨC XĂNG/
TỔNG HỢP CẤP XĂNG."""
import json
from datetime import datetime
from pathlib import Path

import attach_files
import cap_xang_store
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
        self._build_assignment_index()

    def _build_assignment_index(self):
        """Chỉ mục để kiểm tra 1 phiếu cấp xăng có khớp phân công hiện tại
        không (xem orphan_reason) - khớp theo đúng bộ 3 (giáo viên, mã khóa,
        xe) như tab CẤP XĂNG gộp phiếu vào từng dòng khóa/xe."""
        self._course_codes_set = set()
        self._teacher_norms = {attach_files.normalize_name(n) for n in self.all_teacher_names}
        self._assignment_keys = set()
        for c in self.courses:
            ma_khoa = c.get("ma_khoa", "")
            self._course_codes_set.add(ma_khoa)
            for g in c.get("groups", []):
                name_norm = attach_files.normalize_name(g.get("giao_vien", ""))
                self._teacher_norms.add(name_norm)
                self._assignment_keys.add((name_norm, ma_khoa, g.get("xe", "")))

    def orphan_reason(self, record):
        """None nếu phiếu khớp đúng 1 phân công (giáo viên + khóa + xe) đang
        có trong docs/cap_xang_data.json; ngược lại trả về lý do ngắn gọn vì
        sao KHÔNG khớp ("phiếu mồ côi" - thường là phiếu cũ nhập từ Excel của
        khóa đã kết thúc/xe đã đổi, hoặc gõ nhầm mã khóa/tên). Phiếu mồ côi
        vẫn tính vào tổng, nhưng KHÔNG hiện ở lưới "Chi tiết" tab CẤP XĂNG
        (lưới đó dựng theo phân công) và không cộng vào "Đã cấp" của khóa
        nào ở tab ĐỊNH MỨC XĂNG."""
        name_norm = attach_files.normalize_name(record.get("giao_vien", ""))
        ma_khoa = record.get("ma_khoa", "")
        xe = record.get("xe", "")
        if (name_norm, ma_khoa, xe) in self._assignment_keys:
            return None
        if name_norm not in self._teacher_norms:
            return "tên giáo viên không có trong danh sách giáo viên"
        if ma_khoa not in self._course_codes_set:
            return "mã khóa không có trong danh sách khóa"
        return "khóa có nhưng không phân công xe này cho giáo viên này"

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

    def _da_cap_by_khoa_xe(self):
        """Tổng số lít đã cấp (xem cap_xang_store) theo từng cặp (mã khóa, xe)
        - khớp CHÍNH XÁC chuỗi, giống cách tab CẤP XĂNG gộp phiếu theo khóa/xe
        (xem tab_cap_xang.py _unique_assignments) - dùng cho cột "Đã cấp"/
        "Còn lại" ở tab ĐỊNH MỨC XĂNG."""
        totals = {}
        for r in cap_xang_store.load()["records"]:
            key = (r.get("ma_khoa", ""), r.get("xe", ""))
            totals[key] = totals.get(key, 0.0) + (r.get("so_luong") or 0)
        return totals

    def fuel_rows_for_teacher(self, name=None):
        """Trả về list các khóa ĐANG DẠY (trạng thái "Đang đào tạo") của giáo
        viên `name` - hoặc TẤT CẢ giáo viên nếu `name` rỗng/None - kèm định
        mức xăng/học viên của xe, số học viên tính xăng (đã lọc sẵn "Bảo
        lưu" từ docs/cap_xang_data.json), tổng số lít xăng của khóa đó (None
        nếu không tra được định mức cho xe), số lít ĐÃ CẤP thực tế (xem tab
        CẤP XĂNG) và CÒN LẠI = tổng định mức - đã cấp (None nếu chưa có định
        mức để so sánh) - đúng cột "CÒN LẠI" trước đây phải tự tính tay trên
        Excel."""
        name_norm = attach_files.normalize_name(name) if name else None
        da_cap_map = self._da_cap_by_khoa_xe()
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
                da_cap = da_cap_map.get((c["ma_khoa"], g.get("xe", "")), 0.0)
                con_lai = tong - da_cap if tong is not None else None
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
                    "da_cap": da_cap,
                    "con_lai": con_lai,
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                })
        return rows

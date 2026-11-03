"""Nạp data.json (xuất từ Excel) và cung cấp các hàm tra cứu cho app."""
import json
import re
from datetime import datetime
from pathlib import Path

import attach_files


def _parse_date(text):
    text = (text or "").strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


_HANG_PREFIX_RE = re.compile(r"^([A-Za-z0-9.]+)K", re.UNICODE)


def parse_hang(lop):
    """Suy ra hạng đào tạo từ mã lớp (cột 'Khóa' trong bảng phân hạng):
    - Có dấu phẩy (khóa nâng hạng, vd 'NHD2K0726,C-D2') -> lấy phần sau dấu
      phẩy ('C-D2').
    - Không có dấu phẩy -> lấy phần trước chữ 'K', bỏ hậu tố sau dấu chấm
      nếu có (vd 'BK0726' -> 'B', 'B.01K0726' -> 'B', 'C1K0726' -> 'C1')."""
    lop = (lop or "").strip()
    if not lop:
        return ""
    if "," in lop:
        return lop.split(",", 1)[1].strip()
    m = _HANG_PREFIX_RE.match(lop)
    if not m:
        return ""
    return m.group(1).split(".")[0]


class DataModel:
    def __init__(self, json_path, dsgv=None):
        """dsgv (tùy chọn): dict ma_gv -> {"ten": ...} từ roster_xml.parse_dsgv -
        dùng để thay tên viết tắt nhập trong Excel "biên chế xe" bằng tên đầy đủ
        chuẩn bên DSGV, và để tra cứu ra toàn bộ giáo viên (không chỉ những
        người đang có khóa trong Excel)."""
        self.json_path = Path(json_path)
        raw = json.loads(self.json_path.read_text(encoding="utf-8"))
        self.generated_at = raw.get("generated_at", "")
        self.source_files = raw.get("source_files", [])
        self.courses = raw.get("courses", [])
        self.dsgv_names = set()
        if dsgv:
            self.dsgv_names = {rec["ten"] for rec in dsgv.values() if rec.get("ten")}
            self._resolve_teacher_names(dsgv)

    def _resolve_teacher_names(self, dsgv):
        """Thay tên giáo viên trong từng khóa (có thể viết tắt trong Excel)
        bằng tên đầy đủ khớp bên DSGV - khớp hệt trước, không được thì thử
        khớp tên viết tắt (vd "Tr. N. Đ. Phương" -> "Trần Ngô Đông Phương").
        Không khớp được (mơ hồ hoặc không tìm thấy) thì giữ nguyên tên gốc."""
        name_to_magv = {
            attach_files.normalize_name(rec["ten"]): magv for magv, rec in dsgv.items()
        }
        for c in self.courses:
            for g in c.get("groups", []):
                name = g.get("giao_vien")
                if not name:
                    continue
                ma_gv = name_to_magv.get(attach_files.normalize_name(name))
                if not ma_gv:
                    candidates = attach_files.find_teacher_candidates(name, dsgv)
                    if len(candidates) == 1:
                        ma_gv = candidates[0][0]
                if ma_gv:
                    g["giao_vien"] = dsgv[ma_gv]["ten"]

    # ---- truy cập cơ bản ----
    def all_courses(self):
        return self.courses

    def course_codes(self):
        return [c["ma_khoa"] for c in self.courses]

    def find_course(self, ma_khoa):
        for c in self.courses:
            if c["ma_khoa"] == ma_khoa:
                return c
        return None

    def all_teachers(self):
        """Danh sách tên giáo viên duy nhất, đã sắp xếp - gồm toàn bộ giáo
        viên bên DSGV (nếu có), gộp thêm tên trong các khóa (phòng khi chưa
        khớp được với DSGV)."""
        names = set(self.dsgv_names)
        for c in self.courses:
            for g in c.get("groups", []):
                if g.get("giao_vien"):
                    names.add(g["giao_vien"])
        return sorted(names)

    def status_of(self, course):
        today = datetime.now()
        start = _parse_date(course.get("tu_ngay", ""))
        end = _parse_date(course.get("den_ngay", ""))
        if end and today > end:
            return "Đã đào tạo"
        if start and today < start:
            return "Chưa đào tạo"
        return "Đang đào tạo"

    # ---- tra cứu ----
    def search_by_teacher(self, name):
        """Trả về list các dòng (mỗi dòng = 1 khóa mà giáo viên này tham gia)."""
        name = (name or "").strip().lower()
        rows = []
        for c in self.courses:
            for g in c.get("groups", []):
                if name and name not in g.get("giao_vien", "").lower():
                    continue
                if not name:
                    continue
                rows.append(
                    {
                        "giao_vien": g["giao_vien"],
                        "ma_khoa": c["ma_khoa"],
                        "lop": c["lop"],
                        "xe": g["xe"],
                        "ca": g.get("ca", ""),
                        "sdt": g.get("sdt", ""),
                        "so_hoc_vien": len(g.get("hoc_vien", [])),
                        "tu_ngay": c.get("tu_ngay", ""),
                        "den_ngay": c.get("den_ngay", ""),
                        "status": self.status_of(c),
                        "course": c,
                        "group": g,
                    }
                )
        return rows

    def summarize_by_hang(self, rows):
        """Gộp các dòng kết quả tra cứu (xem search_by_teacher/search_by_course)
        theo hạng (suy ra từ mã lớp, xem parse_hang) -> tổng số học viên mỗi
        hạng. Trả về list [(hang, so_hoc_vien)] sắp theo tên hạng."""
        totals = {}
        for r in rows:
            hang = parse_hang(r.get("lop", "")) or "?"
            totals[hang] = totals.get(hang, 0) + r.get("so_hoc_vien", 0)
        return sorted(totals.items())

    def search_by_course(self, ma_khoa):
        """Trả về list các dòng (mỗi dòng = 1 giáo viên/xe trong khóa đó)."""
        c = self.find_course(ma_khoa)
        if not c:
            return []
        rows = []
        for g in c.get("groups", []):
            rows.append(
                {
                    "giao_vien": g["giao_vien"],
                    "ma_khoa": c["ma_khoa"],
                    "lop": c["lop"],
                    "xe": g["xe"],
                    "ca": g.get("ca", ""),
                    "sdt": g.get("sdt", ""),
                    "so_hoc_vien": len(g.get("hoc_vien", [])),
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                    "status": self.status_of(c),
                    "course": c,
                    "group": g,
                }
            )
        return rows

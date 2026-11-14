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
        self._courses_by_code = {}
        self._courses_by_lop = {}
        self._teacher_norms = {attach_files.normalize_name(n) for n in self.all_teacher_names}
        self._assignment_keys = set()
        for c in self.courses:
            ma_khoa = c.get("ma_khoa", "")
            self._courses_by_code[ma_khoa] = c
            self._courses_by_lop.setdefault(c.get("lop", ""), []).append(c)
            for g in c.get("groups", []):
                name_norm = attach_files.normalize_name(g.get("giao_vien", ""))
                self._teacher_norms.add(name_norm)
                self._assignment_keys.add((name_norm, ma_khoa, g.get("xe", "")))

    @staticmethod
    def _assignments_text(course, limit=6):
        pairs = []
        for g in course.get("groups", []):
            p = f"{g.get('giao_vien', '')} - {g.get('xe', '')}"
            if p not in pairs:
                pairs.append(p)
        text = "; ".join(pairs[:limit])
        return text + (f"; ... ({len(pairs)} cặp)" if len(pairs) > limit else "")

    def _suggest_teachers(self, name):
        """Tên giáo viên trong danh sách gần giống `name` (trùng tên gọi -
        từ cuối, hoặc chứa nhau sau khi chuẩn hóa) - để gợi ý sửa."""
        norm = attach_files.normalize_name(name)
        given = norm.split()[-1] if norm.split() else ""
        out = []
        for t in self.all_teachers():
            t_norm = attach_files.normalize_name(t)
            if not t_norm:
                continue
            if (given and t_norm.split()[-1] == given) or norm in t_norm or t_norm in norm:
                out.append(t)
        return out[:5]

    def orphan_reason(self, record):
        """None nếu phiếu khớp đúng 1 phân công (giáo viên + khóa + xe) đang
        có trong docs/cap_xang_data.json; ngược lại trả về dict {"field": ô
        bị sai ("giao_vien"/"ma_khoa"/"xe"), "problem": sai cái gì, "fix":
        cách sửa cụ thể (gợi ý giá trị đúng nếu suy ra được)} - để người dùng
        nhìn là biết sửa ô nào, không phải tự dò.

        Phiếu "mồ côi" thường là phiếu cũ nhập từ Excel (ô Khóa học trống,
        chỉ có mã lớp) của khóa đã kết thúc không còn trong dữ liệu phân công
        -> không sửa được và cũng không cần sửa; hoặc gõ nhầm tên/mã/xe ->
        sửa theo gợi ý. Phiếu mồ côi vẫn tính vào tổng, nhưng KHÔNG hiện ở
        lưới "Chi tiết" tab CẤP XĂNG (lưới đó dựng theo phân công) và không
        trừ vào "Còn lại" của khóa nào ở tab ĐỊNH MỨC XĂNG."""
        giao_vien = record.get("giao_vien", "")
        name_norm = attach_files.normalize_name(giao_vien)
        ma_khoa = record.get("ma_khoa", "")
        lop = record.get("lop", "")
        xe = record.get("xe", "")
        if (name_norm, ma_khoa, xe) in self._assignment_keys:
            return None

        if name_norm not in self._teacher_norms:
            goi_y = self._suggest_teachers(giao_vien)
            fix = "Sửa ô Giáo viên thành đúng tên trong danh sách giáo viên"
            fix += f" - gần giống: {', '.join(goi_y)}." if goi_y else " (không tìm thấy tên nào gần giống)."
            return {
                "field": "giao_vien",
                "problem": f"Tên giáo viên '{giao_vien}' không có trong danh sách giáo viên",
                "fix": fix,
            }

        if ma_khoa not in self._courses_by_code:
            problem = (
                f"Ô Khóa học trống (phiếu nhập từ Excel cũ, chỉ có mã lớp '{lop}')"
                if not ma_khoa else f"Mã khóa '{ma_khoa}' không có trong danh sách khóa"
            )
            cung_lop = self._courses_by_lop.get(lop, []) if lop else []
            if cung_lop:
                ds = "; ".join(
                    f"{c['ma_khoa']} (phân công: {self._assignments_text(c)})" for c in cung_lop
                )
                fix = (
                    f"Nếu đúng là khóa có lớp '{lop}' thì gõ mã khóa vào ô Khóa học: {ds}. "
                    "Giáo viên/Biển số trên phiếu cũng phải đúng như phân công đó."
                )
            else:
                fix = (
                    f"Lớp '{lop}' không còn trong dữ liệu phân công (khóa đã kết thúc trước khi "
                    "dùng app) - KHÔNG cần sửa, để nguyên: phiếu vẫn tính vào tổng/báo cáo. "
                    "Chỉ xóa nếu chắc chắn là phiếu nhập nhầm."
                )
            return {"field": "ma_khoa", "problem": problem, "fix": fix}

        course = self._courses_by_code[ma_khoa]
        gv_cua_xe = sorted({g.get("giao_vien", "") for g in course.get("groups", []) if g.get("xe", "") == xe})
        xe_cua_gv = sorted({
            g.get("xe", "") for g in course.get("groups", [])
            if attach_files.normalize_name(g.get("giao_vien", "")) == name_norm
        })
        if gv_cua_xe:
            return {
                "field": "giao_vien",
                "problem": f"Khóa {ma_khoa}: xe {xe} được phân công cho {', '.join(gv_cua_xe)}, không phải '{giao_vien}'",
                "fix": (
                    f"Nếu đúng là xe {xe} thì sửa ô Giáo viên thành '{gv_cua_xe[0]}'; nếu đúng là "
                    f"'{giao_vien}' thì sửa ô Biển số thành xe được phân công cho người này"
                    + (f": {', '.join(xe_cua_gv)}." if xe_cua_gv else " (người này không có trong khóa này).")
                ),
            }
        if xe_cua_gv:
            return {
                "field": "xe",
                "problem": f"Khóa {ma_khoa}: '{giao_vien}' được phân công xe {', '.join(xe_cua_gv)}, không phải {xe}",
                "fix": f"Sửa ô Biển số thành {xe_cua_gv[0] if len(xe_cua_gv) == 1 else 'một trong: ' + ', '.join(xe_cua_gv)}.",
            }
        return {
            "field": "ma_khoa",
            "problem": f"Khóa {ma_khoa} không phân công cả '{giao_vien}' lẫn xe {xe}",
            "fix": (
                f"Khóa này phân công: {self._assignments_text(course)}. Kiểm tra lại cả 3 ô "
                "Khóa học / Giáo viên / Biển số - có thể gõ nhầm mã khóa."
            ),
        }

    @staticmethod
    def orphan_text(diag):
        """Gộp problem + fix thành 1 đoạn đọc được (tooltip/ô Excel)."""
        return f"{diag['problem']}. Cách sửa: {diag['fix']}" if diag else ""

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

    def allowances_for_teacher(self, name):
        """Định mức/đã cấp/còn lại cho TỪNG khóa+xe giáo viên `name` từng được
        phân công (mọi trạng thái - đúng các dòng tab CẤP XĂNG liệt kê) ->
        dict (ma_khoa, xe) -> {"dinh_muc", "so_hv", "tong", "da_cap",
        "con_lai"}. Nhiều "ca" cùng khóa+xe thì cộng dồn số học viên (phiếu
        cấp xăng không phân biệt ca). Không có định mức cho xe -> tong/con_lai
        None (vẫn có da_cap). Dùng để hiện cột "Còn lại" + cảnh báo vượt định
        mức NGAY LÚC NHẬP ở tab CẤP XĂNG - thay vì phải sang tab ĐỊNH MỨC
        XĂNG mới biết (lúc đó đã cấp xong rồi)."""
        name_norm = attach_files.normalize_name(name)
        da_cap_map = self._da_cap_by_khoa_xe()
        result = {}
        for c in self.courses:
            for g in c.get("groups", []):
                if attach_files.normalize_name(g.get("giao_vien", "")) != name_norm:
                    continue
                key = (c["ma_khoa"], g.get("xe", ""))
                info = result.get(key)
                if info is None:
                    dinh_muc, _ = self.lookup_dinh_muc(g.get("xe", ""), name)
                    info = result[key] = {
                        "dinh_muc": dinh_muc, "so_hv": 0, "tong": None,
                        "da_cap": da_cap_map.get(key, 0.0), "con_lai": None,
                    }
                info["so_hv"] += g.get("so_hoc_vien", 0)
        for info in result.values():
            if info["dinh_muc"] is not None:
                info["tong"] = info["dinh_muc"] * info["so_hv"]
                info["con_lai"] = info["tong"] - info["da_cap"]
        return result

    def fuel_rows_for_teacher(self, name=None, statuses=("Đang đào tạo",)):
        """Trả về list các khóa của giáo viên `name` - hoặc TẤT CẢ giáo viên
        nếu `name` rỗng/None - có trạng thái nằm trong `statuses` (mặc định
        chỉ khóa ĐANG DẠY; None = mọi trạng thái, để quyết toán cả khóa đã
        kết thúc) - kèm định mức xăng/học viên của xe, số học viên tính xăng
        (đã lọc sẵn "Bảo lưu" từ docs/cap_xang_data.json), tổng số lít xăng
        của khóa đó (None nếu không tra được định mức cho xe), số lít ĐÃ CẤP
        thực tế (xem tab CẤP XĂNG) và CÒN LẠI = tổng định mức - đã cấp (None
        nếu chưa có định mức để so sánh) - đúng cột "CÒN LẠI" trước đây phải
        tự tính tay trên Excel."""
        name_norm = attach_files.normalize_name(name) if name else None
        da_cap_map = self._da_cap_by_khoa_xe()
        rows = []
        for c in self.courses:
            if statuses is not None and self.status_of(c) not in statuses:
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

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


def _fmt_g(value):
    return f"{value:g}" if value is not None else "—"


def format_allowance_tooltip(info):
    """Giải thích ngắn (1 dòng) cách tính "Tổng định mức" của 1 (mã khóa, xe)
    - xem CapXangModel._allowance_breakdown - dùng cho tooltip cột "Còn lại"
    ở tab CẤP XĂNG. Bản giải thích đầy đủ hơn (HTML, nhiều dòng) để hiện
    trong tab ĐỊNH MỨC XĂNG tự dựng riêng ở tab_dinh_muc_xang.py vì cần định
    dạng rich text + câu dẫn khác."""
    contribs = info["contribs"]
    if info["tong"] is None:
        thieu = ", ".join(c["giao_vien"] or "?" for c in contribs if c["dinh_muc"] is None)
        return f"Chưa tra được định mức cho: {thieu} - chưa tính được tổng định mức của xe này."
    if len(contribs) == 1:
        c = contribs[0]
        return (
            f"Định mức {_fmt_g(c['dinh_muc'])} lít/HV x {c['so_hoc_vien']} HV = {_fmt_g(info['tong'])} lít; "
            f"đã cấp {_fmt_g(info['da_cap'])} lít."
        )
    chi_tiet = "; ".join(
        f"{c['giao_vien']} ({c['ca'] or '?'}): {_fmt_g(c['dinh_muc'])} lít/HV x {c['so_hoc_vien']} HV = "
        f"{_fmt_g(c['dinh_muc'] * c['so_hoc_vien'])} lít"
        for c in contribs
    )
    return (
        f"Xe này dùng chung {len(contribs)} ca - {chi_tiet}. Cộng lại = tổng định mức "
        f"{_fmt_g(info['tong'])} lít; đã cấp {_fmt_g(info['da_cap'])} lít."
    )


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
        self._known_plates = set()
        for c in self.courses:
            ma_khoa = c.get("ma_khoa", "")
            self._courses_by_code[ma_khoa] = c
            self._courses_by_lop.setdefault(c.get("lop", ""), []).append(c)
            for g in c.get("groups", []):
                name_norm = attach_files.normalize_name(g.get("giao_vien", ""))
                self._teacher_norms.add(name_norm)
                self._assignment_keys.add((name_norm, ma_khoa, g.get("xe", "")))
                if g.get("xe"):
                    self._known_plates.add(attach_files.normalize_plate(g["xe"]))

    def is_known_teacher(self, name):
        return attach_files.normalize_name(name) in self._teacher_norms

    def is_known_course_or_lop(self, value):
        """value có thể là mã khóa (vd "48002K262024") hoặc mã lớp (vd
        "BK0726") - dùng cho tab XÁC THỰC EXCEL (file thô không phải lúc nào
        cũng ghi đúng mã khóa đầy đủ, có khi chỉ ghi mã lớp)."""
        return value in self._courses_by_code or value in self._courses_by_lop

    def is_known_plate(self, plate):
        return attach_files.normalize_plate(plate) in self._known_plates

    def courses_for_lop(self, lop):
        return self._courses_by_lop.get(lop, [])

    def known_khoa_xe_pairs(self):
        """Tập (mã khóa, xe) của MỌI khóa/nhóm hiện có (không lọc theo giáo
        viên) - đúng những cặp mà _da_cap_by_khoa_xe/fuel_rows_for_teacher
        dùng để cộng phiếu vào "Đã cấp" của 1 khóa. Khác với orphan_reason()
        (khớp CẢ BỘ BA giáo viên+khóa+xe, dùng để báo phiếu cần sửa tay) -
        đây chỉ cần khóa+xe đúng, dùng để biết 1 phiếu có được TÍNH vào Đã
        cấp của khóa nào hay không (dù giáo viên trên phiếu có ghi sai)."""
        return {(c["ma_khoa"], g.get("xe", "")) for c in self.courses for g in c.get("groups", [])}

    def overview_stats(self):
        """Số liệu tổng hợp cho tab TỔNG QUAN - tính lại từ các hàm đã có
        (orphan_reason/fuel_rows_for_teacher), không tạo thêm quy tắc mới,
        để không bao giờ lệch với số hiện ở tab TỔNG HỢP CẤP XĂNG/ĐỊNH MỨC
        XĂNG. Phạm vi "khóa vượt định mức"/"thiếu định mức" tính trên TẤT
        CẢ trạng thái khóa (kể cả đã kết thúc), vì mục đích là rà soát toàn
        bộ, không chỉ việc đang làm hôm nay."""
        all_records = cap_xang_store.load()["records"]
        orphans = [(r, self.orphan_reason(r)) for r in all_records]
        orphans = [(r, d) for r, d in orphans if d]

        known_pairs = self.known_khoa_xe_pairs()
        tong_lit_khong_gan_khoa = sum(
            r.get("so_luong") or 0 for r in all_records
            if (r.get("ma_khoa", ""), r.get("xe", "")) not in known_pairs
        )

        # dedup_by_khoa_xe TRƯỚC khi đếm/cộng - 1 xe dùng chung nhiều ca sẽ có
        # nhiều dòng cùng tong/da_cap/con_lai (xem fuel_rows_for_teacher), đếm
        # thẳng theo dòng sẽ đếm/cộng trùng chiếc xe đó.
        all_khoa_rows = self.dedup_by_khoa_xe(self.fuel_rows_for_teacher(None, statuses=None))
        vuot_rows = [r for r in all_khoa_rows if r["con_lai"] is not None and r["con_lai"] < 0]
        thieu_rows = [r for r in all_khoa_rows if r["tong"] is None]

        return {
            "so_phieu_tong": len(all_records),
            "tong_lit_tatca": sum(r.get("so_luong") or 0 for r in all_records),
            "so_phieu_mo_coi": len(orphans),
            "tong_lit_mo_coi": sum(r.get("so_luong") or 0 for r, _ in orphans),
            "tong_lit_khong_gan_khoa": tong_lit_khong_gan_khoa,
            "so_khoa_tong": len(all_khoa_rows),
            "so_khoa_vuot": len(vuot_rows),
            "tong_lit_vuot": sum(-r["con_lai"] for r in vuot_rows),
            "so_khoa_thieu_dinh_muc": len(thieu_rows),
        }

    @staticmethod
    def _assignments_text(course, limit=6):
        pairs = []
        for g in course.get("groups", []):
            p = f"{g.get('giao_vien', '')} - {g.get('xe', '')}"
            if p not in pairs:
                pairs.append(p)
        text = "; ".join(pairs[:limit])
        return text + (f"; ... ({len(pairs)} cặp)" if len(pairs) > limit else "")

    def suggest_teachers(self, name):
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
            goi_y = self.suggest_teachers(giao_vien)
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

    def _allowance_breakdown(self, statuses=None):
        """Lõi tính "Tổng định mức"/"Còn lại" theo TỪNG (mã khóa, xe), tính 1
        LẦN cho MỌI giáo viên - dùng chung cho cả fuel_rows_for_teacher (tab
        ĐỊNH MỨC XĂNG) và allowances_for_teacher (tab CẤP XĂNG), để 2 tab
        luôn khớp nhau và không tính 2 cách khác nhau.

        1 xe có thể được NHIỀU ca/giáo viên dùng chung trong CÙNG 1 khóa (vd
        ca 1 thầy A, ca 2 thầy B) - "Đã cấp" luôn tính theo CẢ XE (không
        phân biệt ca, xem _da_cap_by_khoa_xe), nên "Tổng định mức" của xe đó
        cũng phải CỘNG DỒN định mức từng ca lại (mỗi ca tính theo định mức
        riêng của giáo viên ca đó x số học viên ca đó) rồi mới so với "Đã
        cấp" - nếu chỉ tính theo 1 ca rồi so với "Đã cấp" của cả xe sẽ báo
        "vượt định mức" SAI (đã từng bị vậy - xem lịch sử sửa).

        -> dict (ma_khoa, xe) -> {"contribs": [{"giao_vien","ca","so_hoc_vien",
        "dinh_muc","matched_by"}, ...] (1 phần tử/ca dùng xe đó), "so_hoc_vien":
        tổng mọi ca, "tong": tổng định mức CỘNG DỒN (None nếu THIẾU định mức
        của bất kỳ ca nào dùng chung xe đó - không tính được tổng đầy đủ thì
        thà báo "chưa tính được" còn hơn báo 1 số thiếu gây hiểu lầm), "da_cap",
        "con_lai", "ma_khoa"/"lop"/"xe"/"tu_ngay"/"den_ngay" (của khóa)}."""
        da_cap_map = self._da_cap_by_khoa_xe()
        result = {}
        for c in self.courses:
            if statuses is not None and self.status_of(c) not in statuses:
                continue
            by_xe = {}
            for g in c.get("groups", []):
                xe = g.get("xe", "")
                giao_vien = g.get("giao_vien", "")
                dinh_muc, matched_by = self.lookup_dinh_muc(xe, giao_vien)
                by_xe.setdefault(xe, []).append({
                    "giao_vien": giao_vien, "ca": g.get("ca", ""),
                    "so_hoc_vien": g.get("so_hoc_vien", 0),
                    "dinh_muc": dinh_muc, "matched_by": matched_by,
                })
            for xe, contribs in by_xe.items():
                tong = (
                    None if any(x["dinh_muc"] is None for x in contribs)
                    else sum(x["dinh_muc"] * x["so_hoc_vien"] for x in contribs)
                )
                da_cap = da_cap_map.get((c["ma_khoa"], xe), 0.0)
                result[(c["ma_khoa"], xe)] = {
                    "ma_khoa": c["ma_khoa"], "lop": c.get("lop", ""), "xe": xe,
                    "tu_ngay": c.get("tu_ngay", ""), "den_ngay": c.get("den_ngay", ""),
                    "contribs": contribs,
                    "so_hoc_vien": sum(x["so_hoc_vien"] for x in contribs),
                    "tong": tong, "da_cap": da_cap,
                    "con_lai": tong - da_cap if tong is not None else None,
                }
        return result

    def allowances_for_teacher(self, name):
        """Định mức/đã cấp/còn lại cho TỪNG khóa+xe giáo viên `name` từng được
        phân công (mọi trạng thái - đúng các dòng tab CẤP XĂNG liệt kê) ->
        dict (ma_khoa, xe) -> xem _allowance_breakdown (đã CỘNG DỒN định mức
        mọi ca dùng chung xe đó, không chỉ riêng giáo viên `name`). Dùng để
        hiện cột "Còn lại" + cảnh báo vượt định mức NGAY LÚC NHẬP ở tab CẤP
        XĂNG - thay vì phải sang tab ĐỊNH MỨC XĂNG mới biết (lúc đó đã cấp
        xong rồi)."""
        name_norm = attach_files.normalize_name(name)
        breakdown = self._allowance_breakdown(statuses=None)
        result = {}
        for c in self.courses:
            for g in c.get("groups", []):
                if attach_files.normalize_name(g.get("giao_vien", "")) != name_norm:
                    continue
                key = (c["ma_khoa"], g.get("xe", ""))
                result.setdefault(key, breakdown[key])
        return result

    def phieu_for_khoa_xe(self, ma_khoa, xe):
        """TOÀN BỘ phiếu cấp xăng khớp CHÍNH XÁC (mã khóa, xe) - đúng các
        phiếu được cộng vào "Đã cấp" ở _da_cap_by_khoa_xe/fuel_rows_for_teacher
        - để hiện chi tiết từng lần cấp làm căn cứ cho cột "Còn lại" (xem tab
        ĐỊNH MỨC XĂNG). Sắp theo ngày cấp tăng dần (không đọc được ngày thì
        xếp cuối)."""
        recs = [
            r for r in cap_xang_store.load()["records"]
            if r.get("ma_khoa", "") == ma_khoa and r.get("xe", "") == xe
        ]
        recs.sort(key=lambda r: _parse_date(r.get("ngay_cap", "")) or datetime.max)
        return recs

    def fuel_rows_for_teacher(self, name=None, statuses=("Đang đào tạo",)):
        """Trả về list các khóa của giáo viên `name` - hoặc TẤT CẢ giáo viên
        nếu `name` rỗng/None - có trạng thái nằm trong `statuses` (mặc định
        chỉ khóa ĐANG DẠY; None = mọi trạng thái, để quyết toán cả khóa đã
        kết thúc) - mỗi dòng là 1 (khóa, ca) - "dinh_muc"/"so_hoc_vien" là
        của RIÊNG ca đó (để hiện đúng ai dạy bao nhiêu học viên), còn
        "tong"/"da_cap"/"con_lai" là CHUNG CHO CẢ XE (xem _allowance_breakdown
        - cộng dồn định mức mọi ca dùng chung xe đó) nên 2 dòng cùng khóa+xe
        (khác ca/giáo viên) sẽ hiện CÙNG 1 giá trị tong/da_cap/con_lai - đây
        KHÔNG phải trùng lặp, mà vì 2 ca đó dùng chung 1 "ví" định mức của
        xe. Muốn đếm/cộng theo SỐ XE (không đếm trùng) thì lọc qua
        dedup_by_khoa_xe() trước. "contribs" = chi tiết từng ca đóng góp vào
        tong, dùng để giải thích cho người dùng (xem format_allowance_tooltip)."""
        name_norm = attach_files.normalize_name(name) if name else None
        breakdown = self._allowance_breakdown(statuses)
        rows = []
        for c in self.courses:
            if statuses is not None and self.status_of(c) not in statuses:
                continue
            for g in c.get("groups", []):
                giao_vien = g.get("giao_vien", "")
                if name_norm is not None and attach_files.normalize_name(giao_vien) != name_norm:
                    continue
                xe = g.get("xe", "")
                info = breakdown[(c["ma_khoa"], xe)]
                dinh_muc, matched_by = self.lookup_dinh_muc(xe, giao_vien)
                rows.append({
                    "giao_vien": giao_vien,
                    "ma_khoa": c["ma_khoa"],
                    "lop": c.get("lop", ""),
                    "xe": xe,
                    "ca": g.get("ca", ""),
                    "so_hoc_vien": g.get("so_hoc_vien", 0),
                    "dinh_muc": dinh_muc,
                    "matched_by": matched_by,
                    "tong": info["tong"],
                    "da_cap": info["da_cap"],
                    "con_lai": info["con_lai"],
                    "contribs": info["contribs"],
                    "tu_ngay": c.get("tu_ngay", ""),
                    "den_ngay": c.get("den_ngay", ""),
                })
        return rows

    @staticmethod
    def dedup_by_khoa_xe(rows):
        """Loại dòng trùng (mã khóa, xe) trong kết quả fuel_rows_for_teacher -
        1 xe dùng chung nhiều ca sẽ có nhiều dòng CÙNG tong/da_cap/con_lai
        (xem fuel_rows_for_teacher) - phải lọc qua đây trước khi ĐẾM số
        khóa/xe hay CỘNG tổng, nếu không sẽ đếm/cộng trùng chiếc xe đó."""
        seen, out = set(), []
        for r in rows:
            key = (r["ma_khoa"], r["xe"])
            if key in seen:
                continue
            seen.add(key)
            out.append(r)
        return out

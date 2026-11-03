"""Đọc 2 file XML xuất từ hệ thống quản lý của Sở GTVT:

- DSGV_*.xml: danh sách TOÀN BỘ giáo viên (dùng để sinh mã QR cố định theo
  MaGV - mã chính thức, không bao giờ trùng lặp). File này còn có ảnh chân
  dung (base64) và số CMT/CCCD - CHỈ đọc tên + điện thoại, không bao giờ
  đụng tới các trường nhạy cảm đó.
- DSXTL_*.xml: danh sách xe tập lái (dùng để đối chiếu biển số, cảnh báo
  nếu biển số trong Excel không có trong danh sách chính thức).
"""
import xml.etree.ElementTree as ET

# Không bao giờ đưa 2 trường này vào "extra" (ảnh chân dung base64 rất nặng,
# CMT/CCCD là thông tin nhạy cảm) - mọi trường khác trong DSGV đều hiện ra.
_EXCLUDED_EXTRA_TAGS = {"AnhCD", "SoCMT"}


def _format_ddmmyyyy(s):
    s = (s or "").strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[0:2]}/{s[2:4]}/{s[4:8]}"
    return s


def parse_dsgv(path, active_only=True):
    """Trả về dict: ma_gv (str) -> {"ten", "dien_thoai", "ngay_sinh",
    "dia_chi"}. ngay_sinh/dia_chi chỉ dùng để in thẻ giáo viên cục bộ
    (teacher_card.py) - KHÔNG bao giờ đưa lên trang web public."""
    root = ET.parse(path).getroot()
    body = root.find("GIAOVIENS")
    if body is None:
        return {}
    out = {}
    for rec in body.findall("GIAOVIEN"):

        def g(tag):
            el = rec.find(tag)
            return (el.text or "").strip() if el is not None and el.text else ""

        if active_only and g("TrangThai") not in ("1", ""):
            continue
        ma_gv = g("MaGV")
        if not ma_gv:
            continue
        ho_ten_dem = g("HoTenDem")
        ten = g("TenGV")
        full_name = f"{ho_ten_dem} {ten}".strip()
        if not full_name:
            continue
        extra = {
            child.tag: (child.text or "").strip()
            for child in rec
            if child.tag not in _EXCLUDED_EXTRA_TAGS
        }
        out[ma_gv] = {
            "ten": full_name,
            "dien_thoai": g("DienThoai"),
            "ngay_sinh": _format_ddmmyyyy(g("NgaySinh")),
            "dia_chi": g("NoiCT"),
            "extra": extra,
        }
    return out


def parse_dsxtl(path, active_only=True):
    """Trả về dict: biển số (nguyên văn) -> {"nhan_hieu":..., "loai_xe":...,
    "extra": {...}}. "extra" là TOÀN BỘ các trường còn lại của xe (số khung,
    số động cơ, hạn giấy phép, v.v...) - không có trường nào nhạy cảm trong
    DSXTL (không phải thông tin cá nhân) nên không cần loại trừ gì."""
    root = ET.parse(path).getroot()
    body = root.find("XETAPS")
    if body is None:
        return {}
    out = {}
    for rec in body.findall("XETAP"):

        def g(tag):
            el = rec.find(tag)
            return (el.text or "").strip() if el is not None and el.text else ""

        if active_only and g("TRANGTHAI") not in ("1", ""):
            continue
        bien_so = g("BIENSOXE")
        if not bien_so:
            continue
        extra = {child.tag: (child.text or "").strip() for child in rec}
        out[bien_so] = {
            "nhan_hieu": g("NHANHIEU"),
            "loai_xe": g("LOAIXE"),
            "extra": extra,
        }
    return out

"""Tab TỔNG QUAN: mở app lên là thấy ngay cần xử lý gì, không phải tự đi
lọc/tick từng tab để biết - gộp 3 con số đang nằm rải rác:

- Phiếu không khớp phân công ("mồ côi" - xem CapXangModel.orphan_reason).
- Khóa đã cấp VƯỢT định mức (tính trên MỌI trạng thái khóa, kể cả đã kết
  thúc, để rà soát toàn bộ chứ không chỉ việc đang làm hôm nay).
- Khóa/xe chưa tra được định mức nhiên liệu.

Mỗi mục có nút "Xem chi tiết" nhảy thẳng sang đúng tab + áp sẵn bộ lọc liên
quan, đỡ phải tự tìm lại.

Phần "Đối chiếu tổng số lít" giải thích vì sao tổng lít ở tab TỔNG HỢP CẤP
XĂNG (cộng mọi phiếu) khác tổng ở tab ĐỊNH MỨC XĂNG (chỉ cộng phiếu khớp
khóa+xe vào "Đã cấp" của khóa đó) - để không hiểu nhầm là thiếu/sai dữ liệu
khi so 2 con số này với nhau."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from tab_dinh_muc_xang import STATUS_OPTIONS


def _fmt_so(value):
    return f"{value:g}" if value is not None else "—"


class _StatCard(QGroupBox):
    """1 khối số liệu: số lớn + mô tả + nút "Xem chi tiết" (ẩn nếu không
    truyền on_detail). `warn=True` tô số lớn màu đỏ (dùng cho mục cần xử
    lý); warn=False tô xanh trung tính (chỉ để biết, không phải lỗi)."""

    def __init__(self, title, on_detail=None, parent=None):
        super().__init__(title, parent)
        layout = QVBoxLayout(self)
        self.lbl_number = QLabel("—")
        self.lbl_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_number.setStyleSheet("font-size: 26pt; font-weight: 800;")
        layout.addWidget(self.lbl_number)
        self.lbl_desc = QLabel("")
        self.lbl_desc.setWordWrap(True)
        self.lbl_desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_desc, 1)
        if on_detail is not None:
            btn = QPushButton("Xem chi tiết →")
            btn.clicked.connect(on_detail)
            layout.addWidget(btn)

    def set_value(self, number_text, desc_text, warn):
        self.lbl_number.setText(number_text)
        self.lbl_number.setStyleSheet(
            f"font-size: 26pt; font-weight: 800; color: {'#b3261e' if warn else '#2659c4'};"
        )
        self.lbl_desc.setText(desc_text)


class TongQuanTab(QWidget):
    def __init__(self, model, tabs, tab_dinh_muc, tab_tong_hop, parent=None):
        super().__init__(parent)
        self.model = model
        self.tabs = tabs
        self.tab_dinh_muc = tab_dinh_muc
        self.tab_tong_hop = tab_tong_hop
        self._build_ui()

    def _build_ui(self):
        outer = QVBoxLayout(self)

        lbl_title = QLabel("Tổng quan - cần rà soát gì ở dữ liệu cấp xăng hiện tại")
        lbl_title.setStyleSheet("font-size: 13pt; font-weight: 700;")
        outer.addWidget(lbl_title)

        cards_row = QHBoxLayout()
        self.card_mo_coi = _StatCard("Phiếu không khớp phân công", self._goto_mo_coi)
        cards_row.addWidget(self.card_mo_coi)
        self.card_vuot = _StatCard("Khóa đã cấp vượt định mức", self._goto_vuot)
        cards_row.addWidget(self.card_vuot)
        self.card_thieu = _StatCard("Khóa/xe chưa có định mức", self._goto_thieu)
        cards_row.addWidget(self.card_thieu)
        outer.addLayout(cards_row)

        recon_box = QGroupBox("Đối chiếu tổng số lít giữa các tab")
        recon_layout = QVBoxLayout(recon_box)
        self.lbl_recon = QLabel("")
        self.lbl_recon.setWordWrap(True)
        self.lbl_recon.setTextFormat(Qt.TextFormat.RichText)
        recon_layout.addWidget(self.lbl_recon)
        outer.addWidget(recon_box)

        btn_row = QHBoxLayout()
        btn_refresh = QPushButton("Làm mới")
        btn_refresh.clicked.connect(self.refresh)
        btn_row.addWidget(btn_refresh)
        btn_row.addStretch(1)
        outer.addLayout(btn_row)

        outer.addStretch(1)

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh()

    def refresh(self):
        s = self.model.overview_stats()

        if s["so_phieu_mo_coi"]:
            self.card_mo_coi.set_value(
                str(s["so_phieu_mo_coi"]),
                f"phiếu ({_fmt_so(s['tong_lit_mo_coi'])} lít) không khớp giáo viên/mã khóa/xe "
                "đang phân công - cần rà lại (có thể không cần sửa nếu là khóa đã kết thúc từ trước).",
                warn=True,
            )
        else:
            self.card_mo_coi.set_value("0", "Mọi phiếu đều khớp đúng phân công hiện tại.", warn=False)

        if s["so_khoa_vuot"]:
            self.card_vuot.set_value(
                str(s["so_khoa_vuot"]),
                f"khóa (trên tổng {s['so_khoa_tong']}) đã cấp vượt tổng định mức, vượt tổng cộng "
                f"{_fmt_so(s['tong_lit_vuot'])} lít - tính trên mọi trạng thái khóa.",
                warn=True,
            )
        else:
            self.card_vuot.set_value("0", f"Không có khóa nào vượt định mức (trên tổng {s['so_khoa_tong']} khóa).", warn=False)

        if s["so_khoa_thieu_dinh_muc"]:
            self.card_thieu.set_value(
                str(s["so_khoa_thieu_dinh_muc"]),
                "khóa/xe chưa tra được định mức nhiên liệu - chưa tính được có vượt hay không, "
                "cần bổ sung vào file định mức.",
                warn=True,
            )
        else:
            self.card_thieu.set_value("0", "Mọi khóa đều đã có định mức để đối chiếu.", warn=False)

        khong_gan = s["tong_lit_khong_gan_khoa"]
        da_gan = s["tong_lit_tatca"] - khong_gan
        self.lbl_recon.setText(
            f"Tổng cộng toàn bộ phiếu cấp xăng đã ghi: <b>{_fmt_so(s['tong_lit_tatca'])} lít</b> "
            f"({s['so_phieu_tong']} phiếu) - đây là số hiện ở tab TỔNG HỢP CẤP XĂNG khi bỏ hết bộ lọc.<br>"
            f"Trong đó <b>{_fmt_so(da_gan)} lít</b> khớp đúng khóa+xe đang có, được cộng vào \"Đã cấp\" "
            f"của khóa tương ứng ở tab ĐỊNH MỨC XĂNG; còn <b>{_fmt_so(khong_gan)} lít</b> không khớp khóa+xe "
            "nào nên KHÔNG được tính vào \"Đã cấp\" của khóa nào cả (thường là phiếu của khóa cũ đã kết thúc "
            "trước khi dùng app).<br>"
            "<i>Vì vậy tổng lít ở 2 tab đó không bằng nhau - không phải do thiếu/mất dữ liệu.</i>"
        )

    # ---------------- điều hướng sang tab khác kèm bộ lọc ----------------
    def _goto_mo_coi(self):
        self.tab_tong_hop.cbo_match.setCurrentIndex(2)  # "Chỉ phiếu KHÔNG khớp (mồ côi)"
        self.tabs.setCurrentWidget(self.tab_tong_hop)

    def _goto_vuot(self):
        self.tab_dinh_muc.cbo_status.setCurrentText(STATUS_OPTIONS[-1])  # "Tất cả trạng thái"
        self.tab_dinh_muc.chk_vuot.setChecked(True)
        self.tabs.setCurrentWidget(self.tab_dinh_muc)

    def _goto_thieu(self):
        self.tab_dinh_muc.cbo_status.setCurrentText(STATUS_OPTIONS[-1])
        self.tab_dinh_muc.chk_vuot.setChecked(False)
        self.tabs.setCurrentWidget(self.tab_dinh_muc)

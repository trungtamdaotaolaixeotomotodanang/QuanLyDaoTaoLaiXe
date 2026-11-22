"""Tab CẤP XĂNG: chọn ĐÚNG 1 bộ (giáo viên, khóa học, biển số xe) qua 3
combobox LIÊN KẾT nhau - chọn 1 cái thì 2 cái còn lại (nếu đang để trống) tự
lọc theo lựa chọn đó (xem _on_selector_changed) - rồi nhập số lít cấp. 1 ngày
có thể cấp nhiều lần: bấm "Cấp xăng" nhiều lần trong ngày, hoặc gõ cú pháp
gộp "SỐ PHIẾUxSỐ LÍT, ..." trong 1 lần (xem parse_so_luong/_GHI_CHU_SO_LUONG
- cú pháp này còn dùng lại ở tab TỔNG HỢP CẤP XĂNG, không được đổi chữ ký).

Hiện "Định mức / Đã cấp / Còn lại" của đúng khóa+xe đang chọn, cho SỬA TAY
"Tổng định mức" nếu số tính từ file Excel định mức không đúng thực tế (xem
CapXangModel.set_dinh_muc_override/clear_dinh_muc_override) - định mức Excel
chỉ mang tính tham khảo.

"Lịch sử các lần cấp trước" chỉ hiện đúng phiếu của bộ (giáo viên, khóa, xe)
đang chọn (không phải mọi khóa/xe của giáo viên đó) - bấm 1 dòng để nạp lại
vào khối nhập phía trên mà sửa/xóa.

Thay thế bản cũ (4 chế độ xem: Chi tiết + Tổng theo khóa/tuần/tháng dạng ma
trận ngày) - gộp lại còn đúng 1 luồng nhập duy nhất cho đơn giản."""
import re
from datetime import date, datetime

from PySide6.QtCore import QDate, QRegularExpression, Qt, QTimer
from PySide6.QtGui import QColor, QRegularExpressionValidator
from PySide6.QtWidgets import (
    QComboBox,
    QCompleter,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import attach_files
import cap_xang_store
from cap_xang_model import contrib_sub_tong, format_allowance_tooltip
from excel_export import GHI_CHU_TO_DO, ReportWorkbook
from style import icon as _icon

HISTORY_COLUMNS = ["Ngày cấp", "Biển số", "Số phiếu/Quyển", "Mặt hàng", "Số lượng", "Phân loại sổ", "Ghi chú"]

_SO_LUONG_KY_TU_RE = QRegularExpression(r"^[0-9xX, ]*$")
_SO_LUONG_PHAN_RE = re.compile(r"^(\d+)(?:\s*[xX]\s*(\d+))?$")
_SO_LUONG_MOI_RE = QRegularExpression(r"^\d*[.,]?\d*$")

_GHI_CHU_SO_LUONG = (
    "Nhập 1 số (vd 50) hoặc SỐ PHIẾUxSỐ LÍT (vd 2x30 = 2 phiếu, mỗi phiếu 30 "
    "lít). Nhiều phiếu cùng ngày cách nhau bằng dấu phẩy.\n"
    "Vd: 2x30, 4x20, 20, 50  =  2 phiếu 30 lít + 4 phiếu 20 lít + 1 phiếu 20 "
    "lít + 1 phiếu 50 lít (8 phiếu, tổng 210 lít)."
)

_MAU_CHU_VUOT = "#b3261e"  # "Còn lại" âm - đã cấp vượt định mức
_MAU_CHU_SAP_HET = "#b26a00"  # "Còn lại" còn dưới 10% định mức
_MAU_KHOA_HET_HAN = QColor("#b3261e")  # chữ đỏ: khóa "Đã đào tạo" (combobox Khóa học) / xe không chính thức phân công (combobox Biển số xe)


def _fmt_so(value):
    if value is None:
        return "—"
    return f"{value:g}"


def _parse_ngay(text):
    try:
        return datetime.strptime((text or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _vuot(allowance):
    """Khóa/xe đã cấp vượt định mức (Còn lại âm)? False nếu chưa có định mức."""
    return bool(allowance) and allowance["con_lai"] is not None and allowance["con_lai"] < 0


def parse_so_luong(text):
    """Phân tích chuỗi số lượng kiểu "2x30, 4x20, 20, 50" (xem _GHI_CHU_SO_LUONG
    cho cú pháp) -> trả về (tổng_số_lít, tổng_số_phiếu). Chuỗi rỗng -> (0, 0).
    Có phần sai định dạng -> (None, None). ĐƯỢC tab_tong_hop_cap_xang.py
    import lại - giữ nguyên chữ ký nếu sửa gì ở đây."""
    text = (text or "").strip()
    if not text:
        return 0.0, 0
    tong_lit = 0.0
    tong_phieu = 0
    for phan in text.split(","):
        phan = phan.strip()
        if not phan:
            continue
        m = _SO_LUONG_PHAN_RE.match(phan)
        if not m:
            return None, None
        if m.group(2) is not None:
            so_phieu, lit = int(m.group(1)), int(m.group(2))
        else:
            so_phieu, lit = 1, int(m.group(1))
        tong_lit += so_phieu * lit
        tong_phieu += so_phieu
    return tong_lit, tong_phieu


class _EditableCombo(QComboBox):
    """Combobox gõ tự do + gợi ý từ các giá trị đã nhập trước đó (vd "Mặt
    hàng", "Phân loại sổ") - không ép phải chọn trong danh sách."""

    def __init__(self, suggestions, text, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.set_suggestions(suggestions)
        self.setCurrentText(text)

    def set_suggestions(self, suggestions):
        current = self.currentText()
        self.blockSignals(True)
        self.clear()
        self.addItems(suggestions)
        self.setCurrentText(current)
        completer = QCompleter(suggestions, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.setCompleter(completer)
        self.blockSignals(False)


class _StatChip(QWidget):
    """1 ô nhỏ nổi bật 1 số liệu (vd "426 lít" + nhãn "Tổng định mức") - dùng
    cho khối Định mức thay vì 1 khối chữ dài, để liếc qua là thấy ngay 3 số
    quan trọng (Tổng/Đã cấp/Còn lại) mà không cần đọc câu chữ."""

    def __init__(self, caption, parent=None):
        super().__init__(parent)
        self.setObjectName("statChip")
        v = QVBoxLayout(self)
        v.setContentsMargins(12, 8, 12, 8)
        v.setSpacing(2)
        self.lbl_value = QLabel("—")
        self.lbl_caption = QLabel(caption)
        self.lbl_caption.setStyleSheet("color: #6c757d; font-weight: 600; font-size: 9pt; background: transparent;")
        v.addWidget(self.lbl_value)
        v.addWidget(self.lbl_caption)
        self.setStyleSheet("#statChip { background: #f7f8fa; border: 1px solid #e1e4ea; border-radius: 8px; }")
        self.set_value("—")

    def set_value(self, text, color="#1f2328"):
        self.lbl_value.setText(text)
        self.lbl_value.setStyleSheet(
            f"font-size: 15pt; font-weight: 700; color: {color}; background: transparent;"
        )


class CapXangTab(QWidget):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._all_assignments = []
        self._current_assignment = None  # dict giao_vien/ma_khoa/lop/xe/tu_ngay/den_ngay/status
        self._xe_override = False  # True nếu xe đang chọn KHÁC xe chính thức phân công (xem _resolve_and_reload_detail)
        self._current_allowance = None
        self._history_records = []
        self._editing_record_id = None
        self._updating_selectors = False
        self._last_teacher_text = ""
        self._build_ui()
        self._load_all_assignments()
        self._on_selector_changed()

    def refresh(self):
        """Nạp lại danh sách khóa/giáo viên/xe sau khi tab CẬP NHẬT DỮ LIỆU xử
        lý xong dữ liệu mới - xem main.py."""
        self._load_all_assignments()
        self._on_selector_changed()

    def _load_all_assignments(self):
        self._all_assignments = self.model.all_assignments()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.addWidget(self._build_selector_box())

        # QStackedWidget thay vì ẩn/hiện từng widget riêng lẻ (box_dinh_muc/
        # box_entry/tbl_history/lbl_history_total) - cách cũ để lại 1 mình
        # selector box với layout Preferred-size, không có gì nhận stretch,
        # nên Qt rải khoảng trống thừa vào CHÍNH selector box (và cả vài
        # khoảng trống rải rác khác) làm màn hình rỗng trông vỡ layout. Dồn
        # 2 trạng thái (chưa chọn đủ / đã chọn đủ) vào 2 "trang" của 1
        # QStackedWidget nhận stretch=1 thì trang nào đang hiện cũng tự
        # chiếm đúng phần không gian còn lại, không bị lệch.
        self.stack = QStackedWidget()
        self.stack.addWidget(self._build_placeholder_page())
        self.stack.addWidget(self._build_detail_page())
        outer.addWidget(self.stack, 1)

    def _build_placeholder_page(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.addStretch(1)
        icon_lbl = QLabel("⛽")
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_lbl.setStyleSheet("font-size: 42pt; background: transparent;")
        v.addWidget(icon_lbl)
        self.lbl_placeholder = QLabel("")
        self.lbl_placeholder.setWordWrap(True)
        self.lbl_placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_placeholder.setStyleSheet("font-size: 12pt; color: #5b6372; font-weight: 600; background: transparent;")
        v.addWidget(self.lbl_placeholder)
        v.addStretch(1)
        return page

    def _build_detail_page(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)

        self.lbl_xe_canh_bao = QLabel("")
        self.lbl_xe_canh_bao.setWordWrap(True)
        self.lbl_xe_canh_bao.setStyleSheet(
            "background: #fff1dc; border: 1px solid #f0c48a; border-radius: 4px; "
            "padding: 6px; color: #5a3b00; font-weight: 600;"
        )
        self.lbl_xe_canh_bao.hide()
        v.addWidget(self.lbl_xe_canh_bao)

        self.box_dinh_muc = self._build_dinh_muc_box()
        v.addWidget(self.box_dinh_muc)

        self.box_entry = self._build_entry_box()
        v.addWidget(self.box_entry)

        history_header = QHBoxLayout()
        history_header.addWidget(QLabel("Lịch sử các lần cấp trước:"))
        history_header.addStretch(1)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel...")
        self.btn_export.clicked.connect(self._on_export)
        history_header.addWidget(self.btn_export)
        v.addLayout(history_header)

        self.tbl_history = QTableWidget()
        self.tbl_history.setColumnCount(len(HISTORY_COLUMNS))
        self.tbl_history.setHorizontalHeaderLabels(HISTORY_COLUMNS)
        self.tbl_history.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_history.setAlternatingRowColors(True)
        self.tbl_history.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_history.cellClicked.connect(self._on_history_row_clicked)
        v.addWidget(self.tbl_history, 1)

        self.lbl_history_total = QLabel("")
        self.lbl_history_total.setStyleSheet("font-weight: 700;")
        v.addWidget(self.lbl_history_total)
        return page

    @staticmethod
    def _make_searchable_combo():
        combo = QComboBox()
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        return combo

    def _build_selector_box(self):
        box = QGroupBox("Chọn giáo viên / khóa học")
        row = QHBoxLayout(box)

        row.addWidget(QLabel("Giáo viên:"))
        self.cbo_teacher = self._make_searchable_combo()
        self.cbo_teacher.lineEdit().setPlaceholderText("Nhập hoặc chọn...")
        self.cbo_teacher.activated.connect(lambda _i: self._on_selector_changed("giao_vien"))
        self.cbo_teacher.lineEdit().editingFinished.connect(lambda: self._on_selector_changed("giao_vien"))
        row.addWidget(self.cbo_teacher, 1)

        row.addWidget(QLabel("Khóa học:"))
        # editable để gõ tìm (giống Giáo viên/Biển số xe) - khác 2 ô kia ở
        # chỗ giá trị thật đọc qua currentData() (mã khóa), không phải chữ
        # hiển thị - xem _on_khoa_editing_finished (chặn gõ sai lệch dữ liệu).
        self.cbo_khoa = self._make_searchable_combo()
        self.cbo_khoa.lineEdit().setPlaceholderText("Nhập hoặc chọn...")
        self.cbo_khoa.currentIndexChanged.connect(lambda _i: self._on_selector_changed("khoa"))
        self.cbo_khoa.lineEdit().editingFinished.connect(self._on_khoa_editing_finished)
        row.addWidget(self.cbo_khoa, 1)

        # Biển số xe ĐẶT Ở Ô "Cấp xăng" (_build_entry_box), ngay cạnh Ngày
        # cấp - vẫn là `self.cbo_xe`, chỉ đổi VỊ TRÍ hiển thị, không đổi gì
        # về logic 3 ô liên kết (_settle_selectors_once vẫn đọc/ghi đúng
        # combobox này như trước).

        return box

    def _build_dinh_muc_box(self):
        box = QGroupBox("Định mức")
        v = QVBoxLayout(box)

        chips = QHBoxLayout()
        self.chip_tong = _StatChip("Tổng định mức")
        self.chip_da_cap = _StatChip("Đã cấp")
        self.chip_con_lai = _StatChip("Còn lại")
        chips.addWidget(self.chip_tong, 1)
        chips.addWidget(self.chip_da_cap, 1)
        chips.addWidget(self.chip_con_lai, 1)
        v.addLayout(chips)

        self.lbl_dinh_muc_note = QLabel("")
        self.lbl_dinh_muc_note.setWordWrap(True)
        v.addWidget(self.lbl_dinh_muc_note)

        row = QHBoxLayout()
        self.btn_sua_dinh_muc = QPushButton("Sửa định mức...")
        self.btn_sua_dinh_muc.clicked.connect(self._on_sua_dinh_muc)
        row.addWidget(self.btn_sua_dinh_muc)
        self.btn_dinh_muc_goc = QPushButton("Dùng định mức gốc")
        self.btn_dinh_muc_goc.clicked.connect(self._on_dung_dinh_muc_goc)
        row.addWidget(self.btn_dinh_muc_goc)
        row.addStretch(1)
        v.addLayout(row)
        return box

    def _build_entry_box(self):
        box = QGroupBox("Cấp xăng")
        v = QVBoxLayout(box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Ngày cấp:"))
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        self.date_edit.setDate(QDate.currentDate())
        row1.addWidget(self.date_edit)
        row1.addSpacing(16)

        row1.addWidget(QLabel("Biển số xe:"))
        self.cbo_xe = self._make_searchable_combo()
        self.cbo_xe.lineEdit().setPlaceholderText("Nhập hoặc chọn...")
        self.cbo_xe.activated.connect(lambda _i: self._on_selector_changed("xe"))
        self.cbo_xe.lineEdit().editingFinished.connect(lambda: self._on_selector_changed("xe"))
        row1.addWidget(self.cbo_xe, 1)
        row1.addSpacing(16)

        row1.addWidget(QLabel("Số lượng:"))
        self.txt_so_luong = QLineEdit()
        self.txt_so_luong.setValidator(QRegularExpressionValidator(_SO_LUONG_KY_TU_RE))
        self.txt_so_luong.setPlaceholderText("vd: 2x30, 4x20, 20, 50")
        self.txt_so_luong.setToolTip(_GHI_CHU_SO_LUONG)
        self.txt_so_luong.returnPressed.connect(self._on_submit)
        row1.addWidget(self.txt_so_luong, 1)
        v.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Mặt hàng:"))
        self.cbo_mat_hang = _EditableCombo([], "")
        row2.addWidget(self.cbo_mat_hang, 1)
        row2.addWidget(QLabel("Phân loại sổ:"))
        self.cbo_phan_loai = _EditableCombo([], "")
        row2.addWidget(self.cbo_phan_loai, 1)
        row2.addWidget(QLabel("Số phiếu/Quyển:"))
        self.txt_so_phieu = QLineEdit()
        row2.addWidget(self.txt_so_phieu, 1)
        row2.addWidget(QLabel("Ghi chú:"))
        self.txt_ghi_chu = QLineEdit()
        row2.addWidget(self.txt_ghi_chu, 1)
        v.addLayout(row2)

        row3 = QHBoxLayout()
        self.btn_submit = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogApplyButton), "Cấp xăng")
        self.btn_submit.setObjectName("primary")
        self.btn_submit.clicked.connect(self._on_submit)
        row3.addWidget(self.btn_submit)
        self.btn_cancel_edit = QPushButton("Hủy sửa")
        self.btn_cancel_edit.clicked.connect(self._cancel_edit)
        self.btn_cancel_edit.setVisible(False)
        row3.addWidget(self.btn_cancel_edit)
        self.btn_delete_record = QPushButton("Xóa phiếu")
        self.btn_delete_record.clicked.connect(self._on_delete_current)
        self.btn_delete_record.setVisible(False)
        row3.addWidget(self.btn_delete_record)
        self.lbl_submit_status = QLabel("")
        self.lbl_submit_status.setStyleSheet("color: #1a7f37; font-weight: 600;")
        row3.addWidget(self.lbl_submit_status)
        row3.addStretch(1)
        v.addLayout(row3)
        return box

    # ---------------- 3 combobox liên kết ----------------
    def _on_selector_changed(self, source=None):
        if self._updating_selectors:
            return
        self._updating_selectors = True
        try:
            if source == "giao_vien":
                # Chọn giáo viên MỚI -> RESET Khóa + Xe (khác 2 ô kia, vốn chỉ
                # cập nhật 2 cái còn lại khi đang bỏ trống) - giáo viên là lựa
                # chọn gốc, đổi giáo viên thì lựa chọn khóa/xe cũ (của giáo
                # viên khác) không còn ý nghĩa gì nữa.
                #
                # CHỈ reset khi tên THỰC SỰ đổi (so với _last_teacher_text) -
                # editingFinished phát ra MỖI LẦN ô này mất focus, kể cả khi
                # chữ không đổi (vd tab/click sang ô Số lượng sau khi đã chọn
                # xong cả 3 ô) - nếu cứ reset vô điều kiện thì chỉ cần rời
                # khỏi ô Giáo viên là khóa/xe đã chọn bị xóa sạch, kể cả lúc
                # app tự đưa focus sang ô Số lượng sau khi vừa chọn xong (xem
                # _cancel_edit).
                gv_text = self.cbo_teacher.currentText().strip()
                if gv_text != self._last_teacher_text:
                    self._set_combo_blank(self.cbo_khoa)
                    self._set_combo_blank(self.cbo_xe)

            # Lặp tối đa vài vòng: mỗi vòng có thể tự chọn luôn 1 combobox
            # (xem _repopulate_*_combo - chỉ có đúng 1 lựa chọn thì tự chọn),
            # việc đó lại làm hẹp thêm lựa chọn của combobox còn lại - lặp
            # tới khi không còn gì đổi thêm (3 trường nên tối đa 3 vòng).
            for _ in range(3):
                if not self._settle_selectors_once(source):
                    break
        finally:
            self._updating_selectors = False

        self._last_teacher_text = self.cbo_teacher.currentText().strip()
        self._resolve_and_reload_detail()

    def _settle_selectors_once(self, source):
        """1 vòng cập nhật: lọc theo CẢ 3 ô đang có giá trị, rồi với mỗi ô
        CÒN LẠI (khác `source`) đang BỎ TRỐNG - nạp lại danh sách từ kết quả
        lọc đó; chỉ có đúng 1 lựa chọn thì TỰ CHỌN LUÔN. Trả về True nếu có
        ô nào vừa được tự chọn (cần gọi lại để lọc hẹp thêm ô còn lại)."""
        gv = self.cbo_teacher.currentText().strip() or None
        khoa_ma = self.cbo_khoa.currentData()
        xe = self.cbo_xe.currentText().strip() or None

        gv_norm = attach_files.normalize_name(gv) if gv else None
        xe_norm = attach_files.normalize_plate(xe) if xe else None

        matches = [
            a for a in self._all_assignments
            if (not gv_norm or attach_files.normalize_name(a["giao_vien"]) == gv_norm)
            and (not khoa_ma or a["ma_khoa"] == khoa_ma)
            and (not xe_norm or attach_files.normalize_plate(a["xe"]) == xe_norm)
        ]

        changed = False
        if source != "giao_vien" and not gv and self._repopulate_teacher_combo(matches):
            changed = True
        if source != "khoa" and not khoa_ma and self._repopulate_khoa_combo(matches):
            changed = True
        if source != "xe" and not xe and self._repopulate_xe_combo(matches, gv_khoa_set=bool(gv and khoa_ma)):
            changed = True
        return changed

    @staticmethod
    def _set_combo_blank(combo):
        """Bỏ chọn hẳn (currentText() rỗng) - dùng -1, KHÔNG dùng 0: ô Biển
        số xe lúc đã tự chọn sẵn 1 xe (_repopulate_xe_combo, gv_khoa_set)
        không hề thêm mục rỗng ở đầu danh sách (tu_chon=True), nên mục 0 lúc
        đó là 1 biển số THẬT - setCurrentIndex(0) tưởng là "bỏ trống" nhưng
        lại giữ nguyên biển số xe của lượt chọn giáo viên TRƯỚC ĐÓ, khiến
        giáo viên mới chọn bị lọc nhầm theo xe cũ (tưởng không có khóa nào)."""
        combo.blockSignals(True)
        combo.setCurrentIndex(-1)
        combo.blockSignals(False)

    def _repopulate_teacher_combo(self, matches):
        """Nạp lại danh sách giáo viên từ `matches` - TỰ CHỌN LUÔN nếu chỉ
        còn đúng 1 giáo viên khớp. Trả về True nếu vừa tự chọn."""
        names = sorted({a["giao_vien"] for a in matches if a["giao_vien"]}, key=attach_files.given_name_sort_key)
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.clear()
        self.cbo_teacher.addItem("")
        self.cbo_teacher.addItems(names)
        tu_chon = len(names) == 1
        self.cbo_teacher.setCurrentIndex(1 if tu_chon else 0)
        self._set_completer(self.cbo_teacher, names)
        self.cbo_teacher.blockSignals(False)
        return tu_chon

    def _repopulate_khoa_combo(self, matches):
        """Nạp lại danh sách khóa từ `matches` - TỰ CHỌN LUÔN nếu chỉ còn
        đúng 1 khóa khớp. Trả về True nếu vừa tự chọn."""
        by_khoa = {}
        for a in matches:
            by_khoa.setdefault(a["ma_khoa"], a)
        items = sorted(by_khoa.values(), key=lambda a: (a["status"] == "Đã đào tạo", a["ma_khoa"]))

        self.cbo_khoa.blockSignals(True)
        self.cbo_khoa.clear()
        self.cbo_khoa.addItem("", None)
        labels = []
        for a in items:
            label = f"{a['ma_khoa']} ({a['lop']})"
            if a.get("tu_ngay") and a.get("den_ngay"):
                label += f" [{a['tu_ngay']} → {a['den_ngay']}]"
            da_ket_thuc = a["status"] == "Đã đào tạo"
            if da_ket_thuc:
                label += " - ĐÃ KẾT THÚC"
            self.cbo_khoa.addItem(label, a["ma_khoa"])
            labels.append(label)
            if da_ket_thuc:
                self.cbo_khoa.setItemData(self.cbo_khoa.count() - 1, _MAU_KHOA_HET_HAN, Qt.ForegroundRole)
        tu_chon = len(items) == 1
        self.cbo_khoa.setCurrentIndex(1 if tu_chon else 0)
        self._set_completer(self.cbo_khoa, labels)
        self.cbo_khoa.blockSignals(False)
        return tu_chon

    def _on_khoa_editing_finished(self):
        """Khóa học đọc giá trị qua currentData() (mã khóa thật), không qua
        currentText() (chỉ là nhãn hiển thị) - khác Giáo viên/Biển số xe. Cho
        gõ để tìm (completer) nhưng gõ xong mà KHÔNG khớp đúng mục nào đang
        chọn thì phải trả chữ hiển thị về đúng mục đang chọn (currentIndex
        không đổi) - nếu không, chữ hiển thị sẽ lệch khỏi mã khóa thật đang
        dùng (currentData), gây nhầm lẫn/chọn sai khóa."""
        idx = self.cbo_khoa.currentIndex()
        dung_text = self.cbo_khoa.itemText(idx) if idx >= 0 else ""
        if self.cbo_khoa.lineEdit().text() != dung_text:
            self.cbo_khoa.lineEdit().setText(dung_text)

    def _repopulate_xe_combo(self, matches, gv_khoa_set):
        """Nạp lại danh sách biển số - LUÔN hiện TẤT CẢ biển số đang có trong
        dữ liệu (không chỉ xe của giáo viên/khóa đang chọn - `matches`) để
        giáo viên có thể CẤP CHO XE KHÁC khóa này không chính thức phân công
        (vd xe quen hỏng, mượn tạm xe khác) - xe KHỚP (`matches`, xe chính
        thức của giáo viên/khóa đang chọn) xếp LÊN ĐẦU danh sách, các xe còn
        lại tô đỏ xếp sau để phân biệt, nhưng vẫn chọn được (xem _resolve_
        and_reload_detail - cho phép "override" xe).

        `gv_khoa_set`: đã chọn xong CẢ giáo viên lẫn khóa (nên `matches` phản
        ánh đúng xe chính thức của riêng cặp đó) - chỉ khi đó mới MẶC ĐỊNH
        chọn sẵn xe đúng đầu tiên (không để trống) - còn khi giáo viên/khóa
        chưa chọn xong (matches chưa có nghĩa thật) thì vẫn để trống chờ
        chọn, không tự ý chọn đại 1 xe. Trả về True nếu vừa tự chọn."""
        plates_dung = sorted({a["xe"] for a in matches if a["xe"]})
        plates_dung_set = set(plates_dung)
        plates_khac = sorted({a["xe"] for a in self._all_assignments if a["xe"]} - plates_dung_set)

        self.cbo_xe.blockSignals(True)
        self.cbo_xe.clear()
        tu_chon = gv_khoa_set and bool(plates_dung)
        if not tu_chon:
            self.cbo_xe.addItem("")
        for p in plates_dung:
            self.cbo_xe.addItem(p)
        for p in plates_khac:
            self.cbo_xe.addItem(p)
            self.cbo_xe.setItemData(self.cbo_xe.count() - 1, _MAU_KHOA_HET_HAN, Qt.ForegroundRole)
        self.cbo_xe.setCurrentIndex(0)
        self._set_completer(self.cbo_xe, [*plates_dung, *plates_khac])
        self.cbo_xe.blockSignals(False)
        return tu_chon

    @staticmethod
    def _set_completer(combo, items):
        completer = QCompleter(items, combo)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        combo.setCompleter(completer)

    # ---------------- xác định bộ 3 đang chọn + nạp chi tiết ----------------
    def _resolve_and_reload_detail(self):
        gv = self.cbo_teacher.currentText().strip()
        khoa_ma = self.cbo_khoa.currentData()
        xe = self.cbo_xe.currentText().strip()

        if not gv or not khoa_ma or not xe:
            self._current_assignment = None
            self._show_placeholder("Chọn đủ Giáo viên / Khóa học / Biển số xe để cấp xăng.")
            return

        gv_norm = attach_files.normalize_name(gv)
        xe_norm = attach_files.normalize_plate(xe)
        exact = next(
            (a for a in self._all_assignments
             if attach_files.normalize_name(a["giao_vien"]) == gv_norm
             and a["ma_khoa"] == khoa_ma
             and attach_files.normalize_plate(a["xe"]) == xe_norm),
            None,
        )
        if exact is not None:
            self._current_assignment = exact
            self._xe_override = False
            self._show_detail_content()
            self._reload_allowance_and_history()
            return

        # Không khớp đúng xe nào trong phân công (giáo viên, khóa) này - vẫn
        # CHO PHÉP cấp cho XE KHÁC nếu giáo viên có dạy khóa đó (chỉ khác xe -
        # xem _repopulate_xe_combo, combobox Biển số xe hiện TẤT CẢ xe, tô đỏ
        # xe không chính thức) - vd xe quen hỏng, mượn tạm xe khác. Phiếu cấp
        # theo xe "lạ" này sẽ hiện "mồ côi" (không khớp phân công) ở tab TỔNG
        # HỢP CẤP XĂNG - đúng ý, để biết mà rà soát sau, không mất phiếu.
        base = next(
            (a for a in self._all_assignments
             if attach_files.normalize_name(a["giao_vien"]) == gv_norm and a["ma_khoa"] == khoa_ma),
            None,
        )
        if base is None:
            self._current_assignment = None
            self._xe_override = False
            self._show_placeholder(
                f"'{gv}' không được phân công dạy khóa '{khoa_ma}' - kiểm tra lại lựa chọn."
            )
            return

        self._current_assignment = {**base, "xe": xe}
        self._xe_override = True
        self._show_detail_content()
        self._reload_allowance_and_history()

    def _show_placeholder(self, text):
        self.lbl_placeholder.setText(text)
        self.stack.setCurrentIndex(0)
        self.btn_export.setEnabled(False)

    def _show_detail_content(self):
        self.stack.setCurrentIndex(1)

    def _reload_allowance_and_history(self):
        a = self._current_assignment
        if self._xe_override:
            self.lbl_xe_canh_bao.setText(
                f"⚠ Đang cấp cho xe '{a['xe']}' - KHÁC xe chính thức phân công cho "
                f"'{a['giao_vien']}' trong khóa {a['ma_khoa']} ({a['lop']}). Vẫn cho phép "
                "cấp, nhưng KHÔNG tính được định mức/còn lại cho xe này, và phiếu sẽ hiện "
                "ở mục \"không khớp phân công\" trong tab TỔNG HỢP CẤP XĂNG để rà soát sau."
            )
            self.lbl_xe_canh_bao.show()
        else:
            self.lbl_xe_canh_bao.hide()
        allowances = self.model.allowances_for_teacher(a["giao_vien"])
        self._current_allowance = allowances.get((a["ma_khoa"], a["xe"]))
        self._render_dinh_muc_box()
        self._cancel_edit()
        self._render_history()

    # ---------------- khối định mức ----------------
    def _render_dinh_muc_box(self):
        allowance = self._current_allowance
        if not allowance:
            self.chip_tong.set_value("—")
            self.chip_da_cap.set_value("—")
            self.chip_con_lai.set_value("—")
            self.lbl_dinh_muc_note.setText(
                "Chưa có định mức nào ghi nhận cho khóa/xe này - bấm \"Sửa định mức...\" để nhập tay."
            )
            self.box_dinh_muc.setToolTip("")
            self.btn_sua_dinh_muc.setEnabled(True)
            self.btn_dinh_muc_goc.setEnabled(False)
            return

        a = self._current_assignment
        contrib = self.model.contrib_for(allowance, a["giao_vien"]) or {}
        override = contrib.get("override")

        tong, da_cap, con_lai = allowance["tong"], allowance["da_cap"], allowance["con_lai"]

        color = ""
        if con_lai is not None and con_lai < 0:
            color = _MAU_CHU_VUOT
        elif tong and con_lai is not None and con_lai < 0.1 * tong:
            color = _MAU_CHU_SAP_HET

        self.chip_tong.set_value(f"{_fmt_so(tong)} lít" if tong is not None else "Chưa rõ")
        self.chip_da_cap.set_value(f"{_fmt_so(da_cap)} lít")
        self.chip_con_lai.set_value(
            f"{_fmt_so(con_lai)} lít" if con_lai is not None else "—", color or "#1f2328"
        )

        notes = []
        if override:
            goc_text = (
                f"{_fmt_so(contrib['dinh_muc'] * contrib['so_hoc_vien'])} lít"
                if contrib.get("dinh_muc") is not None else "chưa tính được"
            )
            note = f"Định mức phần của '{a['giao_vien']}' ĐÃ SỬA TAY (gốc theo Excel: {goc_text})"
            if override.get("ghi_chu"):
                note += f" - lý do: {override['ghi_chu']}"
            notes.append(note)
        elif tong is None:
            notes.append("Chưa tra được định mức (lít/học viên) từ file Excel.")
        if con_lai is not None and con_lai < 0:
            notes.append(f"⚠ ĐÃ CẤP VƯỢT ĐỊNH MỨC {_fmt_so(-con_lai)} lít.")
        self.lbl_dinh_muc_note.setText(" ".join(notes))
        self.lbl_dinh_muc_note.setStyleSheet(
            f"color: {_MAU_CHU_VUOT}; font-weight: 600;" if con_lai is not None and con_lai < 0
            else "color: #6c757d;"
        )
        self.box_dinh_muc.setToolTip(format_allowance_tooltip(allowance))
        self.btn_sua_dinh_muc.setEnabled(True)
        self.btn_dinh_muc_goc.setEnabled(bool(override))

    def _on_sua_dinh_muc(self):
        a = self._current_assignment
        if not a:
            return
        allowance = self._current_allowance or {}
        contrib = self.model.contrib_for(allowance, a["giao_vien"]) or {}
        override = contrib.get("override") or {}
        current_tong = contrib_sub_tong(contrib) if contrib else None

        dialog = QDialog(self)
        dialog.setWindowTitle("Sửa định mức")
        layout = QVBoxLayout(dialog)
        label_khoa = f"Khóa {a['ma_khoa']} ({a['lop']}) - xe {a['xe']}"
        if len(allowance.get("contribs") or []) > 1:
            label_khoa += f" - chỉ sửa phần của '{a['giao_vien']}' (xe này dùng chung nhiều giáo viên/ca)"
        layout.addWidget(QLabel(label_khoa))

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Tổng định mức mới (lít):"))
        edit_tong = QLineEdit(_fmt_so(current_tong) if current_tong is not None else "")
        edit_tong.setValidator(QRegularExpressionValidator(_SO_LUONG_MOI_RE))
        row1.addWidget(edit_tong)
        layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Lý do (tùy chọn):"))
        edit_ghi_chu = QLineEdit(override.get("ghi_chu", ""))
        row2.addWidget(edit_ghi_chu)
        layout.addLayout(row2)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        text = edit_tong.text().strip().replace(",", ".")
        try:
            tong_lit = float(text)
        except ValueError:
            QMessageBox.warning(self, "Số không hợp lệ", "Nhập 1 số lít hợp lệ.")
            return
        self.model.set_dinh_muc_override(a["ma_khoa"], a["xe"], a["giao_vien"], tong_lit, edit_ghi_chu.text())
        self._reload_allowance_and_history()

    def _on_dung_dinh_muc_goc(self):
        a = self._current_assignment
        if not a:
            return
        self.model.clear_dinh_muc_override(a["ma_khoa"], a["xe"], a["giao_vien"])
        self._reload_allowance_and_history()

    # ---------------- nhập/sửa phiếu ----------------
    def _confirm_vuot_dinh_muc(self, assignment, allowance, new_lit, old_lit):
        if not allowance or allowance["tong"] is None:
            return True
        sau_khi_cap = allowance["da_cap"] - old_lit + new_lit
        vuot = sau_khi_cap - allowance["tong"]
        if vuot <= 0:
            return True
        con_lai_truoc = allowance["tong"] - (allowance["da_cap"] - old_lit)
        answer = QMessageBox.question(
            self, "Vượt định mức xăng",
            f"Khóa {assignment['ma_khoa']} - xe {assignment['xe']} chỉ còn "
            f"{_fmt_so(con_lai_truoc)} lít (định mức {_fmt_so(allowance['tong'])} lít, "
            f"đã cấp {_fmt_so(allowance['da_cap'] - old_lit)} lít).\n"
            f"Phiếu này {_fmt_so(new_lit)} lít -> VƯỢT định mức {_fmt_so(vuot)} lít.\n\n"
            "Vẫn cấp và lưu phiếu này?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return answer == QMessageBox.StandardButton.Yes

    def _on_submit(self):
        a = self._current_assignment
        if not a:
            return
        so_luong_text = self.txt_so_luong.text().strip()
        if not so_luong_text:
            QMessageBox.information(self, "Thiếu số lượng", "Nhập số lít cần cấp.")
            return
        tong_lit, tong_phieu = parse_so_luong(so_luong_text)
        if tong_lit is None:
            QMessageBox.warning(self, "Số lượng không đúng định dạng", _GHI_CHU_SO_LUONG)
            return
        if tong_lit <= 0:
            QMessageBox.information(self, "Số lượng không hợp lệ", "Số lượng phải lớn hơn 0.")
            return

        ngay_cap = self.date_edit.date().toPython()
        tu = _parse_ngay(a.get("tu_ngay", ""))
        den = _parse_ngay(a.get("den_ngay", ""))
        if tu and ngay_cap < tu:
            QMessageBox.warning(
                self, "Ngày cấp không hợp lệ",
                f"Ngày cấp ({ngay_cap.strftime('%d/%m/%Y')}) trước ngày khóa "
                f"'{a['ma_khoa']}' bắt đầu học ({a['tu_ngay']}).",
            )
            return
        if den and ngay_cap > den:
            QMessageBox.warning(
                self, "Ngày cấp không hợp lệ",
                f"Ngày cấp ({ngay_cap.strftime('%d/%m/%Y')}) sau ngày khóa "
                f"'{a['ma_khoa']}' kết thúc học ({a['den_ngay']}).",
            )
            return

        old_lit = 0
        if self._editing_record_id:
            old_record = next((r for r in self._history_records if r["id"] == self._editing_record_id), None)
            old_lit = (old_record.get("so_luong") or 0) if old_record else 0

        if not self._confirm_vuot_dinh_muc(a, self._current_allowance, tong_lit, old_lit):
            return

        fields = {
            "giao_vien": a["giao_vien"], "ma_khoa": a["ma_khoa"], "lop": a["lop"], "xe": a["xe"],
            "ngay_cap": ngay_cap.strftime("%d/%m/%Y"),
            "so_phieu": self.txt_so_phieu.text().strip(),
            "mat_hang": self.cbo_mat_hang.currentText().strip(),
            "so_luong": tong_lit, "so_luong_text": so_luong_text, "so_phieu_count": tong_phieu,
            "ghi_chu": self.txt_ghi_chu.text().strip(),
            "phan_loai_so": self.cbo_phan_loai.currentText().strip(),
        }
        is_new = self._editing_record_id is None
        if self._editing_record_id:
            cap_xang_store.update_record(self._editing_record_id, fields)
        else:
            cap_xang_store.add_record(fields)
        self._reload_allowance_and_history()
        msg = (
            f"✓ Đã cấp {tong_phieu} phiếu ({_fmt_so(tong_lit)} lít) lúc {datetime.now().strftime('%H:%M')}"
            if is_new else "✓ Đã lưu thay đổi."
        )
        self._flash_submit_status(msg)

    def _flash_submit_status(self, text):
        """Hiện ngắn 1 dòng xác nhận đã lưu cạnh nút Cấp xăng rồi tự ẩn - để
        người dùng biết chắc đã lưu mà không cần hộp thoại chặn thao tác tiếp."""
        self.lbl_submit_status.setText(text)
        QTimer.singleShot(4000, lambda: self.lbl_submit_status.setText(""))

    def _on_history_row_clicked(self, row, _col):
        if row < 0 or row >= len(self._history_records):
            return
        record = self._history_records[row]
        self._select_history_record_xe(record)
        self._load_record_into_form(record)

    def _select_history_record_xe(self, record):
        """Lịch sử giờ liệt kê phiếu của MỌI xe (xem _render_history) - bấm
        vào 1 dòng dùng xe KHÁC xe đang chọn trên ô Biển số xe thì phải đổi
        luôn ô đó sang đúng xe của phiếu này rồi giải quyết lại (Định mức/
        cảnh báo override) - nếu không, lúc bấm "Lưu thay đổi" sẽ ghi NHẦM
        phiếu này sang xe đang chọn trước đó (xem _on_submit dùng xe của
        self._current_assignment)."""
        xe = record.get("xe", "")
        if attach_files.normalize_plate(xe) == attach_files.normalize_plate(self.cbo_xe.currentText()):
            return
        idx = self.cbo_xe.findText(xe)
        if idx >= 0:
            self.cbo_xe.setCurrentIndex(idx)
        else:
            # Biển số của phiếu cũ không còn trong danh sách hiện có (hiếm -
            # vd gõ sai/xe đã thanh lý) - vẫn gán thẳng chữ để sửa/xóa được.
            self.cbo_xe.blockSignals(True)
            self.cbo_xe.setEditText(xe)
            self.cbo_xe.blockSignals(False)
        self._on_selector_changed("xe")

    def _load_record_into_form(self, record):
        self._editing_record_id = record["id"]
        ngay = _parse_ngay(record.get("ngay_cap", ""))
        self.date_edit.setDate(QDate(ngay.year, ngay.month, ngay.day) if ngay else QDate.currentDate())
        self.txt_so_luong.setText(record.get("so_luong_text") or _fmt_so(record.get("so_luong")))
        self.cbo_mat_hang.setCurrentText(record.get("mat_hang", ""))
        self.cbo_phan_loai.setCurrentText(record.get("phan_loai_so", ""))
        self.txt_so_phieu.setText(record.get("so_phieu", ""))
        self.txt_ghi_chu.setText(record.get("ghi_chu", ""))
        self.btn_submit.setText("Lưu thay đổi")
        self.btn_cancel_edit.setVisible(True)
        self.btn_delete_record.setVisible(True)

    def _cancel_edit(self):
        self._editing_record_id = None
        self.date_edit.setDate(QDate.currentDate())
        self.txt_so_luong.clear()
        self.txt_so_phieu.clear()
        self.txt_ghi_chu.clear()
        mat_hang_goi_y = cap_xang_store.distinct_values("mat_hang")
        phan_loai_goi_y = cap_xang_store.distinct_values("phan_loai_so")
        self.cbo_mat_hang.set_suggestions(mat_hang_goi_y)
        self.cbo_mat_hang.setCurrentText(mat_hang_goi_y[0] if mat_hang_goi_y else "")
        self.cbo_phan_loai.set_suggestions(phan_loai_goi_y)
        self.cbo_phan_loai.setCurrentText(phan_loai_goi_y[0] if phan_loai_goi_y else "")
        self.btn_submit.setText("Cấp xăng")
        self.btn_cancel_edit.setVisible(False)
        self.btn_delete_record.setVisible(False)
        # Đưa con trỏ sẵn vào ô Số lượng - đúng ô cần gõ tiếp theo trong luồng
        # thao tác nhanh (chọn xong giáo viên/khóa/xe -> gõ số lượng -> Enter),
        # khỏi phải bấm chuột vào ô mỗi lần cấp. HOÃN LẠI (singleShot 0) chứ
        # không setFocus() ngay: _cancel_edit() này chạy TRONG LÚC
        # _on_selector_changed đang xử lý (vd người dùng vừa gõ xong tên giáo
        # viên, ô đó đang có focus) - setFocus() ngay lập tức làm ô giáo viên
        # mất focus, phát editingFinished NGAY, gọi đệ quy lại
        # _on_selector_changed trong khi vòng gọi ngoài chưa chạy xong, làm
        # self._current_assignment bị set về None giữa chừng -> crash ở
        # _render_history ngay sau đó (đã gặp lúc kiểm tra).
        QTimer.singleShot(0, self.txt_so_luong.setFocus)

    def _on_delete_current(self):
        if not self._editing_record_id:
            return
        confirm = QMessageBox.question(self, "Xác nhận", "Xóa phiếu cấp xăng này?")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        cap_xang_store.delete_record(self._editing_record_id)
        self._reload_allowance_and_history()

    # ---------------- lịch sử ----------------
    def _render_history(self):
        """Lịch sử hiện TOÀN BỘ phiếu của (giáo viên, khóa học) đang chọn -
        KHÔNG lọc theo biển số xe nữa (khác hành vi cũ) - để quản lý dễ theo
        dõi khi giáo viên tạm mượn xe khác (vd xe quen hỏng vài bữa): mọi
        phiếu của đúng giáo viên/khóa này đều nằm chung 1 danh sách, không bị
        tách rời theo từng xe. Cột "Biển số" cho biết từng phiếu dùng xe
        nào."""
        a = self._current_assignment
        gv_norm = attach_files.normalize_name(a["giao_vien"])
        all_records = cap_xang_store.load()["records"]
        records = [
            r for r in all_records
            if r.get("ma_khoa", "") == a["ma_khoa"]
            and attach_files.normalize_name(r.get("giao_vien", "")) == gv_norm
        ]
        records.sort(key=lambda r: _parse_ngay(r.get("ngay_cap", "")) or date.min)
        self._history_records = records

        self.tbl_history.setRowCount(len(records))
        for row_idx, r in enumerate(records):
            so_luong_display = r.get("so_luong_text") or _fmt_so(r.get("so_luong"))
            values = [
                r.get("ngay_cap", ""), r.get("xe", ""), r.get("so_phieu", ""), r.get("mat_hang", ""),
                so_luong_display, r.get("phan_loai_so", ""), r.get("ghi_chu", ""),
            ]
            for c, val in enumerate(values):
                item = QTableWidgetItem(val)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if c == 0:
                    item.setData(Qt.ItemDataRole.UserRole, r["id"])
                self.tbl_history.setItem(row_idx, c, item)

        if records:
            tong = sum(r.get("so_luong") or 0 for r in records)
            self.lbl_history_total.setText(f"Tổng: {_fmt_so(tong)} lít ({len(records)} dòng).")
        else:
            self.lbl_history_total.setText("Chưa có lần cấp nào cho giáo viên/khóa học này.")
        self.btn_export.setEnabled(bool(records))

    # ---------------- xuất Excel ----------------
    def _on_export(self):
        if not self._current_assignment or not self._history_records:
            QMessageBox.information(self, "Không có dữ liệu", "Chưa có phiếu nào để xuất cho lựa chọn hiện tại.")
            return
        a = self._current_assignment
        safe_name = re.sub(r"[^\w]+", "_", a["giao_vien"], flags=re.UNICODE).strip("_")
        default_name = f"CapXang_{safe_name}_{a['ma_khoa']}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất Excel", default_name, "Excel Files (*.xlsx)")
        if not path:
            return

        rows = []
        for r in self._history_records:
            so_luong_display = r.get("so_luong_text") or _fmt_so(r.get("so_luong"))
            rows.append([
                _parse_ngay(r.get("ngay_cap", "")) or r.get("ngay_cap", ""), r.get("xe", ""),
                r.get("so_phieu", ""), r.get("mat_hang", ""), so_luong_display,
                r.get("phan_loai_so", ""), r.get("ghi_chu", ""),
            ])
        vuot = _vuot(self._current_allowance)
        try:
            report = ReportWorkbook()
            report.add_sheet("Lich su cap xang", HISTORY_COLUMNS, rows, highlight=[vuot] * len(rows))
            report.save(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        QMessageBox.information(self, "Hoàn tất", f"Đã xuất {len(rows)} dòng ra:\n{path}\n\n{GHI_CHU_TO_DO}")

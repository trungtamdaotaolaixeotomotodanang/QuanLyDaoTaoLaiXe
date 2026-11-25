"""Tab TỔNG HỢP CẤP XĂNG: xem toàn bộ phiếu cấp xăng của CẢ TRUNG TÂM (mọi
giáo viên) trong 1 bảng, không giới hạn theo 1 giáo viên như tab CẤP XĂNG.

- Lọc theo Giáo viên / Khóa học / khoảng ngày cấp (tùy chọn, mặc định xem hết).
  Ô "Khóa học" chỉ hiện khóa của giáo viên đang lọc (nếu có chọn) và tô đỏ
  khóa đã kết thúc - giống combobox Khóa học ở tab CẤP XĂNG; vẫn gõ tự do
  được mã khóa cũ không còn trong danh sách (vd để lọc phiếu "mồ côi").
- "Nhóm theo": gom các phiếu theo Giáo viên/Khóa học/Biển số thành từng nhóm
  thu/mở được, mỗi nhóm có tổng số lít/số phiếu riêng.
- Bấm vào tên cột để sắp xếp theo cột đó (bấm lại đổi tăng/giảm) - giữ Shift
  bấm thêm cột khác để sắp theo nhiều cấp (vd Giáo viên trước, Ngày cấp sau).
  Kéo tiêu đề cột để đổi thứ tự hiển thị các cột.
- Chọn 1 hoặc nhiều dòng (Ctrl/Shift+click) rồi bấm "Xóa dòng đã chọn" (hoặc
  chuột phải -> Xóa) để xóa. Sau khi xóa có nút "Hoàn tác" (ẩn sau ít giây)
  để khôi phục lại đúng các phiếu vừa xóa nếu lỡ tay.
- SỬA TRỰC TIẾP: double-click vào bất kỳ ô nào (trừ dòng tiêu đề nhóm) để
  sửa, Enter/rời ô để lưu ngay - không cần quay lại tab CẤP XĂNG nữa. Hữu
  ích nhất cho các phiếu không khớp khóa/xe đang phân công hiện tại (nhập
  tay mã khóa cũ/tên viết tắt...) - trước đây các phiếu này chỉ xem/xóa
  được ở đây, không sửa được ở đâu cả.
  KHÔNG lọc theo biển số (giống tab CẤP XĂNG) - xe có thể mượn tạm nên lọc
  theo biển số dễ làm lạc mất phiếu khi rà soát; muốn xem theo xe thì dùng
  "Nhóm theo" -> Biển số. "Kỳ nhanh" chọn sẵn tháng này/tháng trước/quý
  này/năm nay cho báo cáo định kỳ, không phải bấm lịch 2 lần.
- PHIẾU KHÔNG KHỚP PHÂN CÔNG ("mồ côi" - xem CapXangModel.orphan_reason, vd
  phiếu ghi đè biển số xe khác xe chính thức): tô cam nhạt ngay trong bảng +
  rê chuột/bấm vào dòng để xem lý do + cách sửa (không có ô lọc riêng - tô
  cam đã đủ nổi bật để nhận ra khi rà soát, không cần ẩn bớt các phiếu khác).
- "Xuất Excel...": xuất theo đúng bộ lọc đang chọn, nhiều sheet (chi tiết +
  tổng theo giáo viên/khóa-xe/biển số/tháng + bảng giáo viên x tháng + định
  mức), mỗi sheet là bảng Excel có nút lọc/sort sẵn + 1 dòng chú thích ngay
  trên tiêu đề cho biết đang sắp xếp theo cột nào."""
from datetime import date, datetime, timedelta

from PySide6.QtCore import QDate, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QTextDocument
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMenu,
    QMessageBox,
    QPushButton,
    QStyle,
    QStyledItemDelegate,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

import attach_files
import cap_xang_store
from excel_export import GHI_CHU_TO_DO, ReportWorkbook
from style import icon as _icon
from tab_cap_xang import _MAU_KHOA_HET_HAN, parse_so_luong
from tab_dinh_muc_xang import EXPORT_COLUMNS as _DINH_MUC_COLUMNS, dinh_muc_export_row, is_vuot

_SO_LUONG_GHI_CHU = (
    "Gõ 1 số (vd 50) hoặc SỐ PHIẾUxSỐ LÍT, nhiều phiếu cách nhau bằng dấu "
    "phẩy (vd 2x30, 4x20) - xem hướng dẫn ở tab CẤP XĂNG."
)

_ALL_LABEL = "-- Tất cả --"
GROUP_OPTIONS = ["Không nhóm", "Giáo viên", "Khóa học", "Biển số"]
PERIOD_OPTIONS = ["Tùy chọn", "Tuần này", "Tháng này", "Tháng trước", "Quý này", "Năm nay"]
_MAU_MO_COI = QColor("#fff1dc")  # nền dòng phiếu không khớp phân công hiện tại

_COLUMN_LABELS = [
    "Giáo viên", "Khóa học", "Lớp", "Biển số", "Ngày cấp", "Số phiếu/Quyển",
    "Mặt hàng", "Số lượng", "Phân loại sổ", "Ghi chú",
]
_COLUMN_FIELDS = [
    "giao_vien", "ma_khoa", "lop", "xe", "ngay_cap", "so_phieu",
    "mat_hang", "so_luong", "phan_loai_so", "ghi_chu",
]

_HUONG_DAN_HTML = (
    "<p><b>Tab \"TỔNG HỢP CẤP XĂNG\"</b> - xem toàn bộ phiếu cấp xăng của CẢ "
    "TRUNG TÂM (mọi giáo viên), không giới hạn theo 1 giáo viên như tab CẤP "
    "XĂNG.</p>"
    "<ul>"
    "<li>Lọc theo <b>Giáo viên</b> / <b>Khóa học</b> (để \"-- Tất cả --\" để "
    "xem hết). Chọn giáo viên trước thì ô Khóa học chỉ còn hiện khóa của "
    "giáo viên đó, khóa đã kết thúc tô đỏ - vẫn gõ tự do được mã khóa cũ "
    "không còn trong danh sách nếu cần lọc phiếu \"mồ côi\".</li>"
    "<li>KHÔNG lọc theo <b>Biển số</b> (giống tab CẤP XĂNG) - xe có thể mượn "
    "tạm, lọc theo biển số dễ làm lạc mất phiếu khi rà soát. Muốn xem theo xe "
    "thì dùng <b>Nhóm theo</b> -> Biển số.</li>"
    "<li><b>Khoảng ngày</b>: bỏ tick \"Tất cả thời gian\" để lọc theo 1 "
    "khoảng ngày cấp cụ thể, hoặc chọn <b>Kỳ nhanh</b> (tháng này/tháng "
    "trước/quý này/năm nay) - app tự điền Từ ngày/Đến ngày.</li>"
    "<li><b>Phiếu KHÔNG khớp phân công</b> (khóa/xe/giáo viên trên phiếu "
    "không trùng phân công hiện có, vd phiếu ghi đè biển số xe khác xe chính "
    "thức) tự tô <b>cam nhạt</b> ngay trong bảng - không cần lọc riêng. "
    "<b>Bấm vào dòng</b> -> ô cam dưới bảng ghi rõ sai ở ô nào và cách sửa "
    "(gợi ý mã khóa/tên/biển số đúng nếu suy ra được). 3 loại: (1) <i>ô Khóa "
    "học trống</i> - phiếu nhập từ Excel cũ chỉ có mã lớp, khóa đã kết thúc "
    "không còn trong dữ liệu -> KHÔNG cần sửa, vẫn tính vào tổng; (2) <i>tên "
    "giáo viên lạ</i> -> sửa ô Giáo viên theo gợi ý; (3) <i>khóa có nhưng "
    "giáo viên/xe không đúng phân công</i> -> sửa ô Giáo viên hoặc Biển số "
    "theo gợi ý (nếu không phải chủ ý ghi đè xe). Phiếu không khớp không "
    "hiện ở lưới Chi tiết tab CẤP XĂNG và không trừ vào \"Còn lại\" của khóa "
    "nào ở tab ĐỊNH MỨC XĂNG.</li>"
    "<li><b>Xuất Excel...</b>: xuất theo đúng bộ lọc đang chọn ra file nhiều "
    "sheet (chi tiết từng phiếu, tổng theo giáo viên / khóa-xe / biển số / "
    "tháng, bảng giáo viên x tháng, định mức-đã cấp-còn lại) - mỗi sheet là "
    "bảng có nút lọc/sort sẵn khi mở trong Excel.</li>"
    "<li><b>Nhóm theo</b>: gom các phiếu theo Giáo viên/Khóa học/Biển số "
    "thành từng nhóm có tổng số lít/số phiếu riêng - bấm vào nhóm để thu/mở.</li>"
    "<li><b>Bấm vào tên cột</b> (vd \"Ngày cấp\", \"Số lượng\"...) để sắp xếp "
    "theo cột đó, bấm lại để đổi tăng/giảm. Giữ <b>Shift</b> rồi bấm thêm cột "
    "khác để sắp theo nhiều cấp (vd sắp theo Giáo viên trước, trong mỗi giáo "
    "viên sắp theo Ngày cấp).</li>"
    "<li>Có thể <b>kéo tiêu đề cột</b> sang trái/phải để đổi thứ tự hiển thị "
    "các cột theo ý muốn.</li>"
    "<li>Chọn 1 hoặc nhiều dòng (giữ <b>Ctrl/Shift</b> rồi click để chọn thêm) "
    "rồi bấm <b>\"Xóa dòng đã chọn\"</b> (hoặc chuột phải -> Xóa) để xóa. Sau "
    "khi xóa có nút <b>\"Hoàn tác\"</b> hiện ra ít giây để khôi phục lại nếu "
    "lỡ tay.</li>"
    "<li><b>Double-click vào 1 ô</b> (trừ dòng tiêu đề nhóm) để sửa trực "
    "tiếp, Enter/rời ô để lưu ngay - không cần quay lại tab CẤP XĂNG. Cột "
    "\"Ngày cấp\" phải đúng dạng dd/mm/yyyy, cột \"Số lượng\" gõ như hướng "
    "dẫn ở tab CẤP XĂNG (vd 2x30, 4x20) - gõ sai sẽ báo lỗi và giữ nguyên "
    "giá trị cũ.</li>"
    "</ul>"
)


class _TeacherEditDelegate(QStyledItemDelegate):
    """Double-click ô "Giáo viên" hiện combobox có search (giống tab CẤP
    XĂNG) thay vì gõ tự do hoàn toàn - vẫn cho gõ tên không có trong danh
    sách (vd đợi nhập mới), chỉ thêm gợi ý để đỡ gõ sai/viết tắt."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self._model = model

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.setEditable(True)
        combo.setInsertPolicy(QComboBox.NoInsert)
        teachers = self._model.all_teachers()
        combo.addItems(teachers)
        completer = QCompleter(teachers, combo)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        combo.setCompleter(completer)
        return combo

    def setEditorData(self, editor, index):
        editor.setEditText(index.data() or "")

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText().strip())


def _fmt_so(value):
    if value is None:
        return "—"
    return f"{value:g}"


def _parse_ngay(text):
    try:
        return datetime.strptime((text or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def _period_range(label, today):
    """(từ ngày, đến ngày) cho 1 lựa chọn "Kỳ nhanh" - kỳ đang diễn ra thì
    tính tới hôm nay."""
    if label == "Tuần này":
        return today - timedelta(days=today.weekday()), today
    if label == "Tháng này":
        return today.replace(day=1), today
    if label == "Tháng trước":
        end = today.replace(day=1) - timedelta(days=1)
        return end.replace(day=1), end
    if label == "Quý này":
        first_month = 3 * ((today.month - 1) // 3) + 1
        return today.replace(month=first_month, day=1), today
    if label == "Năm nay":
        return today.replace(month=1, day=1), today
    return today, today


def _month_key(r):
    d = _parse_ngay(r.get("ngay_cap", ""))
    return (d.year, d.month) if d else None


def _month_label(key):
    return f"{key[1]:02d}/{key[0]}" if key else "(không rõ ngày)"


def _aggregate(records, key_fn):
    """Gộp phiếu theo key_fn(r) -> {key: {"n": số dòng, "phieu": số phiếu,
    "lit": tổng lít}} (giữ thứ tự xuất hiện)."""
    agg = {}
    for r in records:
        a = agg.setdefault(key_fn(r), {"n": 0, "phieu": 0, "lit": 0.0})
        a["n"] += 1
        a["phieu"] += r.get("so_phieu_count") or 1
        a["lit"] += r.get("so_luong") or 0
    return agg


class TongHopCapXangTab(QWidget):
    # Phát ra sau khi XÓA/SỬA phiếu trực tiếp trên bảng ở tab này (lưu vào
    # shared_data cục bộ) - main.py nghe tín hiệu này để TỰ đẩy lên GitHub
    # ngay (xem _on_delete_selected/_on_undo_delete/_on_item_changed).
    du_lieu_da_doi = Signal()

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._sort_keys = [("ngay_cap", False)]  # mặc định: ngày cấp mới nhất trước
        self._orphan_reasons = {}
        self._build_ui()
        self._populate_filters()
        self._render()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)

        box = QGroupBox("Tổng hợp cấp xăng toàn trung tâm")
        form = QVBoxLayout(box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Giáo viên:"))
        self.cbo_teacher = QComboBox()
        self.cbo_teacher.setEditable(True)
        self.cbo_teacher.setInsertPolicy(QComboBox.NoInsert)
        # Lọc bảng ngay khi gõ (mỗi phím) nhưng CHỈ nạp lại danh sách Khóa
        # học (kèm completer riêng) khi đã CHỌN XONG giáo viên (activated/
        # rời ô) - dựng lại combobox + completer mỗi phím gõ sẽ giật/đóng
        # popup đang gõ, không như _render() (chỉ lọc bảng, nhẹ hơn nhiều).
        self.cbo_teacher.currentTextChanged.connect(lambda _: self._render())
        self.cbo_teacher.activated.connect(lambda _i: self._on_teacher_filter_settled())
        self.cbo_teacher.lineEdit().editingFinished.connect(self._on_teacher_filter_settled)
        row1.addWidget(self.cbo_teacher, 1)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Khóa học:"))
        self.cbo_khoa = QComboBox()
        self.cbo_khoa.setEditable(True)
        self.cbo_khoa.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_khoa.currentTextChanged.connect(lambda _: self._render())
        row1.addWidget(self.cbo_khoa, 1)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Nhóm theo:"))
        self.cbo_group = QComboBox()
        self.cbo_group.addItems(GROUP_OPTIONS)
        self.cbo_group.currentIndexChanged.connect(lambda _: self._render())
        row1.addWidget(self.cbo_group)

        btn_refresh = QPushButton("Làm mới")
        btn_refresh.clicked.connect(self.refresh)
        row1.addWidget(btn_refresh)
        form.addLayout(row1)

        row2 = QHBoxLayout()
        self.chk_all_time = QCheckBox("Tất cả thời gian")
        self.chk_all_time.setChecked(True)
        self.chk_all_time.toggled.connect(self._on_toggle_all_time)
        row2.addWidget(self.chk_all_time)

        row2.addWidget(QLabel("Kỳ nhanh:"))
        self.cbo_period = QComboBox()
        self.cbo_period.addItems(PERIOD_OPTIONS)
        self.cbo_period.currentIndexChanged.connect(self._on_period_changed)
        row2.addWidget(self.cbo_period)

        row2.addWidget(QLabel("Từ ngày:"))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd/MM/yyyy")
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_from.setEnabled(False)
        self.date_from.dateChanged.connect(self._on_date_edited)
        row2.addWidget(self.date_from)

        row2.addWidget(QLabel("Đến ngày:"))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("dd/MM/yyyy")
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setEnabled(False)
        self.date_to.dateChanged.connect(self._on_date_edited)
        row2.addWidget(self.date_to)

        row2.addStretch(1)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel...")
        self.btn_export.setToolTip("Xuất các phiếu đang hiện (theo bộ lọc) ra file Excel nhiều sheet.")
        self.btn_export.clicked.connect(self._on_export)
        row2.addWidget(self.btn_export)
        self.btn_export_pdf = QPushButton("Xuất PDF tóm tắt...")
        self.btn_export_pdf.setToolTip("Xuất 1 bản tóm tắt (tổng theo giáo viên) ra PDF - gọn hơn để in báo cáo giấy.")
        self.btn_export_pdf.clicked.connect(self._on_export_pdf)
        row2.addWidget(self.btn_export_pdf)
        self.btn_delete_selected = QPushButton("Xóa dòng đã chọn")
        self.btn_delete_selected.setEnabled(False)
        self.btn_delete_selected.clicked.connect(self._on_delete_selected)
        row2.addWidget(self.btn_delete_selected)
        btn_help = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogHelpButton), "Hướng dẫn")
        btn_help.clicked.connect(self._show_help)
        row2.addWidget(btn_help)
        form.addLayout(row2)

        # Tóm tắt 1 dòng các bộ lọc ĐANG áp dụng - trước đây chỉ thấy được
        # lúc xuất Excel (_filter_description() ghi vào hộp thoại), chọn
        # "Kỳ nhanh"/tự gõ Từ ngày-Đến ngày không thấy ngay đang lọc gì, dễ
        # tưởng đang xem "tất cả" khi thực ra đã bị lọc hẹp.
        self.lbl_filter_summary = QLabel("")
        self.lbl_filter_summary.setStyleSheet("color: #5a5f6b; font-style: italic;")
        form.addWidget(self.lbl_filter_summary)

        outer.addWidget(box)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(len(_COLUMN_LABELS))
        self.tree.setHeaderLabels(_COLUMN_LABELS)
        self.tree.setSortingEnabled(False)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
        self.tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        self.tree.setItemDelegateForColumn(0, _TeacherEditDelegate(self.model, self.tree))
        self.tree.itemSelectionChanged.connect(self._on_selection_changed)
        self.tree.itemChanged.connect(self._on_item_changed)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._on_context_menu)
        header = self.tree.header()
        header.setSortIndicatorShown(True)
        header.setSectionsMovable(True)
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.sectionClicked.connect(self._on_header_clicked)
        outer.addWidget(self.tree, 1)

        self.lbl_total = QLabel("")
        self.lbl_total.setStyleSheet("font-weight: 700; font-size: 12pt; color: #2659c4;")
        outer.addWidget(self.lbl_total)

        # Lý do + cách sửa của dòng KHÔNG khớp đang chọn - hiện cố định dưới
        # bảng (tooltip chỉ hiện khi rê chuột, dễ bỏ qua).
        self.lbl_orphan = QLabel("")
        self.lbl_orphan.setWordWrap(True)
        self.lbl_orphan.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.lbl_orphan.setStyleSheet(
            "background: #fff1dc; border: 1px solid #f0c48a; border-radius: 4px; padding: 6px; color: #5a3b00;"
        )
        self.lbl_orphan.hide()
        outer.addWidget(self.lbl_orphan)

        # Hiện ngắn sau khi xóa phiếu - cho phép khôi phục lại ngay nếu lỡ
        # tay (xem _on_delete_selected/_on_undo_delete), tự ẩn sau ít giây.
        undo_row = QHBoxLayout()
        self.lbl_undo = QLabel("")
        undo_row.addWidget(self.lbl_undo)
        self.btn_undo = QPushButton("Hoàn tác")
        self.btn_undo.clicked.connect(self._on_undo_delete)
        undo_row.addWidget(self.btn_undo)
        undo_row.addStretch(1)
        self.lbl_undo.hide()
        self.btn_undo.hide()
        outer.addLayout(undo_row)

    def showEvent(self, event):
        super().showEvent(event)
        self.refresh()

    def _show_help(self):
        QMessageBox.information(self, "Hướng dẫn - Tổng hợp cấp xăng", _HUONG_DAN_HTML)

    # ---------------- bộ lọc ----------------
    def _populate_filters(self):
        teacher_text = self.cbo_teacher.currentText().strip()
        khoa_text = self.cbo_khoa.currentText().strip()

        teachers = self.model.all_teachers()
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.clear()
        self.cbo_teacher.addItem(_ALL_LABEL)
        self.cbo_teacher.addItems(teachers)
        idx = self.cbo_teacher.findText(teacher_text) if teacher_text else 0
        self.cbo_teacher.setCurrentIndex(idx if idx >= 0 else 0)
        self.cbo_teacher.blockSignals(False)
        completer_t = QCompleter(teachers, self)
        completer_t.setCaseSensitivity(Qt.CaseInsensitive)
        completer_t.setFilterMode(Qt.MatchContains)
        self.cbo_teacher.setCompleter(completer_t)

        # Chỉ hiện khóa của ĐÚNG giáo viên đang lọc (nếu có chọn) - giống
        # combobox "Khóa học" ở tab CẤP XĂNG; khóa "Đã đào tạo" tô đỏ. Mã
        # khóa THẬT lưu riêng ở itemData (nhãn hiện kèm lớp/trạng thái, có
        # mã khóa chứa khoảng trắng như "...D2022 & K26D2023" nên không thể
        # suy ra mã khóa từ nhãn bằng cách tách chuỗi) - xem _filtered_records().
        assignments = self.model.all_assignments()
        if teacher_text and teacher_text != _ALL_LABEL:
            name_norm = attach_files.normalize_name(teacher_text)
            assignments = [a for a in assignments if attach_files.normalize_name(a["giao_vien"]) == name_norm]
        by_khoa = {}
        for a in assignments:
            by_khoa.setdefault(a["ma_khoa"], a)
        khoa_items = sorted(by_khoa.values(), key=lambda a: (a["status"] == "Đã đào tạo", a["ma_khoa"]))

        self.cbo_khoa.blockSignals(True)
        self.cbo_khoa.clear()
        self.cbo_khoa.addItem(_ALL_LABEL)
        khoa_labels = []
        for a in khoa_items:
            label = f"{a['ma_khoa']} ({a['lop']})"
            da_ket_thuc = a["status"] == "Đã đào tạo"
            if da_ket_thuc:
                label += " - ĐÃ KẾT THÚC"
            self.cbo_khoa.addItem(label, a["ma_khoa"])
            khoa_labels.append(label)
            if da_ket_thuc:
                self.cbo_khoa.setItemData(self.cbo_khoa.count() - 1, _MAU_KHOA_HET_HAN, Qt.ForegroundRole)
        idx2 = self.cbo_khoa.findText(khoa_text) if khoa_text else 0
        self.cbo_khoa.setCurrentIndex(idx2 if idx2 >= 0 else 0)
        self.cbo_khoa.blockSignals(False)
        completer_k = QCompleter(khoa_labels, self)
        completer_k.setCaseSensitivity(Qt.CaseInsensitive)
        completer_k.setFilterMode(Qt.MatchContains)
        self.cbo_khoa.setCompleter(completer_k)

    def refresh(self):
        """Nạp lại danh sách giáo viên/khóa + dữ liệu phiếu từ đĩa - gọi khi
        chuyển sang tab này (xem showEvent) và khi tab CẬP NHẬT DỮ LIỆU xử lý
        xong dữ liệu mới, để luôn thấy ngay các phiếu vừa nhập ở tab CẤP XĂNG."""
        self._populate_filters()
        self._render()

    def _on_teacher_filter_settled(self):
        """Đã CHỌN XONG giáo viên đang lọc (bấm chọn hoặc rời ô) -> nạp lại
        danh sách Khóa học cho khớp (chỉ còn khóa của giáo viên đó) - xem
        _populate_filters()."""
        self._populate_filters()
        self._render()

    def _on_toggle_all_time(self, checked):
        self.date_from.setEnabled(not checked)
        self.date_to.setEnabled(not checked)
        if checked:
            self._set_period_silently(0)
        self._render()

    def _set_period_silently(self, index):
        self.cbo_period.blockSignals(True)
        self.cbo_period.setCurrentIndex(index)
        self.cbo_period.blockSignals(False)

    def _on_date_edited(self, _date):
        # Người dùng tự chỉnh ngày -> không còn đúng kỳ nhanh đã chọn nữa
        self._set_period_silently(0)
        self._render()

    def _on_period_changed(self, index):
        if index == 0:
            return
        start, end = _period_range(PERIOD_OPTIONS[index], date.today())
        self.date_from.blockSignals(True)
        self.date_to.blockSignals(True)
        self.date_from.setDate(QDate(start.year, start.month, start.day))
        self.date_to.setDate(QDate(end.year, end.month, end.day))
        self.date_from.blockSignals(False)
        self.date_to.blockSignals(False)
        if self.chk_all_time.isChecked():
            # toggled -> _on_toggle_all_time -> _render (và không reset kỳ vì checked=False)
            self.chk_all_time.setChecked(False)
        else:
            self._render()

    def _filter_description(self):
        """Mô tả ngắn bộ lọc đang áp dụng - ghi vào hộp thoại sau khi xuất
        Excel để người dùng biết file chứa gì."""
        parts = []
        for label, cbo in (("Giáo viên", self.cbo_teacher), ("Khóa", self.cbo_khoa)):
            text = cbo.currentText().strip()
            if text and text != _ALL_LABEL:
                parts.append(f"{label}: {text}")
        if not self.chk_all_time.isChecked():
            parts.append(
                f"Ngày cấp: {self.date_from.date().toString('dd/MM/yyyy')} - "
                f"{self.date_to.date().toString('dd/MM/yyyy')}"
            )
        return "; ".join(parts) if parts else "Tất cả phiếu, mọi thời gian"

    def _filtered_records(self):
        all_records = cap_xang_store.load()["records"]
        recs = all_records

        teacher = self.cbo_teacher.currentText().strip()
        if teacher and teacher != _ALL_LABEL:
            name_norm = attach_files.normalize_name(teacher)
            recs = [r for r in recs if attach_files.normalize_name(r.get("giao_vien", "")) == name_norm]

        khoa = self.cbo_khoa.currentText().strip()
        if khoa and khoa != _ALL_LABEL:
            # Nhãn khớp đúng 1 item trong danh sách (chọn từ dropdown, hoặc
            # gõ tay trùng khớp) -> dùng mã khóa THẬT lưu ở itemData (nhãn có
            # thể kèm lớp/trạng thái, xem _populate_filters); gõ tự do mã
            # khóa cũ không có trong danh sách (vd lọc phiếu "mồ côi") thì
            # dùng nguyên văn đã gõ.
            idx = self.cbo_khoa.findText(khoa)
            khoa_ma = self.cbo_khoa.itemData(idx) if idx > 0 else khoa
            recs = [r for r in recs if r.get("ma_khoa", "") == khoa_ma]

        if not self.chk_all_time.isChecked():
            start = self.date_from.date().toPython()
            end = self.date_to.date().toPython()

            def _trong_khoang(r):
                d = _parse_ngay(r.get("ngay_cap", ""))
                return d is not None and start <= d <= end

            recs = [r for r in recs if _trong_khoang(r)]

        # id -> lý do không khớp phân công (None = khớp) - tính 1 lần cho các
        # phiếu qua được các bộ lọc trên, dùng lại lúc tô màu/tooltip/xuất.
        # KHÔNG còn ô lọc riêng theo "khớp/không khớp" nữa (tô cam + banner
        # giải thích đã đủ nổi bật để nhận ra ngay trong bảng, không cần ẩn
        # bớt các phiếu khác).
        self._orphan_reasons = {r["id"]: self.model.orphan_reason(r) for r in recs}

        return recs

    # ---------------- sắp xếp ----------------
    def _sort_value(self, r, field):
        if field == "ngay_cap":
            return _parse_ngay(r.get("ngay_cap", "")) or date.min
        if field == "so_luong":
            return r.get("so_luong") or 0.0
        if field == "giao_vien":
            return attach_files.given_name_sort_key(r.get("giao_vien", "") or "")
        return (r.get(field) or "").strip().lower()

    def _sorted_records(self, records):
        result = list(records)
        for field, asc in reversed(self._sort_keys):
            result.sort(key=lambda r: self._sort_value(r, field), reverse=not asc)
        return result

    def _on_header_clicked(self, col):
        if col >= len(_COLUMN_FIELDS) or _COLUMN_FIELDS[col] is None:
            return
        field = _COLUMN_FIELDS[col]
        shift = bool(QApplication.keyboardModifiers() & Qt.KeyboardModifier.ShiftModifier)

        existing_idx = next((i for i, (f, _) in enumerate(self._sort_keys) if f == field), None)
        if shift:
            if existing_idx is not None:
                f, asc = self._sort_keys[existing_idx]
                self._sort_keys[existing_idx] = (f, not asc)
            else:
                self._sort_keys.append((field, True))
        else:
            if existing_idx == 0 and len(self._sort_keys) == 1:
                f, asc = self._sort_keys[0]
                self._sort_keys = [(f, not asc)]
            else:
                self._sort_keys = [(field, True)]

        asc = next(a for f, a in self._sort_keys if f == field)
        self.tree.header().setSortIndicator(col, Qt.AscendingOrder if asc else Qt.DescendingOrder)
        self._render()

    # ---------------- nhóm + dựng bảng ----------------
    def _group_key(self, r, mode):
        if mode == "Giáo viên":
            name = r.get("giao_vien", "")
            return attach_files.normalize_name(name), name or "(không rõ giáo viên)"
        if mode == "Khóa học":
            ma_khoa = r.get("ma_khoa", "")
            label = f"{ma_khoa} ({r.get('lop', '')})" if ma_khoa else "(không rõ khóa học)"
            return ma_khoa, label
        if mode == "Biển số":
            xe = r.get("xe", "")
            return attach_files.normalize_plate(xe), xe or "(không rõ biển số)"
        return None, None

    def _add_leaf(self, r, parent):
        # KHÔNG gắn 1 QPushButton "Xóa" riêng cho mỗi dòng nữa (xem
        # _on_delete_selected/_on_context_menu) - dựng hàng trăm widget cùng
        # lúc mỗi lần _render() làm bảng bị chậm/đơ khi dữ liệu nhiều. Id của
        # phiếu được gắn vào item qua UserRole (cột 0) để lấy lại lúc xóa.
        so_luong_display = r.get("so_luong_text") or _fmt_so(r.get("so_luong"))
        values = [
            r.get("giao_vien", ""), r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""),
            r.get("ngay_cap", ""), r.get("so_phieu", ""), r.get("mat_hang", ""),
            so_luong_display, r.get("phan_loai_so", ""), r.get("ghi_chu", ""),
        ]
        item = QTreeWidgetItem(values)
        item.setData(0, Qt.ItemDataRole.UserRole, r["id"])
        diag = self._orphan_reasons.get(r["id"])
        if diag:
            tip = f"KHÔNG KHỚP PHÂN CÔNG: {diag['problem']}.\nCách sửa: {diag['fix']}"
            for col in range(len(values)):
                item.setBackground(col, _MAU_MO_COI)
                item.setToolTip(col, tip)
        if parent is None:
            self.tree.addTopLevelItem(item)
        else:
            parent.addChild(item)
        return item

    def _render(self):
        self.lbl_filter_summary.setText(f"Đang lọc: {self._filter_description()}")
        records = self._filtered_records()
        records = self._sorted_records(records)
        group_mode = self.cbo_group.currentText()
        # id -> record (dict gốc, SỐNG cùng self qua các lần sửa) - dùng để
        # lấy giá trị cũ lúc sửa 1 ô (xem _on_item_changed), không phải đọc
        # lại từ đĩa mỗi lần gõ.
        self._records_by_id = {r["id"]: r for r in records}

        # setText() khi dựng bảng cũng phát itemChanged - chặn tạm để khỏi
        # chạy nhầm _on_item_changed (tưởng người dùng vừa sửa) lúc đang dựng.
        self.tree.blockSignals(True)
        self.tree.clear()
        if group_mode == "Không nhóm":
            for r in records:
                self._add_leaf(r, parent=None)
        else:
            groups = {}
            order = []
            for r in records:
                key, label = self._group_key(r, group_mode)
                if key not in groups:
                    groups[key] = [label, []]
                    order.append(key)
                groups[key][1].append(r)
            order.sort(key=lambda k: groups[k][0].lower())

            for key in order:
                label, recs = groups[key]
                tong_lit = sum(x.get("so_luong") or 0 for x in recs)
                tong_phieu = sum(x.get("so_phieu_count") or 1 for x in recs)
                group_item = QTreeWidgetItem([f"{label}  —  {len(recs)} phiếu, {_fmt_so(tong_lit)} lít"])
                group_item.setFlags(group_item.flags() & ~Qt.ItemIsEditable)
                group_item.setFirstColumnSpanned(True)
                font = group_item.font(0)
                font.setBold(True)
                group_item.setFont(0, font)
                group_item.setBackground(0, QColor("#eef1f8"))
                self.tree.addTopLevelItem(group_item)
                group_item.setExpanded(True)
                for r in recs:
                    self._add_leaf(r, parent=group_item)

        self.tree.blockSignals(False)

        for col in range(len(_COLUMN_LABELS)):
            self.tree.resizeColumnToContents(col)
        self._on_selection_changed()

        if not records:
            if not cap_xang_store.load()["records"]:
                self.lbl_total.setText(
                    "Chưa có phiếu cấp xăng nào. Qua tab \"CẤP XĂNG\" để bắt đầu nhập phiếu."
                )
            else:
                self.lbl_total.setText("Không có phiếu nào khớp điều kiện lọc.")
            return
        tong_lit = sum(r.get("so_luong") or 0 for r in records)
        tong_phieu = sum(r.get("so_phieu_count") or 1 for r in records)
        text = f"Tổng cộng: {_fmt_so(tong_lit)} lít ({tong_phieu} phiếu, {len(records)} dòng)"
        so_mo_coi = sum(1 for r in records if self._orphan_reasons.get(r["id"]))
        if so_mo_coi:
            text += f"  —  trong đó {so_mo_coi} dòng KHÔNG khớp phân công (tô cam)"
        self.lbl_total.setText(text)

    def _selected_record_ids(self):
        """Id các phiếu của những dòng đang chọn (có thể chọn nhiều dòng qua
        Ctrl/Shift+click) - bỏ qua dòng tiêu đề nhóm (không mang id phiếu)."""
        ids = []
        for item in self.tree.selectedItems():
            rid = item.data(0, Qt.ItemDataRole.UserRole)
            if rid is not None and rid not in ids:
                ids.append(rid)
        return ids

    def _on_selection_changed(self):
        ids = self._selected_record_ids()
        self.btn_delete_selected.setEnabled(bool(ids))
        self.btn_delete_selected.setText(
            f"Xóa {len(ids)} dòng đã chọn" if len(ids) > 1 else "Xóa dòng đã chọn"
        )
        diag = self._orphan_reasons.get(ids[0]) if len(ids) == 1 else None
        if diag:
            self.lbl_orphan.setText(
                f"<b>Dòng đang chọn KHÔNG khớp phân công</b> — {diag['problem']}.<br>"
                f"<b>Cách sửa:</b> {diag['fix']} (double-click vào ô cần sửa ngay trên bảng.)"
            )
            self.lbl_orphan.show()
        else:
            self.lbl_orphan.hide()

    def _on_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if item is None or item.data(0, Qt.ItemDataRole.UserRole) is None:
            return
        if item not in self.tree.selectedItems():
            self.tree.setCurrentItem(item)
        menu = QMenu(self)
        n = len(self._selected_record_ids())
        menu.addAction(f"Xóa ({n} phiếu)" if n > 1 else "Xóa", self._on_delete_selected)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def _on_delete_selected(self):
        ids = self._selected_record_ids()
        if not ids:
            return
        text = f"Xóa {len(ids)} phiếu cấp xăng đã chọn?" if len(ids) > 1 else "Xóa phiếu cấp xăng này?"
        confirm = QMessageBox.question(self, "Xác nhận", text)
        if confirm != QMessageBox.StandardButton.Yes:
            return
        cap_xang_store.delete_records(ids)
        self._show_undo(len(ids))
        self._render()
        self.du_lieu_da_doi.emit()

    def _show_undo(self, n):
        self.lbl_undo.setText(f"Đã xóa {n} phiếu." if n > 1 else "Đã xóa 1 phiếu.")
        self.lbl_undo.show()
        self.btn_undo.show()
        QTimer.singleShot(15000, self._hide_undo)

    def _hide_undo(self):
        self.lbl_undo.hide()
        self.btn_undo.hide()

    def _on_undo_delete(self):
        if cap_xang_store.restore_backup():
            self._hide_undo()
            self._populate_filters()
            self._render()
            self.du_lieu_da_doi.emit()
        else:
            QMessageBox.information(self, "Không có gì để hoàn tác", "Không tìm thấy bản sao để khôi phục.")

    # ---------------- sửa trực tiếp trên bảng ----------------
    def _revert_cell(self, item, column, old_text):
        """Đặt lại ô về giá trị cũ (nhập sai/không hợp lệ) - chặn tạm tín
        hiệu để khỏi gọi lại chính _on_item_changed cho giá trị cũ này."""
        self.tree.blockSignals(True)
        item.setText(column, old_text)
        self.tree.blockSignals(False)

    def _on_item_changed(self, item, column):
        record_id = item.data(0, Qt.ItemDataRole.UserRole)
        if record_id is None or column >= len(_COLUMN_FIELDS):
            return
        record = self._records_by_id.get(record_id)
        if record is None:
            return

        field = _COLUMN_FIELDS[column]
        new_text = item.text(column).strip()

        if field == "ngay_cap":
            old_text = record.get("ngay_cap", "")
            parsed = _parse_ngay(new_text)
            if parsed is None:
                QMessageBox.warning(self, "Ngày không hợp lệ", "Gõ ngày theo dạng dd/mm/yyyy.")
                self._revert_cell(item, column, old_text)
                return
            new_text = parsed.strftime("%d/%m/%Y")
            fields = {"ngay_cap": new_text}
        elif field == "so_luong":
            old_text = record.get("so_luong_text") or _fmt_so(record.get("so_luong"))
            tong_lit, tong_phieu = parse_so_luong(new_text)
            if tong_lit is None:
                QMessageBox.warning(self, "Số lượng không đúng định dạng", _SO_LUONG_GHI_CHU)
                self._revert_cell(item, column, old_text)
                return
            if tong_lit <= 0:
                QMessageBox.warning(self, "Số lượng không hợp lệ", "Số lượng phải lớn hơn 0 - xóa cả phiếu thì dùng \"Xóa dòng đã chọn\".")
                self._revert_cell(item, column, old_text)
                return
            fields = {"so_luong": tong_lit, "so_luong_text": new_text, "so_phieu_count": tong_phieu}
        else:
            old_text = record.get(field, "") or ""
            fields = {field: new_text}

        if new_text == old_text:
            return  # không đổi gì thật sự (vd chỉ bấm vào ô rồi Enter) - khỏi ghi/dựng lại

        cap_xang_store.update_record(record_id, fields)
        record.update(fields)
        self.du_lieu_da_doi.emit()
        # Dựng lại ngay sau khi tín hiệu hiện tại (itemChanged) xử lý xong -
        # tránh sửa cây đang phát tín hiệu ngay trong lúc nó chưa chạy xong,
        # giống cách các tab khác xử lý (xem tab_cap_xang.py _make_commit_handler).
        QTimer.singleShot(0, self._render)

    # ---------------- xuất Excel ----------------
    def _on_export(self):
        records = self._sorted_records(self._filtered_records())
        if not records:
            QMessageBox.information(self, "Không có dữ liệu", "Không có phiếu nào khớp bộ lọc để xuất.")
            return
        default_name = f"CapXang_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất Excel", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            self._write_report(path, records)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã xuất {len(records)} dòng phiếu ra:\n{path}\n\nBộ lọc đã áp dụng: {self._filter_description()}\n\n{GHI_CHU_TO_DO}",
        )

    # ---------------- xuất PDF tóm tắt ----------------
    def _on_export_pdf(self):
        records = self._sorted_records(self._filtered_records())
        if not records:
            QMessageBox.information(self, "Không có dữ liệu", "Không có phiếu nào khớp bộ lọc để xuất.")
            return
        default_name = f"CapXang_TomTat_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất PDF tóm tắt", default_name, "PDF Files (*.pdf)")
        if not path:
            return
        try:
            self._write_pdf_summary(path, records)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        QMessageBox.information(self, "Hoàn tất", f"Đã xuất bản tóm tắt ra:\n{path}")

    def _write_pdf_summary(self, path, records):
        """1 bản tóm tắt gọn (tổng theo giáo viên) - khác "Xuất Excel..."
        (đầy đủ nhiều sheet, để lưu/tra cứu) - dùng in báo cáo giấy định kỳ."""
        tong_lit = sum(r.get("so_luong") or 0 for r in records)
        tong_phieu = sum(r.get("so_phieu_count") or 1 for r in records)
        so_mo_coi = sum(1 for r in records if self._orphan_reasons.get(r["id"]))

        by_gv = _aggregate(records, lambda r: r.get("giao_vien", ""))
        rows_html = "".join(
            f"<tr><td>{gv or '(không rõ)'}</td><td align='right'>{a['phieu']}</td>"
            f"<td align='right'>{_fmt_so(a['lit'])}</td></tr>"
            for gv, a in sorted(by_gv.items(), key=lambda kv: attach_files.given_name_sort_key(kv[0]))
        )
        ghi_chu_mo_coi = (
            f"<p>Trong đó <b>{so_mo_coi}</b> dòng KHÔNG khớp phân công hiện tại (xem chi tiết ở "
            "tab TỔNG HỢP CẤP XĂNG, bộ lọc \"Phân công\").</p>" if so_mo_coi else ""
        )
        html = f"""
        <h2>Báo cáo cấp xăng</h2>
        <p>Bộ lọc: {self._filter_description()}</p>
        <p><b>Tổng cộng:</b> {_fmt_so(tong_lit)} lít ({tong_phieu} phiếu, {len(records)} dòng)</p>
        {ghi_chu_mo_coi}
        <table border="1" cellspacing="0" cellpadding="4" width="100%">
          <tr><th>Giáo viên</th><th>Số phiếu</th><th>Tổng (lít)</th></tr>
          {rows_html}
        </table>
        <p style="color:#888;font-size:9pt;">Xuất lúc {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
        """
        doc = QTextDocument()
        doc.setHtml(html)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(path)
        doc.print_(printer)

    def _chi_tiet_sort_note(self):
        """Mô tả đúng thứ tự sắp xếp ĐANG DÙNG ở bảng trên màn hình (người
        dùng tự chọn bằng cách bấm vào tiêu đề cột, xem _on_header_clicked) -
        sheet "Chi tiết phiếu" xuất RA ĐÚNG thứ tự đó nên cần ghi rõ, kẻo mở
        file sau rồi quên đã lọc/sắp theo gì."""
        label_by_field = dict(zip(_COLUMN_FIELDS, _COLUMN_LABELS))
        parts = [
            f"{label_by_field.get(field, field)} ({'tăng dần' if asc else 'giảm dần'})"
            for field, asc in self._sort_keys
        ]
        return "Sắp xếp theo: " + ", ".join(parts) if parts else None

    def _write_report(self, path, records):
        report = ReportWorkbook()

        def _ngay(r):
            return _parse_ngay(r.get("ngay_cap", "")) or r.get("ngay_cap", "")

        report.add_sheet(
            "Chi tiet phieu",
            ["Giáo viên", "Khóa học", "Lớp", "Biển số", "Ngày cấp", "Số phiếu/Quyển", "Mặt hàng",
             "Số lượng (lít)", "Số phiếu", "Phân loại sổ", "Ghi chú",
             "Khớp phân công", "Không khớp - lý do", "Cách sửa"],
            [
                [
                    r.get("giao_vien", ""), r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""),
                    _ngay(r), r.get("so_phieu", ""), r.get("mat_hang", ""), r.get("so_luong") or 0,
                    r.get("so_phieu_count") or 1, r.get("phan_loai_so", ""), r.get("ghi_chu", ""),
                    "Không" if diag else "Có", diag["problem"] if diag else "", diag["fix"] if diag else "",
                ]
                for r in records
                for diag in (self._orphan_reasons.get(r["id"]),)
            ],
            highlight=[bool(self._orphan_reasons.get(r["id"])) for r in records],
            sort_note=self._chi_tiet_sort_note(),
        )

        by_gv = _aggregate(records, lambda r: r.get("giao_vien", ""))
        report.add_sheet(
            "Theo giao vien",
            ["Giáo viên", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [
                [gv, a["n"], a["phieu"], a["lit"]]
                for gv, a in sorted(by_gv.items(), key=lambda kv: attach_files.given_name_sort_key(kv[0]))
            ],
            sort_note="Sắp xếp theo: Giáo viên (tên, A→Z - không theo họ).",
        )

        by_khoa = _aggregate(
            records, lambda r: (r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""), r.get("giao_vien", ""))
        )
        orphan_keys = {
            (r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""), r.get("giao_vien", ""))
            for r in records if self._orphan_reasons.get(r["id"])
        }
        khoa_keys = sorted(by_khoa)
        report.add_sheet(
            "Theo khoa-xe",
            ["Khóa học", "Lớp", "Biển số", "Giáo viên", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[*key, by_khoa[key]["n"], by_khoa[key]["phieu"], by_khoa[key]["lit"]] for key in khoa_keys],
            highlight=[key in orphan_keys for key in khoa_keys],
            sort_note="Sắp xếp theo: Khóa học → Lớp → Biển số → Giáo viên (A→Z).",
        )

        by_xe = _aggregate(records, lambda r: r.get("xe", ""))
        report.add_sheet(
            "Theo bien so",
            ["Biển số", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[xe, a["n"], a["phieu"], a["lit"]] for xe, a in sorted(by_xe.items())],
            sort_note="Sắp xếp theo: Biển số (A→Z).",
        )

        months = sorted({_month_key(r) for r in records}, key=lambda k: (k is None, k))
        by_thang = _aggregate(records, _month_key)
        report.add_sheet(
            "Theo thang",
            ["Tháng", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[_month_label(m), by_thang[m]["n"], by_thang[m]["phieu"], by_thang[m]["lit"]] for m in months],
            sort_note="Sắp xếp theo: Tháng (tăng dần, tháng/năm trống xếp cuối).",
        )

        by_gv_thang = _aggregate(records, lambda r: (r.get("giao_vien", ""), _month_key(r)))
        pivot_rows = []
        for gv in sorted(by_gv, key=attach_files.given_name_sort_key):
            cells = [by_gv_thang[(gv, m)]["lit"] if (gv, m) in by_gv_thang else None for m in months]
            pivot_rows.append([gv, *cells, by_gv[gv]["lit"]])
        report.add_sheet(
            "GV theo thang (lit)",
            ["Giáo viên", *(_month_label(m) for m in months), "Tổng (lít)"],
            pivot_rows,
            sort_note="Sắp xếp theo: Giáo viên (tên, A→Z); các cột tháng tăng dần từ trái qua phải.",
        )

        teacher = self.cbo_teacher.currentText().strip()
        dinh_muc_rows = self.model.fuel_rows_for_teacher(
            teacher if teacher and teacher != _ALL_LABEL else None
        )
        report.add_sheet(
            "Dinh muc - con lai",
            _DINH_MUC_COLUMNS,
            [dinh_muc_export_row(r) for r in dinh_muc_rows],
            highlight=[is_vuot(r) for r in dinh_muc_rows],
            sort_note="Không sắp xếp lại - giữ theo thứ tự khóa học trong dữ liệu gốc.",
        )

        report.save(path)

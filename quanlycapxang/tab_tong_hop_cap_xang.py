"""Tab TỔNG HỢP CẤP XĂNG: xem toàn bộ phiếu cấp xăng của CẢ TRUNG TÂM (mọi
giáo viên) trong 1 bảng, không giới hạn theo 1 giáo viên như tab CẤP XĂNG.

- Lọc theo Giáo viên / Khóa học / khoảng ngày cấp (tùy chọn, mặc định xem hết).
- "Nhóm theo": gom các phiếu theo Giáo viên/Khóa học/Biển số thành từng nhóm
  thu/mở được, mỗi nhóm có tổng số lít/số phiếu riêng.
- Bấm vào tên cột để sắp xếp theo cột đó (bấm lại đổi tăng/giảm) - giữ Shift
  bấm thêm cột khác để sắp theo nhiều cấp (vd Giáo viên trước, Ngày cấp sau).
  Kéo tiêu đề cột để đổi thứ tự hiển thị các cột.
- Chọn 1 dòng rồi bấm "Xóa dòng đã chọn" (hoặc chuột phải -> Xóa) để xóa
  phiếu đó. (Không còn nút "Xóa" riêng ở mỗi dòng như trước - hàng trăm nút
  cùng lúc làm bảng dựng lại rất chậm/đơ khi dữ liệu nhiều; chọn dòng rồi
  xóa 1 nút chung vừa nhanh vừa nhẹ hơn hẳn.)
- SỬA TRỰC TIẾP: double-click vào bất kỳ ô nào (trừ dòng tiêu đề nhóm) để
  sửa, Enter/rời ô để lưu ngay - không cần quay lại tab CẤP XĂNG nữa. Hữu
  ích nhất cho các phiếu không khớp khóa/xe đang phân công hiện tại (nhập
  tay mã khóa cũ/tên viết tắt...) - trước đây các phiếu này chỉ xem/xóa
  được ở đây, không sửa được ở đâu cả.
- Lọc thêm theo Biển số; "Kỳ nhanh" chọn sẵn tháng này/tháng trước/quý
  này/năm nay cho báo cáo định kỳ, không phải bấm lịch 2 lần.
- PHIẾU KHÔNG KHỚP PHÂN CÔNG ("mồ côi" - xem CapXangModel.orphan_reason): tô
  cam nhạt + rê chuột thấy lý do; ô "Phân công" lọc riêng chỉ các phiếu này
  để rà soát định kỳ (sửa lại mã khóa/xe/tên cho khớp, hoặc xóa).
- "Xuất Excel...": xuất theo đúng bộ lọc đang chọn, nhiều sheet (chi tiết +
  tổng theo giáo viên/khóa-xe/biển số/tháng + bảng giáo viên x tháng + định
  mức), mỗi sheet là bảng Excel có nút lọc/sort sẵn."""
from datetime import date, datetime, timedelta

from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtGui import QColor
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
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

import attach_files
import cap_xang_store
from excel_export import ReportWorkbook
from style import icon as _icon
from tab_cap_xang import parse_so_luong
from tab_dinh_muc_xang import EXPORT_COLUMNS as _DINH_MUC_COLUMNS, dinh_muc_export_row

_SO_LUONG_GHI_CHU = (
    "Gõ 1 số (vd 50) hoặc SỐ PHIẾUxSỐ LÍT, nhiều phiếu cách nhau bằng dấu "
    "phẩy (vd 2x30, 4x20) - xem hướng dẫn ở tab CẤP XĂNG."
)

_ALL_LABEL = "-- Tất cả --"
GROUP_OPTIONS = ["Không nhóm", "Giáo viên", "Khóa học", "Biển số"]
PERIOD_OPTIONS = ["Tùy chọn", "Tuần này", "Tháng này", "Tháng trước", "Quý này", "Năm nay"]
MATCH_OPTIONS = ["Tất cả phiếu", "Chỉ phiếu khớp phân công", "Chỉ phiếu KHÔNG khớp (mồ côi)"]
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
    "xem hết).</li>"
    "<li>Lọc thêm theo <b>Biển số</b> nếu cần xem riêng 1 xe.</li>"
    "<li><b>Khoảng ngày</b>: bỏ tick \"Tất cả thời gian\" để lọc theo 1 "
    "khoảng ngày cấp cụ thể, hoặc chọn <b>Kỳ nhanh</b> (tháng này/tháng "
    "trước/quý này/năm nay) - app tự điền Từ ngày/Đến ngày.</li>"
    "<li><b>Phân công</b>: lọc riêng các phiếu <b>KHÔNG khớp</b> khóa/xe/giáo "
    "viên đang phân công (tô cam nhạt). <b>Bấm vào dòng</b> -> ô cam dưới bảng "
    "ghi rõ sai ở ô nào và cách sửa (gợi ý mã khóa/tên/biển số đúng nếu suy ra "
    "được). 3 loại: (1) <i>ô Khóa học trống</i> - phiếu nhập từ Excel cũ chỉ có mã "
    "lớp, khóa đã kết thúc không còn trong dữ liệu -> KHÔNG cần sửa, vẫn tính "
    "vào tổng; (2) <i>tên giáo viên lạ</i> -> sửa ô Giáo viên theo gợi ý; (3) "
    "<i>khóa có nhưng giáo viên/xe không đúng phân công</i> -> sửa ô Giáo viên "
    "hoặc Biển số theo gợi ý. Phiếu không khớp không hiện ở lưới Chi tiết tab "
    "CẤP XĂNG và không trừ vào \"Còn lại\" của khóa nào ở tab ĐỊNH MỨC XĂNG.</li>"
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
    "<li>Chọn 1 dòng rồi bấm <b>\"Xóa dòng đã chọn\"</b> (hoặc chuột phải vào "
    "dòng đó -> Xóa) để xóa phiếu.</li>"
    "<li><b>Double-click vào 1 ô</b> (trừ dòng tiêu đề nhóm) để sửa trực "
    "tiếp, Enter/rời ô để lưu ngay - không cần quay lại tab CẤP XĂNG. Cột "
    "\"Ngày cấp\" phải đúng dạng dd/mm/yyyy, cột \"Số lượng\" gõ như hướng "
    "dẫn ở tab CẤP XĂNG (vd 2x30, 4x20) - gõ sai sẽ báo lỗi và giữ nguyên "
    "giá trị cũ.</li>"
    "</ul>"
)


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
        self.cbo_teacher.currentTextChanged.connect(lambda _: self._render())
        row1.addWidget(self.cbo_teacher, 1)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Khóa học:"))
        self.cbo_khoa = QComboBox()
        self.cbo_khoa.setEditable(True)
        self.cbo_khoa.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_khoa.currentTextChanged.connect(lambda _: self._render())
        row1.addWidget(self.cbo_khoa, 1)

        row1.addSpacing(16)
        row1.addWidget(QLabel("Biển số:"))
        self.cbo_xe = QComboBox()
        self.cbo_xe.setEditable(True)
        self.cbo_xe.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_xe.currentTextChanged.connect(lambda _: self._render())
        row1.addWidget(self.cbo_xe, 1)

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

        row2.addSpacing(16)
        row2.addWidget(QLabel("Phân công:"))
        self.cbo_match = QComboBox()
        self.cbo_match.addItems(MATCH_OPTIONS)
        self.cbo_match.setToolTip(
            "Phiếu KHÔNG khớp = giáo viên/mã khóa/xe trên phiếu không trùng với "
            "phân công hiện có (rê chuột vào dòng cam để xem lý do)."
        )
        self.cbo_match.currentIndexChanged.connect(lambda _: self._render())
        row2.addWidget(self.cbo_match)

        row2.addStretch(1)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel...")
        self.btn_export.setToolTip("Xuất các phiếu đang hiện (theo bộ lọc) ra file Excel nhiều sheet.")
        self.btn_export.clicked.connect(self._on_export)
        row2.addWidget(self.btn_export)
        self.btn_delete_selected = QPushButton("Xóa dòng đã chọn")
        self.btn_delete_selected.setEnabled(False)
        self.btn_delete_selected.clicked.connect(self._on_delete_selected)
        row2.addWidget(self.btn_delete_selected)
        btn_help = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogHelpButton), "Hướng dẫn")
        btn_help.clicked.connect(self._show_help)
        row2.addWidget(btn_help)
        form.addLayout(row2)

        outer.addWidget(box)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(len(_COLUMN_LABELS))
        self.tree.setHeaderLabels(_COLUMN_LABELS)
        self.tree.setSortingEnabled(False)
        self.tree.setAlternatingRowColors(True)
        self.tree.setRootIsDecorated(True)
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

        khoas = sorted(self.model.course_codes())
        self.cbo_khoa.blockSignals(True)
        self.cbo_khoa.clear()
        self.cbo_khoa.addItem(_ALL_LABEL)
        self.cbo_khoa.addItems(khoas)
        idx2 = self.cbo_khoa.findText(khoa_text) if khoa_text else 0
        self.cbo_khoa.setCurrentIndex(idx2 if idx2 >= 0 else 0)
        self.cbo_khoa.blockSignals(False)
        completer_k = QCompleter(khoas, self)
        completer_k.setCaseSensitivity(Qt.CaseInsensitive)
        completer_k.setFilterMode(Qt.MatchContains)
        self.cbo_khoa.setCompleter(completer_k)

        xe_text = self.cbo_xe.currentText().strip()
        xes = sorted(cap_xang_store.distinct_values("xe"))
        self.cbo_xe.blockSignals(True)
        self.cbo_xe.clear()
        self.cbo_xe.addItem(_ALL_LABEL)
        self.cbo_xe.addItems(xes)
        idx3 = self.cbo_xe.findText(xe_text) if xe_text else 0
        self.cbo_xe.setCurrentIndex(idx3 if idx3 >= 0 else 0)
        self.cbo_xe.blockSignals(False)
        completer_x = QCompleter(xes, self)
        completer_x.setCaseSensitivity(Qt.CaseInsensitive)
        completer_x.setFilterMode(Qt.MatchContains)
        self.cbo_xe.setCompleter(completer_x)

    def refresh(self):
        """Nạp lại danh sách giáo viên/khóa + dữ liệu phiếu từ đĩa - gọi khi
        chuyển sang tab này (xem showEvent) và khi tab CẬP NHẬT DỮ LIỆU xử lý
        xong dữ liệu mới, để luôn thấy ngay các phiếu vừa nhập ở tab CẤP XĂNG."""
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
        for label, cbo in (("Giáo viên", self.cbo_teacher), ("Khóa", self.cbo_khoa), ("Biển số", self.cbo_xe)):
            text = cbo.currentText().strip()
            if text and text != _ALL_LABEL:
                parts.append(f"{label}: {text}")
        if not self.chk_all_time.isChecked():
            parts.append(
                f"Ngày cấp: {self.date_from.date().toString('dd/MM/yyyy')} - "
                f"{self.date_to.date().toString('dd/MM/yyyy')}"
            )
        if self.cbo_match.currentIndex() > 0:
            parts.append(self.cbo_match.currentText())
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
            recs = [r for r in recs if r.get("ma_khoa", "") == khoa]

        xe = self.cbo_xe.currentText().strip()
        if xe and xe != _ALL_LABEL:
            plate = attach_files.normalize_plate(xe)
            recs = [r for r in recs if attach_files.normalize_plate(r.get("xe", "")) == plate]

        if not self.chk_all_time.isChecked():
            start = self.date_from.date().toPython()
            end = self.date_to.date().toPython()

            def _trong_khoang(r):
                d = _parse_ngay(r.get("ngay_cap", ""))
                return d is not None and start <= d <= end

            recs = [r for r in recs if _trong_khoang(r)]

        # id -> lý do không khớp phân công (None = khớp) - tính 1 lần cho các
        # phiếu qua được các bộ lọc trên, dùng lại lúc tô màu/tooltip/xuất.
        self._orphan_reasons = {r["id"]: self.model.orphan_reason(r) for r in recs}
        match_mode = self.cbo_match.currentIndex()
        if match_mode == 1:
            recs = [r for r in recs if self._orphan_reasons[r["id"]] is None]
        elif match_mode == 2:
            recs = [r for r in recs if self._orphan_reasons[r["id"]] is not None]

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
            self.lbl_total.setText("Không có phiếu nào khớp điều kiện lọc.")
            return
        tong_lit = sum(r.get("so_luong") or 0 for r in records)
        tong_phieu = sum(r.get("so_phieu_count") or 1 for r in records)
        text = f"Tổng cộng: {_fmt_so(tong_lit)} lít ({tong_phieu} phiếu, {len(records)} dòng)"
        so_mo_coi = sum(1 for r in records if self._orphan_reasons.get(r["id"]))
        if so_mo_coi:
            text += f"  —  trong đó {so_mo_coi} dòng KHÔNG khớp phân công (tô cam)"
        self.lbl_total.setText(text)

    def _selected_record_id(self):
        """Id phiếu của dòng đang chọn - None nếu không chọn dòng nào hoặc
        đang chọn dòng tiêu đề nhóm (không mang id phiếu)."""
        item = self.tree.currentItem()
        if item is None:
            return None
        return item.data(0, Qt.ItemDataRole.UserRole)

    def _on_selection_changed(self):
        record_id = self._selected_record_id()
        self.btn_delete_selected.setEnabled(record_id is not None)
        diag = self._orphan_reasons.get(record_id) if record_id is not None else None
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
        self.tree.setCurrentItem(item)
        menu = QMenu(self)
        menu.addAction("Xóa", self._on_delete_selected)
        menu.exec(self.tree.viewport().mapToGlobal(pos))

    def _on_delete_selected(self):
        record_id = self._selected_record_id()
        if record_id is None:
            return
        self._on_delete(record_id)

    def _on_delete(self, record_id):
        confirm = QMessageBox.question(self, "Xác nhận", "Xóa phiếu cấp xăng này?")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        cap_xang_store.delete_record(record_id)
        self._render()

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
            f"Đã xuất {len(records)} dòng phiếu ra:\n{path}\n\nBộ lọc đã áp dụng: {self._filter_description()}",
        )

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
        )

        by_gv = _aggregate(records, lambda r: r.get("giao_vien", ""))
        report.add_sheet(
            "Theo giao vien",
            ["Giáo viên", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [
                [gv, a["n"], a["phieu"], a["lit"]]
                for gv, a in sorted(by_gv.items(), key=lambda kv: attach_files.given_name_sort_key(kv[0]))
            ],
        )

        by_khoa = _aggregate(
            records, lambda r: (r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""), r.get("giao_vien", ""))
        )
        report.add_sheet(
            "Theo khoa-xe",
            ["Khóa học", "Lớp", "Biển số", "Giáo viên", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[*key, a["n"], a["phieu"], a["lit"]] for key, a in sorted(by_khoa.items())],
        )

        by_xe = _aggregate(records, lambda r: r.get("xe", ""))
        report.add_sheet(
            "Theo bien so",
            ["Biển số", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[xe, a["n"], a["phieu"], a["lit"]] for xe, a in sorted(by_xe.items())],
        )

        months = sorted({_month_key(r) for r in records}, key=lambda k: (k is None, k))
        by_thang = _aggregate(records, _month_key)
        report.add_sheet(
            "Theo thang",
            ["Tháng", "Số dòng", "Số phiếu", "Tổng (lít)"],
            [[_month_label(m), by_thang[m]["n"], by_thang[m]["phieu"], by_thang[m]["lit"]] for m in months],
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
        )

        teacher = self.cbo_teacher.currentText().strip()
        dinh_muc_rows = self.model.fuel_rows_for_teacher(
            teacher if teacher and teacher != _ALL_LABEL else None
        )
        report.add_sheet(
            "Dinh muc - con lai",
            _DINH_MUC_COLUMNS,
            [dinh_muc_export_row(r) for r in dinh_muc_rows],
        )

        report.save(path)

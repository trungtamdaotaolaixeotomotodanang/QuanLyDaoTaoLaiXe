"""Tab TỔNG HỢP CẤP XĂNG: xem toàn bộ phiếu cấp xăng của CẢ TRUNG TÂM (mọi
giáo viên) trong 1 bảng, không giới hạn theo 1 giáo viên như tab CẤP XĂNG.

- Lọc theo Giáo viên / Khóa học / khoảng ngày cấp (tùy chọn, mặc định xem hết).
- "Nhóm theo": gom các phiếu theo Giáo viên/Khóa học/Biển số thành từng nhóm
  thu/mở được, mỗi nhóm có tổng số lít/số phiếu riêng.
- Bấm vào tên cột để sắp xếp theo cột đó (bấm lại đổi tăng/giảm) - giữ Shift
  bấm thêm cột khác để sắp theo nhiều cấp (vd Giáo viên trước, Ngày cấp sau).
  Kéo tiêu đề cột để đổi thứ tự hiển thị các cột.
- Nút "Xóa" ở cuối mỗi dòng để xóa phiếu đó - sửa chi tiết thì quay lại tab
  CẤP XĂNG (chọn đúng giáo viên)."""
from datetime import date, datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDateEdit,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
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
from style import icon as _icon

_ALL_LABEL = "-- Tất cả --"
GROUP_OPTIONS = ["Không nhóm", "Giáo viên", "Khóa học", "Biển số"]

_COLUMN_LABELS = [
    "Giáo viên", "Khóa học", "Lớp", "Biển số", "Ngày cấp", "Số phiếu/Quyển",
    "Mặt hàng", "Số lượng", "Phân loại sổ", "Ghi chú", "",
]
_COLUMN_FIELDS = [
    "giao_vien", "ma_khoa", "lop", "xe", "ngay_cap", "so_phieu",
    "mat_hang", "so_luong", "phan_loai_so", "ghi_chu", None,
]

_HUONG_DAN_HTML = (
    "<p><b>Tab \"TỔNG HỢP CẤP XĂNG\"</b> - xem toàn bộ phiếu cấp xăng của CẢ "
    "TRUNG TÂM (mọi giáo viên), không giới hạn theo 1 giáo viên như tab CẤP "
    "XĂNG.</p>"
    "<ul>"
    "<li>Lọc theo <b>Giáo viên</b> / <b>Khóa học</b> (để \"-- Tất cả --\" để "
    "xem hết).</li>"
    "<li><b>Khoảng ngày</b>: bỏ tick \"Tất cả thời gian\" để lọc theo 1 "
    "khoảng ngày cấp cụ thể (vd làm báo cáo theo tháng).</li>"
    "<li><b>Nhóm theo</b>: gom các phiếu theo Giáo viên/Khóa học/Biển số "
    "thành từng nhóm có tổng số lít/số phiếu riêng - bấm vào nhóm để thu/mở.</li>"
    "<li><b>Bấm vào tên cột</b> (vd \"Ngày cấp\", \"Số lượng\"...) để sắp xếp "
    "theo cột đó, bấm lại để đổi tăng/giảm. Giữ <b>Shift</b> rồi bấm thêm cột "
    "khác để sắp theo nhiều cấp (vd sắp theo Giáo viên trước, trong mỗi giáo "
    "viên sắp theo Ngày cấp).</li>"
    "<li>Có thể <b>kéo tiêu đề cột</b> sang trái/phải để đổi thứ tự hiển thị "
    "các cột theo ý muốn.</li>"
    "<li>Nút <b>\"Xóa\"</b> ở cuối mỗi dòng để xóa phiếu đó - muốn sửa chi "
    "tiết (ngày/số lượng/ghi chú...) thì quay lại tab CẤP XĂNG, chọn đúng "
    "giáo viên.</li>"
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


class TongHopCapXangTab(QWidget):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._sort_keys = [("ngay_cap", False)]  # mặc định: ngày cấp mới nhất trước
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

        row2.addWidget(QLabel("Từ ngày:"))
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("dd/MM/yyyy")
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_from.setEnabled(False)
        self.date_from.dateChanged.connect(lambda _: self._render())
        row2.addWidget(self.date_from)

        row2.addWidget(QLabel("Đến ngày:"))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("dd/MM/yyyy")
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setEnabled(False)
        self.date_to.dateChanged.connect(lambda _: self._render())
        row2.addWidget(self.date_to)

        row2.addStretch(1)
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
        header = self.tree.header()
        header.setSortIndicatorShown(True)
        header.setSectionsMovable(True)
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.sectionClicked.connect(self._on_header_clicked)
        outer.addWidget(self.tree, 1)

        self.lbl_total = QLabel("")
        self.lbl_total.setStyleSheet("font-weight: 700; font-size: 12pt; color: #2659c4;")
        outer.addWidget(self.lbl_total)

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

    def refresh(self):
        """Nạp lại danh sách giáo viên/khóa + dữ liệu phiếu từ đĩa - gọi khi
        chuyển sang tab này (xem showEvent) và khi tab CẬP NHẬT DỮ LIỆU xử lý
        xong dữ liệu mới, để luôn thấy ngay các phiếu vừa nhập ở tab CẤP XĂNG."""
        self._populate_filters()
        self._render()

    def _on_toggle_all_time(self, checked):
        self.date_from.setEnabled(not checked)
        self.date_to.setEnabled(not checked)
        self._render()

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

        if not self.chk_all_time.isChecked():
            start = self.date_from.date().toPython()
            end = self.date_to.date().toPython()

            def _trong_khoang(r):
                d = _parse_ngay(r.get("ngay_cap", ""))
                return d is not None and start <= d <= end

            recs = [r for r in recs if _trong_khoang(r)]

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
        so_luong_display = r.get("so_luong_text") or _fmt_so(r.get("so_luong"))
        values = [
            r.get("giao_vien", ""), r.get("ma_khoa", ""), r.get("lop", ""), r.get("xe", ""),
            r.get("ngay_cap", ""), r.get("so_phieu", ""), r.get("mat_hang", ""),
            so_luong_display, r.get("phan_loai_so", ""), r.get("ghi_chu", ""), "",
        ]
        item = QTreeWidgetItem(values)
        if parent is None:
            self.tree.addTopLevelItem(item)
        else:
            parent.addChild(item)
        btn = QPushButton("Xóa")
        btn.clicked.connect(lambda _checked=False, rid=r["id"]: self._on_delete(rid))
        self.tree.setItemWidget(item, len(_COLUMN_LABELS) - 1, btn)
        return item

    def _render(self):
        records = self._filtered_records()
        records = self._sorted_records(records)
        group_mode = self.cbo_group.currentText()

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
                group_item.setFirstColumnSpanned(True)
                font = group_item.font(0)
                font.setBold(True)
                group_item.setFont(0, font)
                group_item.setBackground(0, QColor("#eef1f8"))
                self.tree.addTopLevelItem(group_item)
                group_item.setExpanded(True)
                for r in recs:
                    self._add_leaf(r, parent=group_item)

        for col in range(len(_COLUMN_LABELS) - 1):
            self.tree.resizeColumnToContents(col)

        if not records:
            self.lbl_total.setText("Không có phiếu nào khớp điều kiện lọc.")
            return
        tong_lit = sum(r.get("so_luong") or 0 for r in records)
        tong_phieu = sum(r.get("so_phieu_count") or 1 for r in records)
        self.lbl_total.setText(
            f"Tổng cộng: {_fmt_so(tong_lit)} lít ({tong_phieu} phiếu, {len(records)} dòng)"
        )

    def _on_delete(self, record_id):
        confirm = QMessageBox.question(self, "Xác nhận", "Xóa phiếu cấp xăng này?")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        cap_xang_store.delete_record(record_id)
        self._render()

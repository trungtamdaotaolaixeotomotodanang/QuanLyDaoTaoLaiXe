from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCompleter,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

RESULT_COLUMNS = ["Giáo viên", "Mã khóa", "Lớp", "Xe", "SĐT", "Số HV", "Thời gian học", "Tình trạng"]


class TraCuuTab(QWidget):
    """Tab TRA CỨU: tìm theo giáo viên hoặc theo mã khóa, lấy dữ liệu từ data.json."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._build_ui()
        self._load_filters()

    def _build_ui(self):
        outer = QVBoxLayout(self)

        box = QGroupBox("Tra cứu")
        form_layout = QHBoxLayout(box)

        form_layout.addWidget(QLabel("Tên giáo viên:"))
        self.cbo_teacher = QComboBox()
        self.cbo_teacher.setEditable(True)
        self.cbo_teacher.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_teacher.lineEdit().setPlaceholderText("Nhập hoặc chọn tên giáo viên...")
        self.cbo_teacher.currentTextChanged.connect(self._search_by_teacher)
        form_layout.addWidget(self.cbo_teacher, 1)

        form_layout.addSpacing(20)

        form_layout.addWidget(QLabel("Mã khóa:"))
        self.cbo_course = QComboBox()
        self.cbo_course.setEditable(True)
        self.cbo_course.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_course.lineEdit().setPlaceholderText("Nhập hoặc chọn mã khóa...")
        form_layout.addWidget(self.cbo_course, 1)
        btn_course = QPushButton("Tìm theo khóa")
        btn_course.clicked.connect(self._search_by_course)
        form_layout.addWidget(btn_course)

        outer.addWidget(box)

        hang_box = QGroupBox("Tổng hợp số học viên theo hạng")
        hang_layout = QVBoxLayout(hang_box)
        self.tbl_hang = QTableWidget()
        self.tbl_hang.setColumnCount(2)
        self.tbl_hang.setHorizontalHeaderLabels(["Hạng", "Số học viên"])
        self.tbl_hang.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_hang.setMaximumHeight(140)
        self.tbl_hang.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_hang.setSortingEnabled(True)
        hang_layout.addWidget(self.tbl_hang)
        outer.addWidget(hang_box)

        splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget()
        self.table.setColumnCount(len(RESULT_COLUMNS))
        self.table.setHorizontalHeaderLabels(RESULT_COLUMNS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.itemSelectionChanged.connect(self._on_row_selected)
        self.table.setSortingEnabled(True)
        splitter.addWidget(self.table)

        detail_widget = QWidget()
        detail_layout = QVBoxLayout(detail_widget)
        self.lbl_detail = QLabel("Chọn một dòng kết quả để xem chi tiết.")
        detail_layout.addWidget(self.lbl_detail)

        self.detail_tabs = QTabWidget()

        self.tbl_hoc_vien = QTableWidget()
        self.tbl_hoc_vien.setColumnCount(4)
        self.tbl_hoc_vien.setHorizontalHeaderLabels(["Stt", "Họ và tên", "Ngày sinh", "Địa chỉ"])
        self.tbl_hoc_vien.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_hoc_vien.horizontalHeader().setStretchLastSection(True)
        self.tbl_hoc_vien.setSortingEnabled(True)
        self.detail_tabs.addTab(self.tbl_hoc_vien, "Danh sách học viên")

        detail_layout.addWidget(self.detail_tabs)
        splitter.addWidget(detail_widget)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)

        outer.addWidget(splitter, 1)

    def _load_filters(self):
        teachers = self.model.all_teachers()
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.addItem("")
        self.cbo_teacher.addItems(teachers)
        self.cbo_teacher.setCurrentIndex(0)
        self.cbo_teacher.blockSignals(False)

        completer = QCompleter(teachers, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        self.cbo_teacher.setCompleter(completer)

        self.cbo_course.clear()
        course_labels = []
        for c in self.model.all_courses():
            label = f"{c['ma_khoa']} ({c['lop']})"
            self.cbo_course.addItem(label, c["ma_khoa"])
            course_labels.append(label)
        course_completer = QCompleter(course_labels, self)
        course_completer.setCaseSensitivity(Qt.CaseInsensitive)
        course_completer.setFilterMode(Qt.MatchContains)
        self.cbo_course.setCompleter(course_completer)

    def refresh(self):
        """Nạp lại danh sách giáo viên/khóa từ model (sau khi model.reload())
        - gọi sau khi tab CẬP NHẬT WEB đăng dữ liệu mới. Chạy lại tìm kiếm
        theo giáo viên đang chọn (nếu có) để bảng kết quả không bị cũ."""
        teacher_text = self.cbo_teacher.currentText().strip()
        self._load_filters()
        if teacher_text:
            idx = self.cbo_teacher.findText(teacher_text)
            if idx >= 0:
                self.cbo_teacher.blockSignals(True)
                self.cbo_teacher.setCurrentIndex(idx)
                self.cbo_teacher.blockSignals(False)
            self._search_by_teacher(teacher_text)
        else:
            self._show_rows([], "Chọn hoặc nhập tên giáo viên để tra cứu.")

    def _search_by_teacher(self, name=None):
        if name is None:
            name = self.cbo_teacher.currentText()
        name = name.strip()
        if not name:
            self._show_rows([], "Chọn hoặc nhập tên giáo viên để tra cứu.")
            return
        rows = self.model.search_by_teacher(name)
        self._show_rows(rows, f"Kết quả tra cứu theo giáo viên '{name}': {len(rows)} khóa")

    def _search_by_course(self):
        ma_khoa = self.cbo_course.currentData()
        if not ma_khoa:
            return
        rows = self.model.search_by_course(ma_khoa)
        self._show_rows(rows, f"Kết quả tra cứu theo khóa '{ma_khoa}': {len(rows)} giáo viên/xe")

    def _show_rows(self, rows, summary_text):
        # Tắt sort trong lúc đổ dữ liệu (tránh dòng bị xáo lúc đang điền từng
        # ô), bật lại sau - mỗi dòng mang theo dict gốc (UserRole ở cột 0) để
        # _on_row_selected lấy đúng chi tiết dù người dùng đã bấm sort cột.
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows))
        for row_idx, r in enumerate(rows):
            thoi_gian = f"{r['tu_ngay']} → {r['den_ngay']}"
            values = [
                r["giao_vien"],
                r["ma_khoa"],
                r["lop"],
                f"{r['xe']} ({r['ca']})" if r.get("ca") else r["xe"],
                r["sdt"],
                str(r["so_hoc_vien"]),
                thoi_gian,
                r["status"],
            ]
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                if col == 0:
                    item.setData(Qt.ItemDataRole.UserRole, r)
                self.table.setItem(row_idx, col, item)
        self.table.setSortingEnabled(True)
        self._render_hang_summary(rows)
        self.lbl_detail.setText(summary_text)
        if rows:
            self.table.selectRow(0)
            # selectRow() không phát tín hiệu nếu dòng 0 đã đang được chọn từ
            # trước (vd. tìm kiếm mới nhưng trạng thái chọn không đổi), nên
            # luôn cập nhật chi tiết trực tiếp để tránh hiển thị dữ liệu cũ.
            self._render_detail(rows[0])
        else:
            self.tbl_hoc_vien.setRowCount(0)

    def _render_hang_summary(self, rows):
        summary = self.model.summarize_by_hang(rows)
        self.tbl_hang.setSortingEnabled(False)
        self.tbl_hang.setRowCount(len(summary))
        for row_idx, (hang, so_hv) in enumerate(summary):
            self.tbl_hang.setItem(row_idx, 0, QTableWidgetItem(hang))
            item_so_hv = QTableWidgetItem(str(so_hv))
            item_so_hv.setTextAlignment(Qt.AlignCenter)
            self.tbl_hang.setItem(row_idx, 1, item_so_hv)
        self.tbl_hang.setSortingEnabled(True)

    def _on_row_selected(self):
        selected = self.table.selectedItems()
        if not selected:
            return
        item0 = self.table.item(selected[0].row(), 0)
        r = item0.data(Qt.ItemDataRole.UserRole) if item0 else None
        if r is not None:
            self._render_detail(r)

    def _render_detail(self, r):
        self._render_hoc_vien(r["group"])

    def _render_hoc_vien(self, group):
        hoc_vien = group.get("hoc_vien", [])
        self.tbl_hoc_vien.setSortingEnabled(False)
        self.tbl_hoc_vien.setRowCount(len(hoc_vien))
        for row, hv in enumerate(hoc_vien):
            values = [str(hv["stt"]), hv["ho_ten"], hv.get("ngay_sinh", ""), hv.get("dia_chi", "")]
            for col, val in enumerate(values):
                self.tbl_hoc_vien.setItem(row, col, QTableWidgetItem(val))
        self.tbl_hoc_vien.setSortingEnabled(True)
        self.tbl_hoc_vien.resizeColumnsToContents()

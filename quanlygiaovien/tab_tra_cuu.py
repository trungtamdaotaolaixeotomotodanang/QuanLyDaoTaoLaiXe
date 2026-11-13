import re
from datetime import datetime

from openpyxl import Workbook
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCompleter,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

RESULT_COLUMNS = ["Giáo viên", "Mã khóa", "Lớp", "Xe", "SĐT", "Số HV", "Thời gian học", "Tình trạng"]

_SHEET_NAME_INVALID_RE = re.compile(r"[\\/*?:\[\]]")


def _safe_sheet_name(name, used):
    """Tên sheet Excel hợp lệ (bỏ ký tự cấm, tối đa 31 ký tự) và KHÔNG trùng
    tên sheet đã tạo trong cùng workbook (openpyxl báo lỗi nếu trùng) - thêm
    hậu tố " (2)", " (3)"... nếu cần."""
    name = _SHEET_NAME_INVALID_RE.sub("_", name or "").strip() or "Sheet"
    name = name[:31]
    base, i = name, 2
    while name in used:
        suffix = f" ({i})"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name)
    return name


def _add_filter_table(ws, table_name, table_id):
    """Biến vùng dữ liệu đã ghi vào ws thành 1 Excel Table - Excel tự thêm
    nút lọc/sort trên từng cột tiêu đề. table_name/table_id phải duy nhất
    trong cả workbook - Table() mặc định id=1 cho mọi bảng, nếu để nguyên
    thì file có nhiều sheet sẽ bị Excel báo lỗi cần "Repair" khi mở (trùng id
    giữa các bảng ở các sheet khác nhau), nên phải tự gán id tăng dần.

    Table() cũng KHÔNG tự điền tableColumns (openpyxl không tự làm khi gọi
    ws.add_table - phải tự tạo, tên phải khớp đúng text ở dòng tiêu đề),
    thiếu phần này cũng khiến Excel báo lỗi cần "Repair" khi mở."""
    if ws.max_row < 2:
        return
    last_col_idx = ws.max_column
    last_col = ws.cell(row=1, column=last_col_idx).column_letter
    ref = f"A1:{last_col}{ws.max_row}"
    table = Table(id=table_id, displayName=table_name, ref=ref)
    table.tableColumns = [
        TableColumn(id=col, name=str(ws.cell(row=1, column=col).value))
        for col in range(1, last_col_idx + 1)
    ]
    table.autoFilter = AutoFilter(ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True,
    )
    ws.add_table(table)

    for col_idx in range(1, ws.max_column + 1):
        letter = ws.cell(row=1, column=col_idx).column_letter
        width = max(len(str(c.value)) if c.value is not None else 0 for c in ws[letter])
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 45)


_REPORT_COLUMNS = [
    "Mã khóa", "Lớp", "Xe", "SĐT", "Từ ngày", "Đến ngày", "Tình trạng",
    "Họ và tên", "Ngày sinh", "Ghi chú", "Kết quả thi", "Ngày tốt nghiệp", "Địa chỉ",
]


class TraCuuTab(QWidget):
    """Tab TRA CỨU: tìm theo giáo viên hoặc theo mã khóa, lấy dữ liệu từ data.json."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._current_rows = []
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

        form_layout.addSpacing(20)
        self.btn_export = QPushButton("Xuất báo cáo...")
        self.btn_export.setToolTip(
            "Xuất toàn bộ dữ liệu (không phụ thuộc kết quả tra cứu hiện tại) - "
            "mỗi giáo viên 1 sheet, gồm đầy đủ học viên, kết quả thi và ngày tốt nghiệp."
        )
        self.btn_export.clicked.connect(self._on_export)
        form_layout.addWidget(self.btn_export)

        outer.addWidget(box)

        hang_box = QGroupBox("Tổng hợp số học viên theo hạng")
        hang_layout = QVBoxLayout(hang_box)
        self.tbl_hang = QTableWidget()
        self.tbl_hang.setColumnCount(3)
        self.tbl_hang.setHorizontalHeaderLabels(["Năm", "Hạng", "Số học viên"])
        self.tbl_hang.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_hang.setMaximumHeight(140)
        self.tbl_hang.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tbl_hang.setSortingEnabled(True)
        hang_layout.addWidget(self.tbl_hang)
        self.lbl_hang_total = QLabel("Tổng cộng: 0 học viên")
        hang_layout.addWidget(self.lbl_hang_total)
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
        self.tbl_hoc_vien.setColumnCount(6)
        self.tbl_hoc_vien.setHorizontalHeaderLabels(
            ["Stt", "Họ và tên", "Ngày sinh", "Địa chỉ", "Kết quả thi", "Ngày tốt nghiệp"]
        )
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
        self._current_rows = rows
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
        summary, total = self.model.summarize_by_hang(rows)
        self.tbl_hang.setSortingEnabled(False)
        self.tbl_hang.setRowCount(len(summary))
        for row_idx, (nam, hang, so_hv) in enumerate(summary):
            self.tbl_hang.setItem(row_idx, 0, QTableWidgetItem(nam))
            self.tbl_hang.setItem(row_idx, 1, QTableWidgetItem(hang))
            item_so_hv = QTableWidgetItem(str(so_hv))
            item_so_hv.setTextAlignment(Qt.AlignCenter)
            self.tbl_hang.setItem(row_idx, 2, item_so_hv)
        self.tbl_hang.setSortingEnabled(True)
        self.lbl_hang_total.setText(f"Tổng cộng: {total} học viên")

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
            values = [
                str(hv["stt"]), hv["ho_ten"], hv.get("ngay_sinh", ""), hv.get("dia_chi", ""),
                hv.get("ket_qua_thi", ""), hv.get("ngay_tot_nghiep", ""),
            ]
            for col, val in enumerate(values):
                self.tbl_hoc_vien.setItem(row, col, QTableWidgetItem(val))
        self.tbl_hoc_vien.setSortingEnabled(True)
        self.tbl_hoc_vien.resizeColumnsToContents()

    # ---------------- xuất báo cáo ----------------
    def _on_export(self):
        default_name = f"BaoCao_GiaoVien_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất báo cáo", default_name, "Excel Files (*.xlsx)")
        if not path:
            return

        try:
            sheet_count = self._write_report(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        if sheet_count == 0:
            QMessageBox.warning(self, "Không có dữ liệu", "Chưa có khóa/học viên nào để xuất.")
            return
        QMessageBox.information(
            self, "Hoàn tất", f"Đã xuất báo cáo ({sheet_count} giáo viên):\n{path}"
        )

    def _write_report(self, path):
        """Xuất TOÀN BỘ dữ liệu (không phụ thuộc kết quả tra cứu hiện tại đang
        hiển thị) - mỗi giáo viên 1 sheet, mỗi dòng là 1 học viên kèm đầy đủ
        thông tin khóa/xe + kết quả thi/ngày tốt nghiệp của học viên đó, để
        xuất ra là dùng được ngay, không cần tra cứu từng người. Mỗi sheet là
        1 Excel Table (có nút lọc/sort sẵn trên từng cột khi mở trong Excel)."""
        wb = Workbook()
        wb.remove(wb.active)
        used_names = set()

        for idx, teacher in enumerate(self.model.all_teachers(), start=1):
            rows = self.model.search_by_teacher(teacher)
            if not rows:
                continue
            rows = sorted(rows, key=lambda r: (r["ma_khoa"], r["xe"]))
            ws = wb.create_sheet(_safe_sheet_name(teacher, used_names))
            ws.append(_REPORT_COLUMNS)
            for r in rows:
                xe = f"{r['xe']} ({r['ca']})" if r.get("ca") else r["xe"]
                course_cols = [
                    r["ma_khoa"], r["lop"], xe, r["sdt"],
                    r["tu_ngay"], r["den_ngay"], r["status"],
                ]
                hoc_vien = r["group"].get("hoc_vien", [])
                if not hoc_vien:
                    ws.append(course_cols + ["", "", "", "", "", ""])
                    continue
                for hv in hoc_vien:
                    ws.append(course_cols + [
                        hv.get("ho_ten", ""), hv.get("ngay_sinh", ""),
                        hv.get("ghi_chu", ""), hv.get("ket_qua_thi", ""),
                        hv.get("ngay_tot_nghiep", ""), hv.get("dia_chi", ""),
                    ])
            _add_filter_table(ws, f"Tbl{idx}", idx)

        wb.save(path)
        return len(used_names)

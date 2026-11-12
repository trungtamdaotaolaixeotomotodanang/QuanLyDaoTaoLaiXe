"""Tab ĐỊNH MỨC XĂNG: mặc định hiện TẤT CẢ giáo viên đang dạy, hoặc chọn/gõ
1 giáo viên để lọc riêng - hiện danh sách khóa ĐANG DẠY, định mức xăng
(lít/học viên) của xe từng khóa (tra theo file "BẢNG ĐỊNH MỨC NL.xlsx" đã
chọn ở nút "Chọn file định mức..." - xem fuel_norm.py), số học viên tính
xăng (KHÔNG tính học viên "Bảo lưu") và tổng số lít xăng của từng khóa +
tổng cộng toàn bộ các khóa đang hiện trong bảng."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QCompleter,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import fuel_norm

RESULT_COLUMNS = [
    "Giáo viên", "Mã khóa", "Lớp", "Xe", "Thời gian học",
    "Số HV tính xăng", "Định mức/HV (lít)", "Tổng (lít)",
]

ALL_TEACHERS_LABEL = "-- Tất cả giáo viên --"


def _fmt_so(value):
    """Hiện số không có số lẻ .0 thừa (vd 112.0 -> "112", 106.5 giữ nguyên)."""
    if value is None:
        return "—"
    return f"{value:g}"


class DinhMucXangTab(QWidget):
    """Tab ĐỊNH MỨC XĂNG: lấy dữ liệu từ DataModel (data.json + dinh_muc_xang.json)."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._build_ui()
        self._load_filters()
        self._update_dinh_muc_file_label()
        self._search_by_teacher(ALL_TEACHERS_LABEL)

    def _update_dinh_muc_file_label(self):
        src = self.model.dinh_muc_source_file
        self.lbl_dinh_muc_file.setText(
            f"Đang dùng: {src}" if src else "Chưa chọn file định mức nhiên liệu nào."
        )

    def _on_pick_dinh_muc_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file định mức nhiên liệu", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            fuel_norm.load_and_cache(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không đọc được file", str(exc))
            return
        self.model.reload()
        self._update_dinh_muc_file_label()
        self._search_by_teacher()
        QMessageBox.information(
            self, "Hoàn tất",
            "Đã đọc xong file định mức. Bấm \"⬆ Đẩy phiếu/định mức lên\" ở "
            "góc trên để máy khác cũng thấy file này.",
        )

    def _build_ui(self):
        outer = QVBoxLayout(self)

        file_box = QGroupBox("File định mức nhiên liệu")
        file_layout = QHBoxLayout(file_box)
        self.lbl_dinh_muc_file = QLabel("")
        self.lbl_dinh_muc_file.setWordWrap(True)
        file_layout.addWidget(self.lbl_dinh_muc_file, 1)
        btn_pick_dinh_muc = QPushButton("Chọn file định mức...")
        btn_pick_dinh_muc.clicked.connect(self._on_pick_dinh_muc_file)
        file_layout.addWidget(btn_pick_dinh_muc)
        outer.addWidget(file_box)

        box = QGroupBox("Tra cứu định mức xăng theo giáo viên")
        form_layout = QHBoxLayout(box)
        form_layout.addWidget(QLabel("Tên giáo viên:"))
        self.cbo_teacher = QComboBox()
        self.cbo_teacher.setEditable(True)
        self.cbo_teacher.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_teacher.lineEdit().setPlaceholderText("Nhập hoặc chọn tên giáo viên...")
        self.cbo_teacher.currentTextChanged.connect(self._search_by_teacher)
        form_layout.addWidget(self.cbo_teacher, 1)
        outer.addWidget(box)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setWordWrap(True)
        outer.addWidget(self.lbl_summary)

        self.table = QTableWidget()
        self.table.setColumnCount(len(RESULT_COLUMNS))
        self.table.setHorizontalHeaderLabels(RESULT_COLUMNS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setSortingEnabled(True)
        outer.addWidget(self.table, 1)

        self.lbl_total = QLabel("")
        self.lbl_total.setStyleSheet("font-weight: 700; font-size: 12pt; color: #2659c4;")
        outer.addWidget(self.lbl_total)

        self.lbl_missing = QLabel("")
        self.lbl_missing.setWordWrap(True)
        self.lbl_missing.setStyleSheet("color: #b3261e;")
        outer.addWidget(self.lbl_missing)

    def _load_filters(self):
        teachers = self.model.all_teachers()
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.clear()
        self.cbo_teacher.addItem(ALL_TEACHERS_LABEL)
        self.cbo_teacher.addItems(teachers)
        self.cbo_teacher.setCurrentIndex(0)
        self.cbo_teacher.blockSignals(False)

        completer = QCompleter(teachers, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        self.cbo_teacher.setCompleter(completer)

    def refresh(self):
        """Nạp lại danh sách giáo viên + tra cứu lại lựa chọn đang chọn (nếu
        có) - gọi sau khi bấm "Lấy dữ liệu mới nhất" ở main.py."""
        teacher_text = self.cbo_teacher.currentText().strip()
        self._load_filters()
        self._update_dinh_muc_file_label()
        if teacher_text and teacher_text != ALL_TEACHERS_LABEL:
            idx = self.cbo_teacher.findText(teacher_text)
            if idx >= 0:
                self.cbo_teacher.blockSignals(True)
                self.cbo_teacher.setCurrentIndex(idx)
                self.cbo_teacher.blockSignals(False)
        self._search_by_teacher(teacher_text or ALL_TEACHERS_LABEL)

    def _search_by_teacher(self, name=None):
        if name is None:
            name = self.cbo_teacher.currentText()
        name = name.strip()
        if not name or name == ALL_TEACHERS_LABEL:
            rows = self.model.fuel_rows_for_teacher(None)
            self._show_rows(rows, f"Tất cả giáo viên đang dạy: {len(rows)} khóa")
            return
        rows = self.model.fuel_rows_for_teacher(name)
        self._show_rows(rows, f"Các khóa giáo viên '{name}' đang dạy: {len(rows)} khóa")

    def _show_rows(self, rows, summary_text):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows))
        for row_idx, r in enumerate(rows):
            thoi_gian = f"{r['tu_ngay']} → {r['den_ngay']}"
            values = [
                r["giao_vien"],
                r["ma_khoa"],
                r["lop"],
                f"{r['xe']} ({r['ca']})" if r.get("ca") else r["xe"],
                thoi_gian,
                str(r["so_hoc_vien"]),
                _fmt_so(r["dinh_muc"]),
                _fmt_so(r["tong"]),
            ]
            for col, val in enumerate(values):
                item = QTableWidgetItem(val)
                if col >= 5:
                    item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row_idx, col, item)
        self.table.setSortingEnabled(True)
        self.lbl_summary.setText(summary_text)

        tong_cong = sum(r["tong"] for r in rows if r["tong"] is not None)
        self.lbl_total.setText(f"Tổng cộng: {_fmt_so(tong_cong)} lít xăng ({len(rows)} khóa đang dạy)")

        thieu = [r for r in rows if r["dinh_muc"] is None]
        if thieu:
            ds = ", ".join(f"{r['xe']} ({r['ma_khoa']})" for r in thieu)
            self.lbl_missing.setText(
                f"Chưa tra được định mức cho {len(thieu)} xe/khóa - cần bổ sung vào "
                f"file định mức nhiên liệu (nút \"Chọn file định mức...\" ở trên): {ds}"
            )
        else:
            self.lbl_missing.setText("")

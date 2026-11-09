from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import attach_files
import publish_data
from style import icon as _icon


class _LichHocDetailDialog(QDialog):
    """Hộp thoại tổng quan toàn bộ khóa - file lịch học khớp với khóa đó
    (nếu có) - để xem nhanh khóa nào đã có/thiếu file, và gán tay (cập
    nhật file đúng, hoặc chọn file cho khóa chưa có) ngay tại đây - không
    phụ thuộc đúng quy ước đặt tên file PDF như so khớp tự động."""

    def __init__(self, lich_hoc_tab, parent=None):
        super().__init__(parent)
        self.lich_hoc_tab = lich_hoc_tab
        self.setWindowTitle("Thông tin chi tiết - Lịch học thực hành")
        self.resize(760, 480)

        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Mã khóa", "Lớp", "File lịch học đang khớp", ""]
        )
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.setObjectName("primary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        self._reload_table()
        self.table.setSortingEnabled(True)

    def _reload_table(self):
        self.table.setSortingEnabled(False)
        courses = self.lich_hoc_tab.model.all_courses()
        self.table.setRowCount(len(courses))
        for row, c in enumerate(courses):
            ma_khoa, lop = c.get("ma_khoa", ""), c.get("lop", "")
            pdf_path = attach_files.resolve_lich_hoc_pdf(
                ma_khoa, lop, self.lich_hoc_tab.lich_hoc_files, self.lich_hoc_tab.lich_hoc_overrides
            )
            self.table.setItem(row, 0, QTableWidgetItem(ma_khoa))
            self.table.setItem(row, 1, QTableWidgetItem(lop))
            ten_file = Path(pdf_path).name if pdf_path else "(chưa có file lịch học khớp)"
            self.table.setItem(row, 2, QTableWidgetItem(ten_file))
            self.table.setCellWidget(row, 3, self._build_action(ma_khoa, pdf_path))
        self.table.setSortingEnabled(True)

    def _build_action(self, ma_khoa, pdf_path):
        widget = QWidget()
        h = QHBoxLayout(widget)
        h.setContentsMargins(4, 2, 4, 2)
        btn = QPushButton("Cập nhật..." if pdf_path else "Chọn file...")
        btn.clicked.connect(lambda: self._on_update(ma_khoa, pdf_path))
        h.addWidget(btn)
        return widget

    def _on_update(self, ma_khoa, old_pdf_path):
        start_dir = str(Path(old_pdf_path).parent) if old_pdf_path else ""
        new_path, _ = QFileDialog.getOpenFileName(
            self, f"Chọn file lịch học - {ma_khoa}", start_dir, "PDF Files (*.pdf)"
        )
        if not new_path:
            return
        overrides = publish_data.load_lich_hoc_overrides()
        overrides[ma_khoa] = new_path
        publish_data.save_lich_hoc_overrides(overrides)
        self.lich_hoc_tab.lich_hoc_overrides = overrides
        self._reload_table()
        self.lich_hoc_tab._on_course_changed(self.lich_hoc_tab.cbo_course.currentIndex())


class LichHocTab(QWidget):
    """Tab LỊCH HỌC: chọn khóa, xem trực tiếp nội dung file PDF "LICH HOC
    THUC HANH ..." của khóa đó (lịch học giờ chỉ còn ở file PDF, không còn
    nhập tay theo từng ngày trong Excel)."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self.lich_hoc_files = publish_data.load_data_sources().get("lich_hoc_files", [])
        self.lich_hoc_overrides = publish_data.load_lich_hoc_overrides()
        self.current_pdf_path = None
        self._build_ui()
        self._load_courses()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        top = QHBoxLayout()
        top.addWidget(QLabel("Chọn khóa:"))
        self.cbo_course = QComboBox()
        self.cbo_course.currentIndexChanged.connect(self._on_course_changed)
        top.addWidget(self.cbo_course, 1)
        self.btn_open_external = QPushButton("Mở file PDF gốc...")
        self.btn_open_external.setEnabled(False)
        self.btn_open_external.clicked.connect(self._open_external)
        top.addWidget(self.btn_open_external)
        self.btn_detail = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Thông tin chi tiết..."
        )
        self.btn_detail.setToolTip(
            "Xem tổng quan mọi khóa và file lịch học khớp với khóa đó - cập nhật "
            "lại file đúng, hoặc chọn file cho khóa chưa có, ngay tại đây."
        )
        self.btn_detail.clicked.connect(self._open_detail_dialog)
        top.addWidget(self.btn_detail)
        layout.addLayout(top)

        self.lbl_summary = QLabel("")
        layout.addWidget(self.lbl_summary)

        self.pdf_doc = QPdfDocument(self)
        self.pdf_view = QPdfView()
        self.pdf_view.setDocument(self.pdf_doc)
        self.pdf_view.setPageMode(QPdfView.PageMode.MultiPage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitToWidth)
        layout.addWidget(self.pdf_view, 1)

    def _load_courses(self):
        self.cbo_course.clear()
        for c in self.model.all_courses():
            label = f"{c['ma_khoa']} ({c['lop']})"
            self.cbo_course.addItem(label, c["ma_khoa"])

    def refresh(self):
        """Nạp lại danh sách file lịch học + khóa (sau khi model.reload())
        - gọi sau khi tab CẬP NHẬT WEB đăng dữ liệu mới. Giữ nguyên khóa
        đang chọn nếu còn tồn tại, luôn vẽ lại để không còn PDF cũ."""
        ma_khoa_truoc = self.cbo_course.currentData()
        self.lich_hoc_files = publish_data.load_data_sources().get("lich_hoc_files", [])
        self.lich_hoc_overrides = publish_data.load_lich_hoc_overrides()
        self._load_courses()
        idx = self.cbo_course.findData(ma_khoa_truoc) if ma_khoa_truoc else -1
        if idx >= 0:
            self.cbo_course.blockSignals(True)
            self.cbo_course.setCurrentIndex(idx)
            self.cbo_course.blockSignals(False)
        self._on_course_changed(self.cbo_course.currentIndex())

    def _on_course_changed(self, _index):
        ma_khoa = self.cbo_course.currentData()
        course = self.model.find_course(ma_khoa)
        if not course:
            self.lbl_summary.setText("")
            self._show_pdf(None)
            return

        self.lbl_summary.setText(
            f"Lớp {course['lop']} | Thời gian: {course['tu_ngay']} → {course['den_ngay']}"
        )
        pdf_path = attach_files.resolve_lich_hoc_pdf(
            ma_khoa, course.get("lop", ""), self.lich_hoc_files, self.lich_hoc_overrides
        )
        self._show_pdf(pdf_path)

    def _show_pdf(self, pdf_path):
        self.current_pdf_path = pdf_path
        self.btn_open_external.setEnabled(pdf_path is not None)
        if pdf_path is None:
            self.pdf_doc.close()
            suffix = " (chưa có file PDF lịch học khớp với lớp này)"
            self.lbl_summary.setText(self.lbl_summary.text() + suffix)
            return
        self.pdf_doc.load(str(pdf_path))

    def _open_external(self):
        if self.current_pdf_path:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.current_pdf_path)))

    def _open_detail_dialog(self):
        dlg = _LichHocDetailDialog(self, self)
        dlg.exec()

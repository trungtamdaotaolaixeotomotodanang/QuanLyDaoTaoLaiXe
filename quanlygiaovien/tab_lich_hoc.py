from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

import attach_files
import publish_data


class LichHocTab(QWidget):
    """Tab LỊCH HỌC: chọn khóa, xem trực tiếp nội dung file PDF "LICH HOC
    THUC HANH ..." của khóa đó (lịch học giờ chỉ còn ở file PDF, không còn
    nhập tay theo từng ngày trong Excel)."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self.lich_hoc_files = publish_data.load_data_sources().get("lich_hoc_files", [])
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
        pdf_path = attach_files.find_lich_hoc_pdf(course.get("lop", ""), self.lich_hoc_files)
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

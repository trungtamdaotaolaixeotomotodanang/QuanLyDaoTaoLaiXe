from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
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
    """Hộp thoại tổng quan toàn bộ khóa - thư mục chứa file lịch học khớp
    với khóa đó (nếu có) - để xem nhanh khóa nào đã có/thiếu file, và mở
    thư mục đó ra ngay để thêm/thay file mới mà không cần dò tay qua
    Explorer."""

    def __init__(self, courses, lich_hoc_files, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Thông tin chi tiết - Lịch học thực hành")
        self.resize(760, 480)

        layout = QVBoxLayout(self)
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Mã khóa", "Lớp", "Thư mục chứa file lịch học", ""]
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

        self.table.setRowCount(len(courses))
        for row, c in enumerate(courses):
            pdf_path = attach_files.find_lich_hoc_pdf(c.get("lop", ""), lich_hoc_files)
            self.table.setItem(row, 0, QTableWidgetItem(c.get("ma_khoa", "")))
            self.table.setItem(row, 1, QTableWidgetItem(c.get("lop", "")))
            folder = str(Path(pdf_path).parent) if pdf_path else "(chưa có file lịch học khớp)"
            self.table.setItem(row, 2, QTableWidgetItem(folder))
            self.table.setCellWidget(row, 3, self._build_action(pdf_path))
        self.table.setSortingEnabled(True)

    def _build_action(self, pdf_path):
        widget = QWidget()
        h = QHBoxLayout(widget)
        h.setContentsMargins(4, 2, 4, 2)
        btn = QPushButton("Mở thư mục")
        if pdf_path:
            btn.clicked.connect(lambda: self._on_open_folder(pdf_path))
        else:
            btn.setEnabled(False)
        h.addWidget(btn)
        return widget

    def _on_open_folder(self, pdf_path):
        folder = Path(pdf_path).parent
        if not folder.exists():
            QMessageBox.warning(
                self, "Không tìm thấy thư mục", f"Thư mục không còn tồn tại:\n{folder}"
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))


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
        self.btn_detail = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Thông tin chi tiết..."
        )
        self.btn_detail.setToolTip(
            "Xem tổng quan mọi khóa và thư mục chứa file lịch học khớp với khóa đó - "
            "mở thư mục ra để thêm/thay file mới nhanh hơn."
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

    def _open_detail_dialog(self):
        dlg = _LichHocDetailDialog(self.model.all_courses(), self.lich_hoc_files, self)
        dlg.exec()

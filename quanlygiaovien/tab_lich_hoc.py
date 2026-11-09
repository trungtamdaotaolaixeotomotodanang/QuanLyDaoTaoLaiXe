from pathlib import Path

from PySide6.QtCore import QUrl, Qt, QThread, Signal
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtPdf import QPdfDocument
from PySide6.QtPdfWidgets import QPdfView
from PySide6.QtWidgets import (
    QComboBox,
    QCompleter,
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


class _ActionWorker(QThread):
    """Chạy 1 hành động ngắn (gán file lịch học - có git push) ở thread
    riêng để không treo UI. fn không nhận tham số, trả về (ok, message)."""

    finished_result = Signal(bool, str)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self):
        try:
            ok, message = self.fn()
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            ok, message = False, str(exc)
        self.finished_result.emit(ok, message)


class _LichHocDetailDialog(QDialog):
    """Hộp thoại tổng quan toàn bộ khóa - file lịch học khớp với khóa đó
    (nếu có) - để xem nhanh khóa nào đã có/thiếu file, và gán tay (cập
    nhật file đúng, hoặc chọn file cho khóa chưa có) ngay tại đây - không
    phụ thuộc đúng quy ước đặt tên file PDF như so khớp tự động."""

    def __init__(self, lich_hoc_tab, parent=None):
        super().__init__(parent)
        self.lich_hoc_tab = lich_hoc_tab
        self._worker = None
        self.setWindowTitle("Thông tin chi tiết - Lịch học thực hành")
        self.resize(760, 480)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        layout = QVBoxLayout(self)
        note = QLabel(
            "Dòng tô đỏ là khóa chưa có file lịch học khớp. Bấm \"Chọn file...\"/\"Cập "
            "nhật...\" để gán - áp dụng ngay cho khung xem PDF ở tab này. Lưu ý: web công "
            "khai chỉ hiện đúng file này sau khi bạn bấm \"Cập nhật && Đăng lên web\" ở tab "
            "CẬP NHẬT WEB."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

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
            item_ma_khoa = QTableWidgetItem(ma_khoa)
            item_lop = QTableWidgetItem(lop)
            ten_file = Path(pdf_path).name if pdf_path else "(chưa có file lịch học khớp)"
            item_file = QTableWidgetItem(ten_file)
            action = self._build_action(ma_khoa, pdf_path)
            if not pdf_path:
                # Khóa chưa có file lịch học khớp - bôi đỏ cả dòng để dễ nhận ra.
                for item in (item_ma_khoa, item_lop, item_file):
                    item.setBackground(QColor("#f8d7da"))
                action.setStyleSheet("background-color: #f8d7da;")
            self.table.setItem(row, 0, item_ma_khoa)
            self.table.setItem(row, 1, item_lop)
            self.table.setItem(row, 2, item_file)
            self.table.setCellWidget(row, 3, action)
        self.table.setSortingEnabled(True)

    def _build_action(self, ma_khoa, pdf_path):
        widget = QWidget()
        widget.setAttribute(Qt.WA_StyledBackground, True)
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
        # Đăng ký file này vào nguồn "lich_hoc_files" của máy này ngay (nếu
        # chưa có) để so khớp được luôn, không phụ thuộc máy này đã "Thêm"
        # đúng file vào tab CẬP NHẬT WEB hay chưa - xem register_lich_hoc_source.
        self.lich_hoc_tab.register_lich_hoc_source(new_path)
        ten_file = Path(new_path).name

        self.setEnabled(False)
        self._worker = _ActionWorker(lambda: publish_data.set_lich_hoc_override(ma_khoa, ten_file))
        self._worker.finished_result.connect(self._on_update_done)
        self._worker.start()

    def _on_update_done(self, ok, message):
        self.setEnabled(True)
        if ok:
            self.lich_hoc_tab.lich_hoc_overrides = publish_data.load_lich_hoc_overrides()
            self._reload_table()
            self.lich_hoc_tab._on_course_changed(self.lich_hoc_tab.cbo_course.currentIndex())
            QMessageBox.information(
                self, "Hoàn tất",
                f"{message}\n\nLưu ý: web công khai chỉ hiện đúng file này sau khi bạn bấm "
                "\"Cập nhật && Đăng lên web\" ở tab CẬP NHẬT WEB.",
            )
        else:
            QMessageBox.warning(self, "Chưa đăng lên web được", message)


class LichHocTab(QWidget):
    """Tab LỊCH HỌC: chọn khóa, xem trực tiếp nội dung file PDF "LICH HOC
    THUC HANH ..." của khóa đó (lịch học giờ chỉ còn ở file PDF, không còn
    nhập tay theo từng ngày trong Excel)."""

    # Phát khi _LichHocDetailDialog đăng ký thêm 1 file nguồn lịch học mới
    # (xem register_lich_hoc_source) - MainWindow nối sang tab CẬP NHẬT WEB
    # để khung chọn file ở đó cũng có file này, không bị ghi đè mất lúc
    # người dùng bấm "Cập nhật && Đăng lên web" sau.
    lich_hoc_source_added = Signal(str)

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
        self.cbo_course.setEditable(True)
        self.cbo_course.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_course.lineEdit().setPlaceholderText("Nhập hoặc chọn mã khóa...")
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
        labels = []
        for c in self.model.all_courses():
            label = f"{c['ma_khoa']} ({c['lop']})"
            self.cbo_course.addItem(label, c["ma_khoa"])
            labels.append(label)
        completer = QCompleter(labels, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.cbo_course.setCompleter(completer)

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

    def register_lich_hoc_source(self, path):
        """Thêm 1 file vào danh sách nguồn lịch học của máy này
        (data_sources.json) nếu chưa có - để file vừa gán tay (xem
        _LichHocDetailDialog._on_update) so khớp được ngay trên máy này,
        không cần mở tab CẬP NHẬT WEB thêm tay trước. Phát
        lich_hoc_source_added để khung chọn file ở tab CẬP NHẬT WEB (đang
        giữ danh sách riêng trong bộ nhớ từ lúc mở app) cũng cập nhật theo,
        không bị ghi đè mất lúc bấm "Cập nhật && Đăng lên web"."""
        if path in self.lich_hoc_files:
            return
        self.lich_hoc_files.append(path)
        sources = publish_data.load_data_sources()
        files = sources.setdefault("lich_hoc_files", [])
        if path not in files:
            files.append(path)
            publish_data.save_data_sources(sources)
        self.lich_hoc_source_added.emit(path)

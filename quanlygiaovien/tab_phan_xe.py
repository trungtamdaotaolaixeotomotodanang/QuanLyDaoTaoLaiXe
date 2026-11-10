from pathlib import Path

from PySide6.QtCore import QUrl, Qt, QThread, Signal
from PySide6.QtGui import QBrush, QColor, QDesktopServices
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
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

import convert_excel_to_json
import data_model
import publish_data
import roster_xml
from style import icon as _icon


class _ActionWorker(QThread):
    """Chạy 1 hành động ngắn (gán ngày học - có git push) ở thread riêng để
    không treo UI. fn không nhận tham số, trả về (ok, message)."""

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


class _PhanXeDetailDialog(QDialog):
    """Hộp thoại tổng quan toàn bộ khóa - ngày bắt đầu/kết thúc, cho sửa
    trực tiếp (khóa Excel chưa có ngày), mở file Excel gốc để sửa, hoặc tải
    lại dữ liệu khóa đó từ Excel sau khi đã sửa."""

    def __init__(self, phan_xe_tab, parent=None):
        super().__init__(parent)
        self.phan_xe_tab = phan_xe_tab
        self._worker = None
        self.setWindowTitle("Thông tin chi tiết - Phân xe")
        self.resize(920, 520)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        self._loading = False
        self._invalid_cells = set()
        self._original_dates = {}

        layout = QVBoxLayout(self)
        note = QLabel(
            "Sửa trực tiếp \"Ngày bắt đầu\"/\"Ngày kết thúc\" (định dạng dd/mm/yyyy) rồi bấm "
            "\"Lưu thay đổi\" ở dưới - chỉ có tác dụng với khóa mà file Excel chưa có ngày "
            "(Excel luôn ưu tiên hơn). Sai định dạng sẽ báo ngay khi nhập (ô tô đỏ đậm, trỏ "
            "vào ô để xem chi tiết). Dòng tô đỏ nhạt là khóa chưa có đủ ngày bắt đầu/kết thúc; "
            "ô tô vàng ở \"Số học viên\"/\"Số xe/nhóm\"/\"Trạng thái dữ liệu\" là khóa có khả "
            "năng file Excel bị lỗi/thiếu dữ liệu (0 học viên, 0 xe/nhóm, hoặc có nhóm thiếu "
            "xe/giáo viên). \"Mở Excel\" để sửa file gốc, \"Tải lại từ Excel\" để đọc lại ngay "
            "sau khi sửa. Lưu ý: \"Lưu thay đổi\" đăng lên web công khai ngay (nếu khóa đang "
            "hiển thị); riêng \"Tải lại từ Excel\" (số học viên, xe/nhóm...) chỉ cập nhật trên "
            "máy này, vẫn cần bấm \"Cập nhật && Đăng lên web\" ở tab CẬP NHẬT WEB để đăng lên web."
        )
        note.setWordWrap(True)
        layout.addWidget(note)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            [
                "Mã khóa", "Lớp", "Ngày bắt đầu", "Ngày kết thúc",
                "Số học viên", "Số xe/nhóm", "Trạng thái dữ liệu", "",
            ]
        )
        self.table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.EditKeyPressed
        )
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        for col in range(7):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table, 1)
        self.table.itemChanged.connect(self._on_item_changed)

        btn_row = QHBoxLayout()
        self.btn_save_all = QPushButton("Lưu thay đổi")
        self.btn_save_all.setObjectName("primary")
        self.btn_save_all.setToolTip(
            "Lưu tất cả ngày bắt đầu/kết thúc vừa sửa ở bảng trên (chỉ các khóa có thay đổi)."
        )
        self.btn_save_all.clicked.connect(self._on_save_all_changes)
        btn_row.addWidget(self.btn_save_all)
        btn_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        self._reload_table()
        self.table.setSortingEnabled(True)

    def _reload_table(self):
        self._loading = True
        self._invalid_cells = set()
        self.table.setSortingEnabled(False)
        courses = self.phan_xe_tab.model.all_courses()
        self.table.setRowCount(len(courses))
        self._original_dates = {
            c.get("ma_khoa", ""): (c.get("tu_ngay", ""), c.get("den_ngay", ""))
            for c in courses
        }
        for row, c in enumerate(courses):
            item_ma_khoa = QTableWidgetItem(c.get("ma_khoa", ""))
            item_ma_khoa.setFlags(item_ma_khoa.flags() & ~Qt.ItemIsEditable)
            item_lop = QTableWidgetItem(c.get("lop", ""))
            item_lop.setFlags(item_lop.flags() & ~Qt.ItemIsEditable)
            item_tu = QTableWidgetItem(c.get("tu_ngay", ""))
            item_den = QTableWidgetItem(c.get("den_ngay", ""))

            issues = self._data_issues(c)
            item_so_hv = QTableWidgetItem(str(c.get("so_hoc_vien", 0)))
            item_so_xe = QTableWidgetItem(str(len(c.get("groups", []))))
            item_trang_thai = QTableWidgetItem("; ".join(issues) if issues else "✓ Ổn")
            for item in (item_so_hv, item_so_xe, item_trang_thai):
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)

            action = self._build_action(c)
            if not c.get("tu_ngay") or not c.get("den_ngay"):
                # Khóa chưa có đủ ngày bắt đầu/kết thúc - bôi đỏ cả dòng.
                for item in (item_ma_khoa, item_lop, item_tu, item_den):
                    item.setBackground(QColor("#f8d7da"))
                action.setStyleSheet("background-color: #f8d7da;")
            if issues:
                # Khả năng file Excel lỗi/thiếu dữ liệu - bôi vàng riêng 3 cột này
                # (tách biệt với màu đỏ thiếu ngày, vì là 2 vấn đề khác nhau).
                for item in (item_so_hv, item_so_xe, item_trang_thai):
                    item.setBackground(QColor("#fff3cd"))

            self.table.setItem(row, 0, item_ma_khoa)
            self.table.setItem(row, 1, item_lop)
            self.table.setItem(row, 2, item_tu)
            self.table.setItem(row, 3, item_den)
            self.table.setItem(row, 4, item_so_hv)
            self.table.setItem(row, 5, item_so_xe)
            self.table.setItem(row, 6, item_trang_thai)
            self.table.setCellWidget(row, 7, action)
        self.table.setSortingEnabled(True)
        self._loading = False

    @staticmethod
    def _data_issues(course):
        """Vài dấu hiệu dễ gặp khi file Excel nguồn bị lỗi/đọc sai sheet -
        không đầy đủ hay chắc chắn là lỗi, chỉ để gợi ý người dùng kiểm tra
        lại file Excel của khóa đó."""
        issues = []
        groups = course.get("groups", [])
        if not groups:
            issues.append("Thiếu xe/nhóm")
        if not course.get("so_hoc_vien"):
            issues.append("Thiếu học viên")
        missing = sum(1 for g in groups if not g.get("xe") or not g.get("giao_vien"))
        if missing:
            issues.append(f"{missing} nhóm thiếu xe/GV")
        return issues

    def _build_action(self, course):
        ma_khoa = course.get("ma_khoa", "")
        source_path = course.get("source_file_path", "")
        widget = QWidget()
        widget.setAttribute(Qt.WA_StyledBackground, True)
        h = QHBoxLayout(widget)
        h.setContentsMargins(4, 2, 4, 2)
        btn_open = QPushButton("Mở Excel")
        btn_open.setToolTip("Mở file Excel \"BIÊN CHẾ XE KHÓA ...\" gốc của khóa này để sửa.")
        btn_open.clicked.connect(lambda: self._on_open_excel(source_path))
        h.addWidget(btn_open)
        btn_reload = QPushButton("Tải lại từ Excel")
        btn_reload.setToolTip(
            "Đọc lại file Excel gốc (sau khi đã sửa) để cập nhật ngay trên máy này."
        )
        btn_reload.clicked.connect(lambda: self._on_reload_from_excel(ma_khoa, source_path))
        h.addWidget(btn_reload)
        return widget

    def _on_item_changed(self, item):
        """Báo ngay lúc nhập nếu "Ngày bắt đầu"/"Ngày kết thúc" sai định dạng
        - tô đỏ đậm + tooltip cho đúng ô đó, không cần chờ bấm "Lưu thay
        đổi" mới biết."""
        if self._loading or item.column() not in (2, 3):
            return
        text = item.text().strip()
        key = (item.row(), item.column())
        if text and not data_model._parse_date(text):
            item.setBackground(QColor("#f1948a"))
            item.setToolTip(f"Sai định dạng ngày (cần dd/mm/yyyy, ví dụ 01/06/2026): '{text}'")
            self._invalid_cells.add(key)
        else:
            item.setBackground(QBrush())
            item.setToolTip("")
            self._invalid_cells.discard(key)

    def _on_save_all_changes(self):
        if self._invalid_cells:
            rows = sorted({row for row, _ in self._invalid_cells})
            ten_khoa = ", ".join(self.table.item(row, 0).text() for row in rows)
            QMessageBox.warning(
                self, "Sai định dạng ngày",
                "Sửa đúng định dạng dd/mm/yyyy ở ô tô đỏ đậm trước khi lưu - khóa: " + ten_khoa,
            )
            return

        changes = {}
        for row in range(self.table.rowCount()):
            ma_khoa = self.table.item(row, 0).text()
            tu_text = self.table.item(row, 2).text().strip()
            den_text = self.table.item(row, 3).text().strip()
            if (tu_text, den_text) != self._original_dates.get(ma_khoa, ("", "")):
                changes[ma_khoa] = (tu_text, den_text)

        if not changes:
            QMessageBox.information(self, "Không có thay đổi", "Chưa có ngày nào được sửa.")
            return

        self.setEnabled(False)
        self._worker = _ActionWorker(lambda: publish_data.set_ngay_hoc_overrides_bulk(changes))
        self._worker.finished_result.connect(
            lambda ok, msg: self._on_save_all_done(ok, msg, changes)
        )
        self._worker.start()

    def _on_save_all_done(self, ok, message, changes):
        self.setEnabled(True)
        if ok:
            for ma_khoa, (tu_ngay, den_ngay) in changes.items():
                course = self.phan_xe_tab.model.find_course(ma_khoa)
                if course:
                    if not course.get("tu_ngay"):
                        course["tu_ngay"] = tu_ngay
                    if not course.get("den_ngay"):
                        course["den_ngay"] = den_ngay
            self._reload_table()
            if self.phan_xe_tab.cbo_course.currentData() in changes:
                self.phan_xe_tab._render_course(self.phan_xe_tab.cbo_course.currentData())
            QMessageBox.information(
                self, "Hoàn tất",
                f"{message}\n\nĐã lưu {len(changes)} khóa và đăng ngay lên web công khai (nếu "
                "khóa đang hiển thị). Khóa chưa từng đăng lên web (ví dụ khóa mới) sẽ hiện "
                "ngày này sau khi bạn bấm \"Cập nhật && Đăng lên web\" ở tab CẬP NHẬT WEB.",
            )
        else:
            QMessageBox.warning(self, "Chưa đăng lên web được", message)

    def _on_open_excel(self, source_path):
        if not source_path or not Path(source_path).exists():
            QMessageBox.warning(
                self, "Không tìm thấy file",
                "Chưa có file Excel nguồn cho khóa này (hoặc file đã bị di chuyển/xóa).",
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(source_path)))

    def _on_reload_from_excel(self, ma_khoa, source_path):
        if not source_path or not Path(source_path).exists():
            QMessageBox.warning(
                self, "Không tìm thấy file",
                "Chưa có file Excel nguồn cho khóa này (hoặc file đã bị di chuyển/xóa).",
            )
            return
        try:
            new_course = convert_excel_to_json.load_course_from_file(source_path)
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            QMessageBox.warning(self, "Lỗi đọc file Excel", str(exc))
            return
        publish_data.replace_course_in_data_json(new_course)
        self.phan_xe_tab.reload_model_from_disk()
        self._reload_table()
        QMessageBox.information(
            self, "Đã tải lại",
            f"Đã tải lại dữ liệu khóa {ma_khoa} từ file Excel (chỉ máy này - vẫn cần bấm "
            "\"Cập nhật && Đăng lên web\" ở tab CẬP NHẬT WEB để đăng thay đổi này lên web).",
        )


class PhanXeTab(QWidget):
    """Tab PHÂN XE: chọn khóa, xem danh sách xe/giáo viên và học viên theo từng xe."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
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
        self.btn_detail = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Xem chi tiết..."
        )
        self.btn_detail.setToolTip(
            "Xem tổng quan ngày bắt đầu/kết thúc mọi khóa - sửa trực tiếp khóa chưa có "
            "ngày, hoặc mở file Excel gốc để sửa rồi tải lại."
        )
        self.btn_detail.clicked.connect(self._open_detail_dialog)
        top.addWidget(self.btn_detail)
        self.lbl_summary = QLabel("")
        top.addWidget(self.lbl_summary, 2)
        layout.addLayout(top)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(7)
        self.tree.setHeaderLabels(
            [
                "Họ và tên / Xe - GV - SĐT",
                "Ngày sinh",
                "Địa chỉ",
                "Ghi chú",
                "Kết quả thi",
                "Ngày tốt nghiệp",
                "Stt",
            ]
        )
        self.tree.setColumnHidden(6, True)
        self.tree.setAlternatingRowColors(True)
        self.tree.setSortingEnabled(True)
        layout.addWidget(self.tree)

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
        """Nạp lại danh sách khóa từ model (sau khi model.reload()) - gọi
        sau khi tab CẬP NHẬT WEB đăng dữ liệu mới. Giữ nguyên khóa đang chọn
        nếu còn tồn tại, luôn vẽ lại để không còn dữ liệu cũ."""
        ma_khoa_truoc = self.cbo_course.currentData()
        self._load_courses()
        idx = self.cbo_course.findData(ma_khoa_truoc) if ma_khoa_truoc else -1
        if idx >= 0:
            self.cbo_course.blockSignals(True)
            self.cbo_course.setCurrentIndex(idx)
            self.cbo_course.blockSignals(False)
        self._render_course(self.cbo_course.currentData())

    def _on_course_changed(self, _index):
        ma_khoa = self.cbo_course.currentData()
        self._render_course(ma_khoa)

    def _render_course(self, ma_khoa):
        # Tắt sort trong lúc dựng lại cây (tránh top_item.addChild() bị ảnh
        # hưởng bởi sort giữa đường), bật lại ở cuối.
        self.tree.setSortingEnabled(False)
        self.tree.clear()
        course = self.model.find_course(ma_khoa)
        if not course:
            self.lbl_summary.setText("")
            self.tree.setSortingEnabled(True)
            return

        self.lbl_summary.setText(
            f"Lớp {course['lop']} | Thời gian: {course['tu_ngay']} → {course['den_ngay']} | "
            f"Tổng học viên: {course['so_hoc_vien']} | Số xe: {len(course['groups'])}"
        )

        for g in course.get("groups", []):
            header_text = f"{g['giao_vien']}  |  Xe {g['xe']}  |  SĐT {g.get('sdt','')}  ({len(g['hoc_vien'])} HV)"
            top_item = QTreeWidgetItem([header_text, "", "", "", "", "", ""])
            top_item.setFirstColumnSpanned(False)
            font = top_item.font(0)
            font.setBold(True)
            top_item.setFont(0, font)
            self.tree.addTopLevelItem(top_item)

            for hv in g.get("hoc_vien", []):
                child = QTreeWidgetItem(
                    [
                        f"  {hv['stt']}. {hv['ho_ten']}",
                        hv.get("ngay_sinh", ""),
                        hv.get("dia_chi", ""),
                        hv.get("ghi_chu", ""),
                        hv.get("ket_qua_thi", ""),
                        hv.get("ngay_tot_nghiep", ""),
                        str(hv.get("stt", "")),
                    ]
                )
                top_item.addChild(child)

            top_item.setExpanded(True)

        self.tree.setSortingEnabled(True)
        for col in range(6):
            self.tree.resizeColumnToContents(col)

    def _open_detail_dialog(self):
        dlg = _PhanXeDetailDialog(self, self)
        dlg.exec()

    def reload_model_from_disk(self):
        """Nạp lại model từ data.json - dùng sau khi _PhanXeDetailDialog
        ghi đè 1 khóa bằng bản đọc lại từ file Excel nguồn (xem
        publish_data.replace_course_in_data_json). Chỉ cục bộ, không ảnh
        hưởng dữ liệu đã đăng lên web."""
        dsgv_file = publish_data.load_data_sources().get("dsgv_file") or ""
        dsgv = roster_xml.parse_dsgv(dsgv_file) if dsgv_file and Path(dsgv_file).exists() else {}
        self.model.reload(dsgv)
        publish_data.apply_ngay_hoc_overrides_fallback(self.model.courses)
        self._load_courses()
        self._render_course(self.cbo_course.currentData())

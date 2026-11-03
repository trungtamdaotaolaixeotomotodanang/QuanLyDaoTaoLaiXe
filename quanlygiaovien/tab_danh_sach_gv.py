"""Tab DANH SÁCH GIÁO VIÊN: duyệt danh sách toàn bộ giáo viên (DSGV), bấm
vào 1 người là xem ngay thẻ (ảnh có mã QR) của người đó, sửa ảnh thẻ (chọn
file hoặc dán từ clipboard) + hạng lái xe dạy ngay tại đây - ảnh/hạng hiện
cập nhật ngay trên bản xem trước - rồi xuất PDF để in, 1 người hoặc tất cả.

Đây là nơi DUY NHẤT quản lý ảnh/hạng/xuất thẻ giáo viên (trước nằm ở tab
CẬP NHẬT WEB, đã chuyển hẳn qua đây cho gọn, khỏi trùng lặp).

Danh sách giáo viên chỉ đọc 1 lần lúc mở app (từ DSGV_*.xml đã chọn ở tab
CẬP NHẬT WEB) - giống các tab chỉ đọc khác, cần khởi động lại app để thấy
DSGV mới.
"""
import tempfile
from io import BytesIO
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QColor, QGuiApplication, QPixmap
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# Màu tô nền cột "Hạng" trong dialog Danh sách chi tiết - để phân biệt rõ đây
# là dữ liệu nhập/cập nhật trong app, không phải lấy từ file XML.
_MAU_COT_HANG = QColor("#fff3b0")

# Các hạng lái xe thường gặp - hiện trong danh sách chọn, nhưng vẫn cho gõ
# tay hạng khác (vd ghép nhiều hạng "B, C1") vì đây là ô combo có thể sửa.
CAC_HANG_LAI_XE = ["B", "C1", "C", "D1", "D2", "CE"]

import detail_dialog
import publish_data
import roster_xml
from style import icon as _icon


def _given_name_key(full_name):
    """Khóa sắp xếp theo TÊN (từ cuối cùng), không theo HỌ - vd "Nguyễn Văn
    An" sắp xếp theo "An"."""
    parts = (full_name or "").split()
    return (parts[-1], full_name) if parts else (full_name, full_name)


def _pil_to_pixmap(img):
    buf = BytesIO()
    img.save(buf, format="PNG")
    pix = QPixmap()
    pix.loadFromData(buf.getvalue(), "PNG")
    return pix


class _ActionWorker(QThread):
    """Chạy 1 hành động ngắn (thêm/xóa ảnh - có git push) ở thread riêng để
    không treo UI. fn không nhận tham số, trả về (ok: bool, message: str)."""

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


class DanhSachGiaoVienTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.photo_worker = None
        self._current_ma_gv = None
        self._teachers = {}

        sources = publish_data.load_data_sources()
        dsgv_path = sources.get("dsgv_file") or ""
        if dsgv_path and Path(dsgv_path).exists():
            self._teachers = roster_xml.parse_dsgv(dsgv_path)

        self._build_ui()
        self._load_list()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(splitter, 1)

        left = QWidget()
        left.setMinimumWidth(240)
        left.setMaximumWidth(360)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_detail_list = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Danh sách chi tiết"
        )
        self.btn_detail_list.clicked.connect(self._on_show_detail_list)
        left_layout.addWidget(self.btn_detail_list)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Tìm kiếm giáo viên theo tên hoặc mã GV...")
        self.txt_search.textChanged.connect(self._on_search_changed)
        left_layout.addWidget(self.txt_search)

        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_select)
        left_layout.addWidget(self.list_widget, 1)

        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)

        self.lbl_card = QLabel("Chọn 1 giáo viên ở danh sách bên trái.")
        self.lbl_card.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_card.setMinimumHeight(320)
        self.lbl_card.setStyleSheet(
            "background:#ffffff; border:1px solid #d0d4db; border-radius:6px; color:#6b7280;"
        )
        right_layout.addWidget(self.lbl_card, 1)

        photo_row = QHBoxLayout()
        self.btn_choose_photo = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileIcon), "Chọn ảnh..."
        )
        self.btn_choose_photo.clicked.connect(self._on_choose_photo)
        photo_row.addWidget(self.btn_choose_photo)
        self.btn_paste_photo = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogOpenButton), "Dán ảnh (Ctrl+V)"
        )
        self.btn_paste_photo.clicked.connect(self._on_paste_photo)
        photo_row.addWidget(self.btn_paste_photo)
        self.btn_remove_photo = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_TrashIcon), "Xóa ảnh"
        )
        self.btn_remove_photo.clicked.connect(self._on_remove_photo)
        photo_row.addWidget(self.btn_remove_photo)
        right_layout.addLayout(photo_row)

        hang_row = QHBoxLayout()
        hang_row.addWidget(QLabel("Dạy lái xe hạng:"))
        self.txt_hang = QComboBox()
        self.txt_hang.setEditable(True)
        self.txt_hang.addItems(CAC_HANG_LAI_XE)
        self.txt_hang.setCurrentIndex(-1)
        self.txt_hang.setPlaceholderText("Chọn hoặc gõ hạng, vd B, C1...")
        self.txt_hang.editTextChanged.connect(self._on_hang_edited_live)
        hang_row.addWidget(self.txt_hang, 1)
        self.btn_save_hang = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Lưu hạng"
        )
        self.btn_save_hang.clicked.connect(self._on_save_hang)
        hang_row.addWidget(self.btn_save_hang)
        right_layout.addLayout(hang_row)

        export_row = QHBoxLayout()
        self.btn_export_one = QPushButton("Xuất thẻ GV này (PDF)")
        self.btn_export_one.clicked.connect(self._on_export_one)
        export_row.addWidget(self.btn_export_one)
        self.btn_export_all = QPushButton("Xuất tất cả giáo viên (PDF)")
        self.btn_export_all.clicked.connect(self._on_export_all)
        export_row.addWidget(self.btn_export_all)
        export_row.addStretch(1)
        right_layout.addLayout(export_row)

        right_layout.addWidget(QLabel("Toàn bộ thông tin từ file DSGV_*.xml:"))
        self.txt_xml_info = QTextEdit()
        self.txt_xml_info.setReadOnly(True)
        self.txt_xml_info.setMaximumHeight(220)
        right_layout.addWidget(self.txt_xml_info)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        self._set_detail_enabled(False)

    def _set_detail_enabled(self, enabled):
        for w in (
            self.btn_choose_photo, self.btn_paste_photo, self.btn_remove_photo,
            self.txt_hang, self.btn_save_hang, self.btn_export_one,
        ):
            w.setEnabled(enabled)

    # ---- danh sách ----
    def _load_list(self, filter_text=""):
        self.list_widget.clear()
        teachers = sorted(self._teachers.items(), key=lambda kv: _given_name_key(kv[1]["ten"]))
        filter_text = (filter_text or "").strip().lower()
        if filter_text:
            teachers = [
                (ma_gv, rec) for ma_gv, rec in teachers
                if filter_text in rec["ten"].lower() or filter_text in ma_gv.lower()
            ]
        for stt, (ma_gv, rec) in enumerate(teachers, start=1):
            item = QListWidgetItem(f"{stt}. {rec['ten']} ({ma_gv})")
            item.setData(Qt.ItemDataRole.UserRole, ma_gv)
            self.list_widget.addItem(item)
        if not self._teachers:
            self.lbl_card.setText(
                "Chưa có danh sách giáo viên.\n"
                "Hãy chọn file DSGV_*.xml ở tab CẬP NHẬT WEB rồi khởi động lại app."
            )
        elif not teachers:
            empty_item = QListWidgetItem("Không tìm thấy giáo viên phù hợp.")
            empty_item.setFlags(Qt.ItemFlag.NoItemFlags)
            self.list_widget.addItem(empty_item)

    def _on_search_changed(self, text):
        self._load_list(text)

    # ---- dialog danh sách chi tiết (toàn bộ GV + toàn bộ thông tin XML) ----
    def _on_show_detail_list(self):
        if not self._teachers:
            QMessageBox.warning(
                self, "Chưa có DS giáo viên",
                "Hãy chọn file DSGV_*.xml ở tab CẬP NHẬT WEB rồi khởi động lại app.",
            )
            return
        teachers = sorted(self._teachers.items(), key=lambda kv: _given_name_key(kv[1]["ten"]))

        columns = []
        seen = set()
        for _, rec in teachers:
            for tag in rec.get("extra", {}):
                if tag not in seen:
                    seen.add(tag)
                    columns.append(tag)
        columns.append("Hạng")
        hang_col = len(columns) - 1

        dialog = QDialog(self)
        dialog.setWindowTitle("Danh sách chi tiết giáo viên")
        dialog.setWindowFlags(
            dialog.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
        )
        dialog.resize(1100, 600)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "Toàn bộ thông tin từ file DSGV_*.xml. Cột \"Hạng\" (tô vàng) là hạng lái "
            "xe dạy đã cập nhật trong app, không nằm trong file XML."
        ))

        filter_row = QHBoxLayout()
        lbl_total = QLabel()
        filter_row.addWidget(lbl_total)
        filter_row.addStretch(1)
        filter_row.addWidget(QLabel("Lọc:"))
        txt_filter = QLineEdit()
        txt_filter.setPlaceholderText("Lọc theo bất kỳ cột...")
        filter_row.addWidget(txt_filter, 1)
        layout.addLayout(filter_row)

        table = QTableWidget()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setRowCount(len(teachers))
        for row, (ma_gv, rec) in enumerate(teachers):
            extra = rec.get("extra", {})
            for col, tag in enumerate(columns[:-1]):
                table.setItem(row, col, QTableWidgetItem(extra.get(tag, "")))
            hang_item = QTableWidgetItem(publish_data.get_hang_lai_xe(ma_gv))
            hang_item.setBackground(_MAU_COT_HANG)
            table.setItem(row, hang_col, hang_item)
        table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)

        col_key = "giao_vien"
        saved_hidden = publish_data.get_detail_dialog_hidden_columns(col_key)
        if saved_hidden is None:
            hidden_columns = detail_dialog.default_hidden_columns(
                columns[:-1], [rec.get("extra", {}) for _, rec in teachers], always_visible=("Hạng",)
            )
        else:
            hidden_columns = set(saved_hidden) & set(columns)
        detail_dialog.apply_column_visibility(table, columns, hidden_columns)

        def _update_total():
            visible = detail_dialog.filter_table_rows(table, txt_filter.text())
            lbl_total.setText(f"Tổng: {visible}/{len(teachers)} giáo viên")

        txt_filter.textChanged.connect(_update_total)
        _update_total()

        def _on_choose_columns():
            new_hidden = detail_dialog.choose_visible_columns(dialog, columns, hidden_columns)
            if new_hidden is None:
                return
            hidden_columns.clear()
            hidden_columns.update(new_hidden)
            detail_dialog.apply_column_visibility(table, columns, hidden_columns)
            publish_data.set_detail_dialog_hidden_columns(col_key, sorted(hidden_columns))

        bottom_row = QHBoxLayout()
        btn_columns = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogListView), "Chọn cột..."
        )
        btn_columns.clicked.connect(_on_choose_columns)
        bottom_row.addWidget(btn_columns)
        btn_export = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel"
        )
        btn_export.clicked.connect(
            lambda: detail_dialog.export_table_to_excel(
                dialog, table, columns, hang_col, "danh_sach_chi_tiet_giao_vien.xlsx", "Giáo viên"
            )
        )
        bottom_row.addWidget(btn_export)
        bottom_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(dialog.accept)
        bottom_row.addWidget(btn_close)
        layout.addLayout(bottom_row)

        dialog.exec()

    def _on_select(self, current, _previous):
        if not current:
            self._current_ma_gv = None
            self._set_detail_enabled(False)
            self.txt_xml_info.clear()
            return
        ma_gv = current.data(Qt.ItemDataRole.UserRole)
        self._current_ma_gv = ma_gv
        self._set_detail_enabled(True)
        self.txt_hang.blockSignals(True)
        self.txt_hang.setCurrentText(publish_data.get_hang_lai_xe(ma_gv))
        self.txt_hang.blockSignals(False)
        self._refresh_preview()
        self._refresh_xml_info(ma_gv)

    # ---- xem trước thẻ ----
    def _refresh_preview(self, hang_override=None):
        ma_gv = self._current_ma_gv
        if not ma_gv:
            return
        rec = self._teachers.get(ma_gv, {})
        try:
            img = publish_data.render_teacher_card_preview(ma_gv, rec, hang_override=hang_override)
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            self.lbl_card.setText(f"Lỗi xem trước thẻ: {exc}")
            return
        pix = _pil_to_pixmap(img)
        scaled = pix.scaled(
            max(self.lbl_card.width() - 24, 100), max(self.lbl_card.height() - 24, 60),
            Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation,
        )
        self.lbl_card.setPixmap(scaled)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._current_ma_gv:
            self._refresh_preview(hang_override=self.txt_hang.currentText())

    # ---- toàn bộ thông tin từ XML ----
    def _refresh_xml_info(self, ma_gv):
        extra = self._teachers.get(ma_gv, {}).get("extra", {})
        lines = [f"{tag}: {value}" for tag, value in extra.items()]
        self.txt_xml_info.setPlainText("\n".join(lines))

    # ---- hạng lái xe dạy ----
    def _on_hang_edited_live(self, text):
        self._refresh_preview(hang_override=text)

    def _on_save_hang(self):
        if not self._current_ma_gv:
            return
        publish_data.set_hang_lai_xe(self._current_ma_gv, self.txt_hang.currentText())
        QMessageBox.information(self, "Đã lưu", "Đã lưu hạng lái xe dạy cho giáo viên này.")
        self._refresh_preview()

    # ---- ảnh thẻ ----
    def _on_choose_photo(self):
        if not self._current_ma_gv:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn ảnh giáo viên", "", "Ảnh (*.png *.jpg *.jpeg)"
        )
        if not path:
            return
        self._apply_photo(path)

    def _on_paste_photo(self):
        if not self._current_ma_gv:
            return
        image = QGuiApplication.clipboard().image()
        if image.isNull():
            QMessageBox.warning(self, "Không có ảnh", "Clipboard hiện không có ảnh nào để dán.")
            return
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp_path = tmp.name
        tmp.close()
        image.save(tmp_path, "PNG")
        self._apply_photo(tmp_path, cleanup=True)

    def _apply_photo(self, path, cleanup=False):
        ma_gv = self._current_ma_gv
        ten = self._teachers.get(ma_gv, {}).get("ten", "")

        def fn():
            try:
                return publish_data.add_teacher_photo(ma_gv, ten, path)
            finally:
                if cleanup:
                    try:
                        Path(path).unlink(missing_ok=True)
                    except OSError:
                        pass

        self._run_photo_action(fn)

    def _on_remove_photo(self):
        if not self._current_ma_gv:
            return
        ma_gv = self._current_ma_gv
        ten = self._teachers.get(ma_gv, {}).get("ten", "")
        confirm = QMessageBox.question(self, "Xác nhận", f"Xóa ảnh của '{ten}' khỏi trang web?")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._run_photo_action(lambda: publish_data.remove_teacher_photo(ma_gv, ten))

    def _run_photo_action(self, fn):
        self._set_detail_enabled(False)
        self.photo_worker = _ActionWorker(fn)
        self.photo_worker.finished_result.connect(self._on_photo_action_done)
        self.photo_worker.start()

    def _on_photo_action_done(self, ok, message):
        self._set_detail_enabled(True)
        self._refresh_preview(hang_override=self.txt_hang.currentText())
        if ok:
            QMessageBox.information(self, "Hoàn tất", message)
        else:
            QMessageBox.warning(self, "Chưa đăng lên web được", message)

    # ---- xuất PDF ----
    def _on_export_one(self):
        if not self._current_ma_gv:
            return
        ma_gv = self._current_ma_gv
        rec = self._teachers.get(ma_gv, {})
        self._export_cards([(ma_gv, rec)], f"the_{ma_gv}.pdf")

    def _on_export_all(self):
        if not self._teachers:
            QMessageBox.warning(
                self, "Chưa có DS giáo viên",
                "Hãy chọn file DSGV_*.xml ở tab CẬP NHẬT WEB rồi khởi động lại app.",
            )
            return
        teachers = sorted(self._teachers.items(), key=lambda kv: _given_name_key(kv[1]["ten"]))
        self._export_cards(teachers, "the_giao_vien.pdf")

    def _export_cards(self, teachers, default_name):
        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file PDF thẻ giáo viên", default_name, "PDF Files (*.pdf)"
        )
        if not path:
            return
        try:
            so_trang = publish_data.export_teacher_cards(teachers, path)
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            QMessageBox.critical(self, "Lỗi", str(exc))
            return
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã xuất {len(teachers)} thẻ giáo viên ({so_trang} trang) vào:\n{path}",
        )

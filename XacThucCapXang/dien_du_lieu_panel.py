"""Panel cho tab "Điền dữ liệu" (xem dien_du_lieu_logic.py) - chọn file/sheet
nguồn (mặc định quanlycapxang/data/BKeHDMST.xlsx, sheet Sheet1), bấm
"Điền dữ liệu" để tự suy luận Giáo viên quản lý/Khóa cho từng dòng dựa vào
docs/cap_xang_data.json: CHẮC CHẮN thì điền, mơ hồ/không khớp thì để trống
(tô CAM) - bấm vào ô trống đó để xem các khả năng khớp (biển số + ngày hóa
đơn của dòng) và chọn 1 trong số đó (xem dien_du_lieu_logic.candidates_for_row/
autofill_for_row). Cột Số phiếu luôn để trống sẵn, gõ tay trực tiếp. Cả 3 cột
mới đều SỬA TAY được ngay trong bảng (kể cả sau khi đã tự điền) - các cột gốc
chỉ xem, không sửa.

Xuất kết quả (gồm mọi chỉnh sửa tay) ra file Excel mới - dòng nào còn thiếu
Giáo viên quản lý hoặc Khóa thì tô đỏ, để biết ngay cần bổ sung thêm."""
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import excel_validate
import dien_du_lieu_logic as logic
from excel_export import ReportWorkbook
from style import icon as _icon

_MAU_CAN_DIEN = QColor("#fff1dc")  # cam nhạt - ô Giáo viên/Khóa còn trống, cần bấm chọn
_MAU_TRANG = QColor("#ffffff")


def _preview_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value)


class _CandidateDialog(QDialog):
    """Hộp thoại chọn 1 trong các cặp (giáo viên, khóa) khớp biển số + ngày
    của 1 dòng - xem ValidateFillPanel._on_result_cell_clicked."""

    def __init__(self, candidates, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chọn Giáo viên / Khóa")
        self.selected = None
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Các khả năng khớp biển số + ngày hóa đơn của dòng này:"))
        self.list_widget = QListWidget()
        for c in candidates:
            text = f"{c['giao_vien']}  —  {c['khoa_lop']}  ({c['tu_ngay']} - {c['den_ngay']})"
            item = QListWidgetItem(text)
            item.setData(Qt.ItemDataRole.UserRole, c)
            self.list_widget.addItem(item)
        self.list_widget.setCurrentRow(0)
        self.list_widget.itemDoubleClicked.connect(lambda _it: self.accept())
        layout.addWidget(self.list_widget)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.resize(480, 280)

    def accept(self):
        item = self.list_widget.currentItem()
        if item is not None:
            self.selected = item.data(Qt.ItemDataRole.UserRole)
        super().accept()


class DienDuLieuPanel(QWidget):
    def __init__(self, model, default_data_dir, default_file, default_sheet, parent=None):
        super().__init__(parent)
        self.model = model
        self.default_data_dir = Path(default_data_dir)
        self.default_file = Path(default_file)
        self.default_sheet = default_sheet
        self._source_headers, self._source_rows, self._source_row_numbers = [], [], []
        self._output_headers = []
        self._candidates_by_row = []
        self._gv_col = self._khoa_col = self._sp_col = None
        self._map_combos = {}
        self._build_ui()
        self._scan_folder()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)

        file_box = QGroupBox(f"Chọn file & sheet nguồn (mặc định: {self.default_file.name} / {self.default_sheet})")
        file_layout = QVBoxLayout(file_box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("File Excel:"))
        self.cbo_file = QComboBox()
        self.cbo_file.setMinimumWidth(360)
        self.cbo_file.currentIndexChanged.connect(self._on_file_changed)
        row1.addWidget(self.cbo_file, 1)
        btn_browse = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogOpenButton), "Duyệt file khác...")
        btn_browse.clicked.connect(self._on_browse_file)
        row1.addWidget(btn_browse)
        file_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Sheet:"))
        self.cbo_sheet = QComboBox()
        self.cbo_sheet.currentIndexChanged.connect(self._on_sheet_changed)
        row2.addWidget(self.cbo_sheet, 1)
        row2.addSpacing(16)
        row2.addWidget(QLabel("Dòng tiêu đề:"))
        self.spin_header = QSpinBox()
        self.spin_header.setRange(1, 1000)
        self.spin_header.valueChanged.connect(self._reload_source)
        row2.addWidget(self.spin_header)
        file_layout.addLayout(row2)

        row3 = QHBoxLayout()
        row3.addWidget(QLabel("Cột Ngày HĐ:"))
        self.cbo_ngay = QComboBox()
        self.cbo_ngay.setMinimumWidth(160)
        row3.addWidget(self.cbo_ngay)
        row3.addSpacing(16)
        row3.addWidget(QLabel("Cột Biển số (BSX):"))
        self.cbo_bsx = QComboBox()
        self.cbo_bsx.setMinimumWidth(160)
        row3.addWidget(self.cbo_bsx)
        row3.addStretch(1)
        file_layout.addLayout(row3)

        self.lbl_info = QLabel("Chưa chọn file nào.")
        self.lbl_info.setWordWrap(True)
        file_layout.addWidget(self.lbl_info)

        outer.addWidget(file_box)

        action_row = QHBoxLayout()
        self.btn_fill = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogApplyButton), "Điền dữ liệu")
        self.btn_fill.clicked.connect(self._on_fill)
        action_row.addWidget(self.btn_fill)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel...")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._on_export)
        action_row.addWidget(self.btn_export)
        action_row.addStretch(1)
        outer.addLayout(action_row)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setStyleSheet("font-weight: 700; font-size: 11pt;")
        self.lbl_summary.setWordWrap(True)
        outer.addWidget(self.lbl_summary)

        self.tbl_result = QTableWidget()
        self.tbl_result.setAlternatingRowColors(True)
        self.tbl_result.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_result.cellClicked.connect(self._on_result_cell_clicked)
        self.tbl_result.cellChanged.connect(self._on_result_cell_changed)
        self.tbl_result.setToolTip(
            "3 cột cuối (Giáo viên quản lý/Khóa/Số phiếu) gõ tay được trực tiếp. Ô CAM = chưa "
            "xác định được - bấm vào ô đó để xem các khả năng khớp và chọn."
        )
        outer.addWidget(self.tbl_result, 1)

    # ---------------- chọn file/sheet ----------------
    def _scan_folder(self):
        self.cbo_file.blockSignals(True)
        self.cbo_file.clear()
        files = []
        if self.default_data_dir.exists():
            files = sorted(
                [p for p in self.default_data_dir.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm")],
                key=lambda p: p.name.lower(),
            )
        for p in files:
            self.cbo_file.addItem(p.name, str(p))
        if not files:
            self.lbl_info.setText(
                f"Không tìm thấy file Excel nào trong:\n{self.default_data_dir}\n"
                "Bấm \"Duyệt file khác...\" để chọn file ở nơi khác."
            )
        idx = self.cbo_file.findData(str(self.default_file))
        self.cbo_file.blockSignals(False)
        target_idx = idx if idx >= 0 else self.cbo_file.currentIndex()
        if target_idx == self.cbo_file.currentIndex():
            # setCurrentIndex() không phát tín hiệu nếu chỉ số không đổi (vd
            # file mặc định tình cờ đứng đầu danh sách sau khi sắp xếp) - gọi
            # thẳng handler để vẫn nạp sheet/cột, không để trống combo.
            self._on_file_changed(target_idx)
        else:
            self.cbo_file.setCurrentIndex(target_idx)

    def _on_browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel", str(self.default_data_dir), "Excel Files (*.xlsx *.xlsm)"
        )
        if not path:
            return
        idx = self.cbo_file.findData(path)
        if idx < 0:
            self.cbo_file.insertItem(0, Path(path).name, path)
            idx = 0
        self.cbo_file.setCurrentIndex(idx)

    def _on_file_changed(self, _index):
        path = self.cbo_file.currentData()
        self.cbo_sheet.blockSignals(True)
        self.cbo_sheet.clear()
        if path:
            try:
                sheets = excel_validate.list_sheets(path)
            except Exception as exc:
                self.cbo_sheet.blockSignals(False)
                self.lbl_info.setText(f"Không đọc được file: {exc}")
                self._clear_result()
                return
            self.cbo_sheet.addItems(sheets)
        self.cbo_sheet.blockSignals(False)
        default_idx = self.cbo_sheet.findText(self.default_sheet)
        start_idx = default_idx if default_idx >= 0 else self.cbo_sheet.currentIndex()
        if start_idx == self.cbo_sheet.currentIndex():
            self._on_sheet_changed(start_idx)
        else:
            self.cbo_sheet.setCurrentIndex(start_idx)

    def _on_sheet_changed(self, _index):
        path, sheet = self.cbo_file.currentData(), self.cbo_sheet.currentText()
        if not path or not sheet:
            self._clear_result()
            return
        try:
            guess = excel_validate.guess_header_row(path, sheet)
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet: {exc}")
            self._clear_result()
            return
        self.spin_header.blockSignals(True)
        self.spin_header.setValue(guess)
        self.spin_header.blockSignals(False)
        self._reload_source()

    def _clear_result(self):
        self._source_headers, self._source_rows, self._source_row_numbers = [], [], []
        self._output_headers = []
        self._candidates_by_row = []
        self.tbl_result.setRowCount(0)
        self.tbl_result.setColumnCount(0)
        self.btn_export.setEnabled(False)
        self.lbl_summary.setText("")

    def _reload_source(self):
        path, sheet = self.cbo_file.currentData(), self.cbo_sheet.currentText()
        if not path or not sheet:
            self._clear_result()
            return
        try:
            self._source_headers, rows, row_numbers = excel_validate.read_sheet(
                path, sheet, self.spin_header.value()
            )
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet với dòng tiêu đề này: {exc}")
            self._clear_result()
            return
        self._source_rows, self._source_row_numbers, skipped = excel_validate.filter_substantive_rows(
            rows, row_numbers
        )

        self.lbl_info.setText(
            f"{Path(path).name}  —  sheet \"{sheet}\"  —  {len(self._source_rows)} dòng dữ liệu "
            f"(đã bỏ {skipped} dòng trống/mẫu), {len(self._source_headers)} cột nguồn "
            f"(dòng tiêu đề: {self.spin_header.value()})."
        )
        self._fill_column_combos()
        self._clear_result_table_only()

    def _clear_result_table_only(self):
        self._output_headers = []
        self._candidates_by_row = []
        self.tbl_result.setRowCount(0)
        self.tbl_result.setColumnCount(0)
        self.btn_export.setEnabled(False)
        self.lbl_summary.setText("Bấm \"Điền dữ liệu\" để xử lý.")

    def _fill_column_combos(self):
        guesses = logic.detect_columns(self._source_headers)
        for combo, key in ((self.cbo_ngay, "ngay"), (self.cbo_bsx, "bien_so")):
            combo.blockSignals(True)
            combo.clear()
            combo.addItems(self._source_headers)
            guess_idx = guesses.get(key)
            if guess_idx is not None:
                combo.setCurrentIndex(guess_idx)
            combo.blockSignals(False)

    # ---------------- điền dữ liệu ----------------
    def _on_fill(self):
        if not self._source_headers:
            QMessageBox.information(self, "Chưa có dữ liệu", "Chọn file/sheet có dữ liệu trước.")
            return
        ngay_idx = self.cbo_ngay.currentIndex()
        bsx_idx = self.cbo_bsx.currentIndex()
        if ngay_idx < 0 or bsx_idx < 0:
            QMessageBox.information(self, "Chưa chọn cột", "Chọn cột Ngày HĐ và Biển số trước.")
            return

        self._output_headers = logic.build_output_headers(self._source_headers)
        n_kept = len(self._output_headers) - len(logic.NEW_COLUMNS)
        self._gv_col, self._khoa_col, self._sp_col = n_kept, n_kept + 1, n_kept + 2

        self._candidates_by_row = []
        fill_values = []
        for row in self._source_rows:
            candidates = logic.candidates_for_row(self.model, row[ngay_idx], row[bsx_idx])
            self._candidates_by_row.append(candidates)
            fill_values.append(logic.autofill_for_row(candidates))

        self._build_result_table(n_kept, fill_values)

        so_dong = len(self._source_rows)
        ca_hai = sum(1 for gv, khoa in fill_values if gv and khoa)
        mot_cot = sum(1 for gv, khoa in fill_values if bool(gv) != bool(khoa))
        thieu_ca_hai = sum(1 for gv, khoa in fill_values if not gv and not khoa)
        self.lbl_summary.setText(
            f"{so_dong} dòng  —  {ca_hai} dòng tự điền được cả Giáo viên lẫn Khóa, "
            f"{mot_cot} dòng chỉ điền được 1 cột, {thieu_ca_hai} dòng chưa điền được cột nào "
            "(ô cam - bấm vào để chọn)."
        )
        self.btn_export.setEnabled(so_dong > 0)

    def _build_result_table(self, n_kept, fill_values):
        # Chặn tín hiệu cellChanged trong lúc dựng bảng hàng loạt - tránh gọi
        # _on_result_cell_changed (dành cho người dùng gõ tay) hàng trăm lần
        # vô ích lúc setItem().
        self.tbl_result.blockSignals(True)
        self.tbl_result.setColumnCount(len(self._output_headers))
        self.tbl_result.setHorizontalHeaderLabels(self._output_headers)
        self.tbl_result.setRowCount(len(self._source_rows))

        for r, row in enumerate(self._source_rows):
            for c in range(n_kept):
                item = QTableWidgetItem(_preview_text(row[c]))
                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                self.tbl_result.setItem(r, c, item)

            gv, khoa = fill_values[r]
            self._set_fillable_cell(r, self._gv_col, gv or "")
            self._set_fillable_cell(r, self._khoa_col, khoa or "")
            self._set_fillable_cell(r, self._sp_col, "", can_be_empty_without_color=True)

        self.tbl_result.blockSignals(False)
        self.tbl_result.resizeRowsToContents()

    def _set_fillable_cell(self, row, col, text, can_be_empty_without_color=False):
        item = QTableWidgetItem(text)
        item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable)
        if not text and not can_be_empty_without_color:
            item.setBackground(_MAU_CAN_DIEN)
        else:
            item.setBackground(_MAU_TRANG)
        self.tbl_result.setItem(row, col, item)

    def _on_result_cell_changed(self, row, col):
        """Người dùng gõ tay trực tiếp vào ô Giáo viên/Khóa - hết trống thì bỏ
        màu cam, trống lại thì tô cam lại (nhắc còn thiếu)."""
        if col not in (self._gv_col, self._khoa_col):
            return
        item = self.tbl_result.item(row, col)
        if item is None:
            return
        item.setBackground(_MAU_TRANG if item.text().strip() else _MAU_CAN_DIEN)

    # ---------------- bấm ô trống để chọn ----------------
    def _on_result_cell_clicked(self, row, col):
        if col not in (self._gv_col, self._khoa_col):
            return
        item = self.tbl_result.item(row, col)
        if item is None or item.text().strip():
            return  # đã có giá trị - không mở hộp thoại, để sửa tay bình thường
        if row >= len(self._candidates_by_row):
            return
        candidates = self._candidates_by_row[row]
        if not candidates:
            QMessageBox.information(
                self, "Không tìm thấy",
                "Không khớp giáo viên/khóa nào với biển số + ngày hóa đơn của dòng này trong "
                "docs/cap_xang_data.json - kiểm tra lại biển số/ngày, hoặc điền tay."
            )
            return
        dialog = _CandidateDialog(candidates, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.selected:
            c = dialog.selected
            self.tbl_result.item(row, self._gv_col).setText(c["giao_vien"])
            self.tbl_result.item(row, self._khoa_col).setText(c["khoa_lop"])

    # ---------------- xuất Excel ----------------
    def _on_export(self):
        if not self._output_headers:
            return
        default_name = f"DienDuLieu_{self.cbo_sheet.currentText()}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất Excel", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        rows_out, highlight = [], []
        for r in range(self.tbl_result.rowCount()):
            row_vals = [self.tbl_result.item(r, c).text() for c in range(self.tbl_result.columnCount())]
            rows_out.append(row_vals)
            gv_missing = not row_vals[self._gv_col].strip()
            khoa_missing = not row_vals[self._khoa_col].strip()
            highlight.append(gv_missing or khoa_missing)
        try:
            report = ReportWorkbook()
            report.add_sheet("Du lieu da dien", self._output_headers, rows_out, highlight=highlight)
            report.save(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        so_thieu = sum(highlight)
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã xuất {len(rows_out)} dòng ra:\n{path}\n\n"
            f"{so_thieu} dòng tô đỏ còn thiếu Giáo viên quản lý và/hoặc Khóa - cần bổ sung thêm."
        )

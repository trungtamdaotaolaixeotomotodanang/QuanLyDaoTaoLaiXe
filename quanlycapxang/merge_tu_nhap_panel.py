"""Panel chọn 2 sổ (ĐÍCH cần bổ sung cột + NGUỒN đã có người gõ 3 cột Tự
nhập) rồi ghép theo merge_tu_nhap_logic, xem trước kết quả, lưu ra file -
dùng cho merge_tu_nhap.py. Phần chọn file/sheet/dòng tiêu đề lặp lại cấu trúc
_FileSheetPicker ở dưới (giống ValidateSheetPanel trong validate_panel.py)
cho cả 2 sổ, vì mỗi sổ cần set cột khác nhau sau khi đọc (sổ đích: chỉ cột
Mã tra cứu; sổ nguồn: cột Mã tra cứu + 3 cột Tự nhập) nên không dùng lại
nguyên panel đó được."""
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
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
import merge_tu_nhap_logic as logic
import merge_tu_nhap_write as writer
import validate1_logic
from style import icon as _icon

_NO_MAPPING = "-- Không dùng --"
_MAU_DIEN = QColor("#e6f4ea")
_MAU_TRONG = QColor("#f4f5f7")
_MAU_BO_QUA = QColor("#fff1dc")


def _preview_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value)


class _FileSheetPicker(QWidget):
    """1 khối chọn file Excel + sheet + dòng tiêu đề + xem trước 5 dòng đầu.
    Khi đổi file/sheet/dòng tiêu đề xong (đọc được headers/rows) thì phát
    `changed` để panel cha tự dò lại các cột cần dùng."""

    changed = Signal()

    def __init__(self, box_title, default_data_dir, default_file, default_sheet, parent=None):
        super().__init__(parent)
        self.default_data_dir = Path(default_data_dir)
        self.default_file = Path(default_file)
        self.default_sheet = default_sheet
        self.headers, self.rows, self.row_numbers = [], [], []
        self.header_row = 1

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        box = QGroupBox(box_title)
        box_layout = QVBoxLayout(box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("File Excel:"))
        self.cbo_file = QComboBox()
        self.cbo_file.setMinimumWidth(320)
        self.cbo_file.currentIndexChanged.connect(self._on_file_changed)
        row1.addWidget(self.cbo_file, 1)
        btn_browse = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogOpenButton), "Duyệt...")
        btn_browse.clicked.connect(self._on_browse_file)
        row1.addWidget(btn_browse)
        box_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Sheet:"))
        self.cbo_sheet = QComboBox()
        self.cbo_sheet.currentIndexChanged.connect(self._on_sheet_changed)
        row2.addWidget(self.cbo_sheet, 1)
        row2.addSpacing(16)
        row2.addWidget(QLabel("Dòng tiêu đề:"))
        self.spin_header = QSpinBox()
        self.spin_header.setRange(1, 1000)
        self.spin_header.valueChanged.connect(self._reload)
        row2.addWidget(self.spin_header)
        box_layout.addLayout(row2)

        self.lbl_info = QLabel("Chưa chọn file nào.")
        self.lbl_info.setWordWrap(True)
        box_layout.addWidget(self.lbl_info)

        self.tbl_preview = QTableWidget()
        self.tbl_preview.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_preview.setMaximumHeight(120)
        self.tbl_preview.setAlternatingRowColors(True)
        self.tbl_preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        box_layout.addWidget(self.tbl_preview)

        outer.addWidget(box)
        self._scan_folder()

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
                "Bấm \"Duyệt...\" để chọn file ở nơi khác."
            )
        idx = self.cbo_file.findData(str(self.default_file))
        target_idx = idx if idx >= 0 else self.cbo_file.currentIndex()
        self.cbo_file.setCurrentIndex(target_idx)
        self.cbo_file.blockSignals(False)
        # Gọi thẳng (không chỉ dựa vào signal currentIndexChanged) - nếu
        # target_idx trùng currentIndex đang có sẵn (vd file mặc định đứng
        # đầu danh sách do tên xếp theo thứ tự chữ cái) thì setCurrentIndex ở
        # trên KHÔNG phát signal, sheet/preview sẽ không được nạp nếu thiếu
        # dòng này.
        self._on_file_changed(target_idx)

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
                self._clear()
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
            self._clear()
            return
        try:
            guess = excel_validate.guess_header_row(path, sheet)
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet: {exc}")
            self._clear()
            return
        self.spin_header.blockSignals(True)
        self.spin_header.setValue(guess)
        self.spin_header.blockSignals(False)
        self._reload()

    def _clear(self):
        self.headers, self.rows, self.row_numbers = [], [], []
        self.tbl_preview.setRowCount(0)
        self.tbl_preview.setColumnCount(0)
        self.changed.emit()

    def _reload(self):
        path, sheet = self.cbo_file.currentData(), self.cbo_sheet.currentText()
        if not path or not sheet:
            self._clear()
            return
        try:
            self.headers, self.rows, self.row_numbers = excel_validate.read_sheet(
                path, sheet, self.spin_header.value()
            )
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet với dòng tiêu đề này: {exc}")
            self._clear()
            return
        self.header_row = self.spin_header.value()
        self.lbl_info.setText(
            f"{Path(path).name}  —  sheet \"{sheet}\"  —  {len(self.rows)} dòng dữ liệu, "
            f"{len(self.headers)} cột (dòng tiêu đề: {self.header_row})."
        )
        self.tbl_preview.setColumnCount(len(self.headers))
        self.tbl_preview.setHorizontalHeaderLabels(self.headers)
        preview_rows = self.rows[:5]
        self.tbl_preview.setRowCount(len(preview_rows))
        for r, row in enumerate(preview_rows):
            for c, val in enumerate(row):
                self.tbl_preview.setItem(r, c, QTableWidgetItem(_preview_text(val)))
        self.changed.emit()

    @property
    def file_path(self):
        return self.cbo_file.currentData()

    @property
    def sheet_name(self):
        return self.cbo_sheet.currentText()


class MergeTuNhapPanel(QWidget):
    def __init__(self, default_data_dir, default_target_file, default_target_sheet,
                 default_source_file, default_source_sheet, parent=None):
        super().__init__(parent)
        self._fill_results = []
        self._build_ui(
            default_data_dir, default_target_file, default_target_sheet,
            default_source_file, default_source_sheet,
        )
        self.picker_target.changed.connect(self._on_target_changed)
        self.picker_source.changed.connect(self._on_source_changed)
        self._on_target_changed()
        self._on_source_changed()

    def _build_ui(self, default_data_dir, target_file, target_sheet, source_file, source_sheet):
        outer = QVBoxLayout(self)

        pickers_row = QHBoxLayout()
        self.picker_target = _FileSheetPicker(
            "Sổ ĐÍCH cần bổ sung 3 cột", default_data_dir, target_file, target_sheet,
        )
        self.picker_source = _FileSheetPicker(
            "Sổ NGUỒN đã có người gõ 3 cột Tự nhập", default_data_dir, source_file, source_sheet,
        )
        pickers_row.addWidget(self.picker_target, 1)
        pickers_row.addWidget(self.picker_source, 1)
        outer.addLayout(pickers_row)

        map_box = QGroupBox("Các cột dùng để ghép (app tự dò - sửa lại nếu dò sai)")
        map_layout = QHBoxLayout(map_box)
        map_layout.addWidget(QLabel("Mã tra cứu (sổ đích):"))
        self.cbo_target_key = QComboBox()
        self.cbo_target_key.setMinimumWidth(130)
        map_layout.addWidget(self.cbo_target_key)
        map_layout.addSpacing(12)
        map_layout.addWidget(QLabel("Mã tra cứu (sổ nguồn):"))
        self.cbo_source_key = QComboBox()
        self.cbo_source_key.setMinimumWidth(130)
        map_layout.addWidget(self.cbo_source_key)
        map_layout.addSpacing(12)
        map_layout.addWidget(QLabel("GV Quản lý:"))
        self.cbo_source_gv = QComboBox()
        map_layout.addWidget(self.cbo_source_gv)
        map_layout.addWidget(QLabel("Khóa:"))
        self.cbo_source_khoa = QComboBox()
        map_layout.addWidget(self.cbo_source_khoa)
        map_layout.addWidget(QLabel("Số phiếu:"))
        self.cbo_source_sp = QComboBox()
        map_layout.addWidget(self.cbo_source_sp)
        outer.addWidget(map_box)

        action_row = QHBoxLayout()
        self.btn_merge = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogApplyButton), "Ghép cột")
        self.btn_merge.clicked.connect(self._on_merge)
        action_row.addWidget(self.btn_merge)
        self.btn_save = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Lưu file đã bổ sung...")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self._on_save)
        action_row.addWidget(self.btn_save)
        action_row.addStretch(1)
        outer.addLayout(action_row)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setStyleSheet("font-weight: 700; font-size: 11pt;")
        outer.addWidget(self.lbl_summary)

        self.tbl_result = QTableWidget()
        cols = ["Dòng", "Mã tra cứu", "GV Quản lý (mới)", "Khóa (mới)", "Số phiếu (mới)", "Trạng thái"]
        self.tbl_result.setColumnCount(len(cols))
        self.tbl_result.setHorizontalHeaderLabels(cols)
        self.tbl_result.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_result.setAlternatingRowColors(True)
        self.tbl_result.setWordWrap(True)
        header = self.tbl_result.horizontalHeader()
        for c in range(5):
            header.setSectionResizeMode(c, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.Stretch)
        self.tbl_result.setSortingEnabled(True)
        outer.addWidget(self.tbl_result, 1)

    # ------------- dò cột khi đổi file/sheet -------------
    def _fill_combo(self, combo, headers, guess_idx):
        combo.blockSignals(True)
        combo.clear()
        combo.addItem(_NO_MAPPING)
        combo.addItems(headers)
        combo.setCurrentIndex((guess_idx + 1) if guess_idx is not None else 0)
        combo.blockSignals(False)

    def _on_target_changed(self):
        headers = self.picker_target.headers
        guess = logic.detect_key_column(headers)
        self._fill_combo(self.cbo_target_key, headers, guess)
        self._reset_result()

    def _on_source_changed(self):
        headers = self.picker_source.headers
        guess_key = logic.detect_key_column(headers)
        self._fill_combo(self.cbo_source_key, headers, guess_key)
        guess_cols = validate1_logic.detect_columns(headers) if headers else {}
        self._fill_combo(self.cbo_source_gv, headers, guess_cols.get("giao_vien"))
        self._fill_combo(self.cbo_source_khoa, headers, guess_cols.get("khoa"))
        self._fill_combo(self.cbo_source_sp, headers, guess_cols.get("so_phieu"))
        self._reset_result()

    def _reset_result(self):
        self.tbl_result.setRowCount(0)
        self._fill_results = []
        self.btn_save.setEnabled(False)
        self.lbl_summary.setText("")

    @staticmethod
    def _combo_idx(combo):
        idx = combo.currentIndex()
        return (idx - 1) if idx > 0 else None

    # ------------- ghép -------------
    def _on_merge(self):
        if not self.picker_target.headers or not self.picker_source.headers:
            QMessageBox.information(self, "Chưa có dữ liệu", "Chọn đủ file/sheet có dữ liệu cho cả 2 sổ trước.")
            return
        target_key = self._combo_idx(self.cbo_target_key)
        source_key = self._combo_idx(self.cbo_source_key)
        if target_key is None or source_key is None:
            QMessageBox.information(self, "Chưa chọn cột", "Chọn cột \"Mã tra cứu\" cho cả 2 sổ để ghép.")
            return

        new_cols = {
            "giao_vien": self._combo_idx(self.cbo_source_gv),
            "khoa": self._combo_idx(self.cbo_source_khoa),
            "so_phieu": self._combo_idx(self.cbo_source_sp),
        }
        source_check = logic.detect_check_columns(self.picker_source.headers)
        target_check = logic.detect_check_columns(self.picker_target.headers)
        # chỉ dùng được vế kiểm tra chéo nào mà CẢ 2 sổ đều dò ra cột tương ứng
        check_keys = [k for k in source_check if source_check[k] is not None and target_check.get(k) is not None]
        source_check = {k: source_check[k] for k in check_keys}
        target_check = {k: target_check[k] for k in check_keys}

        source_map = logic.build_source_map(self.picker_source.rows, source_key, new_cols, source_check)
        self._fill_results = logic.compute_fill(
            self.picker_target.rows, self.picker_target.row_numbers, target_key, target_check, source_map,
        )
        self._show_results(target_key)

    def _show_results(self, target_key):
        self.tbl_result.setSortingEnabled(False)
        self.tbl_result.setRowCount(len(self._fill_results))
        n_dien, n_trong, n_bo_qua = 0, 0, 0
        key_lookup = {rn: row[target_key] for row, rn in zip(self.picker_target.rows, self.picker_target.row_numbers)}
        for r, item in enumerate(self._fill_results):
            vals = item["values"]
            da_dien = any(v is not None for v in vals.values())
            if da_dien:
                n_dien += 1
                mau = _MAU_DIEN
            elif "lệch" in item["reason"] or "mơ hồ" in item["reason"]:
                n_bo_qua += 1
                mau = _MAU_BO_QUA
            else:
                n_trong += 1
                mau = _MAU_TRONG
            values = [
                str(item["row"]), _preview_text(key_lookup.get(item["row"])),
                _preview_text(vals["giao_vien"]), _preview_text(vals["khoa"]), _preview_text(vals["so_phieu"]),
                item["reason"],
            ]
            for c, val in enumerate(values):
                cell = QTableWidgetItem(val)
                cell.setBackground(mau)
                if c == 0:
                    cell.setData(Qt.ItemDataRole.UserRole, item["row"])
                self.tbl_result.setItem(r, c, cell)
        self.tbl_result.setSortingEnabled(True)
        self.tbl_result.resizeRowsToContents()
        self.lbl_summary.setText(
            f"{len(self._fill_results)} dòng  —  {n_dien} dòng đã điền chắc chắn, "
            f"{n_bo_qua} dòng bỏ qua vì mơ hồ/lệch dữ liệu, {n_trong} dòng để trống (không khớp)."
        )
        self.btn_save.setEnabled(n_dien > 0)

    # ------------- lưu -------------
    def _on_save(self):
        if not self._fill_results:
            return
        target_path = self.picker_target.file_path
        default_path = Path(target_path)
        default_name = default_path.with_name(f"{default_path.stem}_da_bo_sung.xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self, "Lưu file đã bổ sung", str(default_name), "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            writer.write_filled_columns(
                target_path, self.picker_target.sheet_name, self.picker_target.header_row,
                len(self.picker_target.headers), self._fill_results, path,
            )
        except Exception as exc:
            QMessageBox.warning(self, "Không lưu được", str(exc))
            return
        n_dien = sum(1 for it in self._fill_results if any(v is not None for v in it["values"].values()))
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã điền chắc chắn {n_dien}/{len(self._fill_results)} dòng, lưu vào:\n{path}\n\n"
            "Các dòng còn trống (không khớp/mơ hồ) - mở file lên tự nhập tiếp.",
        )

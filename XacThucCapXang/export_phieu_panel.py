"""Panel "Xuất phiếu cấp xăng (JSON)" - kế thừa ValidateSheetPanel (chọn file/
sheet, bảng xem trước, "Xác thực" đối chiếu, bảng lỗi - xem validate_panel.py)
và THÊM: 4 cột chỉ cần chép nguyên (Mặt hàng/Số lượng/Ghi chú/Phân loại sổ,
không cần đối chiếu) + nút "Xuất JSON cho quanlycapxang..." để xuất các dòng
KHÔNG có lỗi thành file JSON đúng schema phieu_cap_xang.json (xem
export_phieu_logic.py), cho quanlycapxang NHẬP vào qua nút "⬇ Nhập phiếu từ
JSON...".

Kế thừa (không sửa ValidateSheetPanel gốc) để tab "Xác thực dữ liệu" (sheet
PETROLIMEX) không bị ảnh hưởng gì - 2 tab dùng chung logic đối chiếu nhưng
mục đích cuối khác nhau (báo cáo Excel vs xuất JSON)."""
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QStyle,
)

import export_phieu_logic
from style import icon as _icon
from validate_panel import ValidateSheetPanel, _NO_MAPPING

_EXTRA_MAP_KEYS = ("mat_hang", "so_luong", "ghi_chu", "phan_loai_so")
_EXTRA_MAP_LABELS = {
    "mat_hang": "Cột Mặt hàng:",
    "so_luong": "Cột Số lượng:",
    "ghi_chu": "Cột Ghi chú:",
    "phan_loai_so": "Cột Phân loại sổ:",
}


class ExportPhieuPanel(ValidateSheetPanel):
    def _build_ui(self):
        super()._build_ui()

        extra_box = QGroupBox("Các cột chỉ cần chép nguyên vào JSON (không đối chiếu)")
        extra_layout = QHBoxLayout(extra_box)
        self._extra_map_combos = {}
        for key in _EXTRA_MAP_KEYS:
            extra_layout.addWidget(QLabel(_EXTRA_MAP_LABELS[key]))
            combo = QComboBox()
            combo.setMinimumWidth(140)
            extra_layout.addWidget(combo, 1)
            self._extra_map_combos[key] = combo
        self.layout().addWidget(extra_box)

        export_row = QHBoxLayout()
        self.btn_export_json = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất JSON cho quanlycapxang..."
        )
        self.btn_export_json.setEnabled(False)
        self.btn_export_json.clicked.connect(self._on_export_json)
        export_row.addWidget(self.btn_export_json)
        export_row.addStretch(1)
        self.layout().addLayout(export_row)

    # ---------------- mở rộng vòng đời chọn file/sheet/xác thực ----------------
    def _fill_mapping_combos(self):
        super()._fill_mapping_combos()
        guesses = export_phieu_logic.detect_extra_columns(self._headers)
        for key in _EXTRA_MAP_KEYS:
            combo = self._extra_map_combos[key]
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(_NO_MAPPING)
            combo.addItems(self._headers)
            guess_idx = guesses.get(key)
            combo.setCurrentIndex((guess_idx + 1) if guess_idx is not None else 0)
            combo.blockSignals(False)

    def _reload_preview(self):
        super()._reload_preview()
        self.btn_export_json.setEnabled(False)

    def _clear_preview(self):
        super()._clear_preview()
        self.btn_export_json.setEnabled(False)

    def _on_validate(self):
        super()._on_validate()
        self.btn_export_json.setEnabled(bool(self._validated_rows))

    def _current_extra_mapping(self):
        mapping = {}
        for key, combo in self._extra_map_combos.items():
            idx = combo.currentIndex()
            mapping[key] = (idx - 1) if idx > 0 else None
        return mapping

    # ---------------- xuất JSON ----------------
    def _on_export_json(self):
        if not self._validated_rows:
            return
        mapping = self._current_mapping()
        extra_mapping = self._current_extra_mapping()
        records, skipped = export_phieu_logic.build_records(
            self._headers, self._validated_rows, self._validated_row_numbers,
            self.model, mapping, extra_mapping, self._current_issues,
        )
        if not records:
            QMessageBox.warning(
                self, "Không có phiếu hợp lệ",
                "Không có dòng nào xuất được - xem lý do bị bỏ qua sau khi bấm lại \"Xác thực\"."
            )
            return

        sheet_name = self.cbo_sheet.currentText()
        default_name = f"PhieuCapXang_{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
        path, _ = QFileDialog.getSaveFileName(
            self, "Xuất JSON cho quanlycapxang", default_name, "JSON Files (*.json)"
        )
        if not path:
            return

        source_file = Path(self.cbo_file.currentData()).name if self.cbo_file.currentData() else ""
        try:
            export_phieu_logic.write_json(records, path, source_file, sheet_name)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return

        msg = f"Đã xuất {len(records)} phiếu hợp lệ ra:\n{path}"
        if skipped:
            msg += f"\n\nĐã bỏ qua {len(skipped)} dòng:\n"
            msg += "\n".join(f"- Dòng {s['row']}: {s['reason']}" for s in skipped[:20])
            if len(skipped) > 20:
                msg += f"\n... và {len(skipped) - 20} dòng khác."
        QMessageBox.information(self, "Hoàn tất", msg)

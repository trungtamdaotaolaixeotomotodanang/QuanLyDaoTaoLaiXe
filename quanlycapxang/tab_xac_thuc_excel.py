"""Tab XÁC THỰC EXCEL: chọn 1 file Excel + 1 sheet bất kỳ (mặc định liệt kê
sẵn các file trong thư mục data/ cạnh app - các sổ cấp nhiên liệu/định mức
gốc) rồi rà lỗi:

- Lỗi/cảnh báo KHÔNG cần biết trước cấu trúc file: ô trống bất thường, dòng
  trùng lặp, lệch kiểu dữ liệu so với phần còn lại của cột, ngày gõ sai định
  dạng, khoảng trắng thừa (xem excel_validate.basic_issues).
- ĐỐI CHIẾU (tùy chọn): chỉ định cột nào là Giáo viên/Khóa-Lớp/Biển số rồi so
  với dữ liệu khóa/giáo viên/xe đang có trong app - báo giáo viên lạ/khóa-lớp
  lạ/biển số lạ, kèm gợi ý sửa (xem excel_validate.cross_reference_issues).

"Dòng tiêu đề" tự đoán khi đổi file/sheet (nhiều sổ có vài dòng tiêu đề lớn/
trống phía trên, không phải lúc nào dòng 1 cũng là tiêu đề cột) - sửa tay
được nếu đoán sai, bảng xem trước cập nhật ngay để kiểm tra."""
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt
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
from excel_export import ReportWorkbook
from runtime_paths import app_dir
from style import icon as _icon

DEFAULT_DATA_DIR = app_dir() / "data"
_NO_MAPPING = "-- Không dùng --"
_MAU_LOI = QColor("#fde7e7")
_MAU_CANH_BAO = QColor("#fff1dc")

ISSUE_COLUMNS = ["Dòng", "Cột", "Mức độ", "Vấn đề", "Cách sửa"]


def _preview_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value)


class XacThucExcelTab(QWidget):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._headers = []
        self._rows = []
        self._row_numbers = []
        self._current_issues = []
        self._build_ui()
        self._scan_default_folder()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)

        file_box = QGroupBox("Chọn file & sheet cần xác thực")
        file_layout = QVBoxLayout(file_box)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("File Excel:"))
        self.cbo_file = QComboBox()
        self.cbo_file.setMinimumWidth(320)
        self.cbo_file.currentIndexChanged.connect(self._on_file_changed)
        row1.addWidget(self.cbo_file, 1)
        btn_browse = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogOpenButton), "Duyệt file khác...")
        btn_browse.clicked.connect(self._on_browse_file)
        row1.addWidget(btn_browse)
        btn_rescan = QPushButton(_icon(self, QStyle.StandardPixmap.SP_BrowserReload), "Làm mới danh sách")
        btn_rescan.setToolTip(f"Quét lại thư mục:\n{DEFAULT_DATA_DIR}")
        btn_rescan.clicked.connect(self._scan_default_folder)
        row1.addWidget(btn_rescan)
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
        self.spin_header.setToolTip(
            "Dòng chứa tên cột (vd \"Giáo viên\", \"Biển số\"...) - app tự đoán khi đổi file/"
            "sheet, sửa tay nếu đoán sai (xem bảng xem trước bên dưới có đúng không)."
        )
        self.spin_header.valueChanged.connect(self._reload_preview)
        row2.addWidget(self.spin_header)
        file_layout.addLayout(row2)

        self.lbl_info = QLabel("Chưa chọn file nào.")
        self.lbl_info.setWordWrap(True)
        file_layout.addWidget(self.lbl_info)

        self.tbl_preview = QTableWidget()
        self.tbl_preview.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_preview.setMaximumHeight(140)
        self.tbl_preview.setAlternatingRowColors(True)
        self.tbl_preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        file_layout.addWidget(self.tbl_preview)

        outer.addWidget(file_box)

        map_box = QGroupBox("Đối chiếu với dữ liệu khóa/giáo viên/xe trong app (tùy chọn)")
        map_layout = QHBoxLayout(map_box)
        map_layout.addWidget(QLabel("Cột Giáo viên:"))
        self.cbo_map_gv = QComboBox()
        map_layout.addWidget(self.cbo_map_gv, 1)
        map_layout.addWidget(QLabel("Cột Khóa học/Lớp:"))
        self.cbo_map_khoa = QComboBox()
        map_layout.addWidget(self.cbo_map_khoa, 1)
        map_layout.addWidget(QLabel("Cột Biển số:"))
        self.cbo_map_xe = QComboBox()
        map_layout.addWidget(self.cbo_map_xe, 1)
        map_box.setToolTip(
            "Để \"-- Không dùng --\" nếu file không có cột đó/không muốn đối chiếu. Có chọn thì "
            "app so giá trị cột đó với danh sách giáo viên/khóa/xe hiện có, báo giá trị lạ."
        )
        outer.addWidget(map_box)

        action_row = QHBoxLayout()
        self.btn_validate = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogApplyButton), "Xác thực")
        self.btn_validate.clicked.connect(self._on_validate)
        action_row.addWidget(self.btn_validate)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất báo cáo lỗi...")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._on_export)
        action_row.addWidget(self.btn_export)
        action_row.addStretch(1)
        outer.addLayout(action_row)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setStyleSheet("font-weight: 700; font-size: 11pt;")
        outer.addWidget(self.lbl_summary)

        self.tbl_issues = QTableWidget()
        self.tbl_issues.setColumnCount(len(ISSUE_COLUMNS))
        self.tbl_issues.setHorizontalHeaderLabels(ISSUE_COLUMNS)
        self.tbl_issues.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_issues.setAlternatingRowColors(True)
        self.tbl_issues.setWordWrap(True)
        self.tbl_issues.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.tbl_issues.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.tbl_issues.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.tbl_issues.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.tbl_issues.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.tbl_issues.setSortingEnabled(True)
        outer.addWidget(self.tbl_issues, 1)

    # ---------------- chọn file/sheet ----------------
    def _scan_default_folder(self):
        current_path = self.cbo_file.currentData()
        self.cbo_file.blockSignals(True)
        self.cbo_file.clear()
        files = []
        if DEFAULT_DATA_DIR.exists():
            files = sorted(
                [p for p in DEFAULT_DATA_DIR.iterdir() if p.suffix.lower() in (".xlsx", ".xlsm")],
                key=lambda p: p.name.lower(),
            )
        for p in files:
            self.cbo_file.addItem(p.name, str(p))
        if not files:
            self.lbl_info.setText(
                f"Không tìm thấy file Excel nào trong:\n{DEFAULT_DATA_DIR}\n"
                "Bấm \"Duyệt file khác...\" để chọn file ở nơi khác."
            )
        restored = False
        if current_path:
            idx = self.cbo_file.findData(current_path)
            if idx >= 0:
                self.cbo_file.setCurrentIndex(idx)
                restored = True
        self.cbo_file.blockSignals(False)
        if not restored:
            self._on_file_changed(self.cbo_file.currentIndex())

    def _on_browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file Excel", str(DEFAULT_DATA_DIR), "Excel Files (*.xlsx *.xlsm)"
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
                self._clear_preview()
                return
            self.cbo_sheet.addItems(sheets)
        self.cbo_sheet.blockSignals(False)
        self._on_sheet_changed(self.cbo_sheet.currentIndex())

    def _on_sheet_changed(self, _index):
        path, sheet = self.cbo_file.currentData(), self.cbo_sheet.currentText()
        if not path or not sheet:
            self._clear_preview()
            return
        try:
            guess = excel_validate.guess_header_row(path, sheet)
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet: {exc}")
            self._clear_preview()
            return
        self.spin_header.blockSignals(True)
        self.spin_header.setValue(guess)
        self.spin_header.blockSignals(False)
        self._reload_preview()

    def _clear_preview(self):
        self._headers, self._rows, self._row_numbers = [], [], []
        self.tbl_preview.setRowCount(0)
        self.tbl_preview.setColumnCount(0)
        self.tbl_issues.setRowCount(0)
        self._current_issues = []
        self.btn_export.setEnabled(False)
        self.lbl_summary.setText("")

    def _reload_preview(self):
        path, sheet = self.cbo_file.currentData(), self.cbo_sheet.currentText()
        if not path or not sheet:
            self._clear_preview()
            return
        try:
            self._headers, self._rows, self._row_numbers = excel_validate.read_sheet(
                path, sheet, self.spin_header.value()
            )
        except Exception as exc:
            self.lbl_info.setText(f"Không đọc được sheet với dòng tiêu đề này: {exc}")
            self._clear_preview()
            return

        self.lbl_info.setText(
            f"{Path(path).name}  —  sheet \"{sheet}\"  —  {len(self._rows)} dòng dữ liệu, "
            f"{len(self._headers)} cột (dòng tiêu đề: {self.spin_header.value()})."
        )
        self._fill_preview_table()
        self._fill_mapping_combos()
        self.tbl_issues.setRowCount(0)
        self._current_issues = []
        self.btn_export.setEnabled(False)
        self.lbl_summary.setText("")

    def _fill_preview_table(self):
        self.tbl_preview.setColumnCount(len(self._headers))
        self.tbl_preview.setHorizontalHeaderLabels(self._headers)
        preview_rows = self._rows[:5]
        self.tbl_preview.setRowCount(len(preview_rows))
        for r, row in enumerate(preview_rows):
            for c, val in enumerate(row):
                self.tbl_preview.setItem(r, c, QTableWidgetItem(_preview_text(val)))

    def _fill_mapping_combos(self):
        guesses = {
            self.cbo_map_gv: excel_validate.guess_column(self._headers, excel_validate.TEACHER_KEYWORDS),
            self.cbo_map_khoa: excel_validate.guess_column(self._headers, excel_validate.COURSE_KEYWORDS),
            self.cbo_map_xe: excel_validate.guess_column(self._headers, excel_validate.PLATE_KEYWORDS),
        }
        for combo, guess_idx in guesses.items():
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(_NO_MAPPING)
            combo.addItems(self._headers)
            combo.setCurrentIndex((guess_idx + 1) if guess_idx is not None else 0)
            combo.blockSignals(False)

    def refresh(self):
        """Nạp lại dữ liệu khóa/giáo viên/xe (self.model đã model.reload() ở
        chỗ gọi) - gọi sau khi bấm "Lấy dữ liệu mới nhất" ở main.py, để đối
        chiếu theo đúng dữ liệu mới nhất."""
        pass  # self.model là tham chiếu sống, không cần làm gì thêm ở đây

    # ---------------- xác thực ----------------
    def _on_validate(self):
        if not self._headers:
            QMessageBox.information(self, "Chưa có dữ liệu", "Chọn file/sheet có dữ liệu trước.")
            return
        rows, row_numbers, skipped = excel_validate.filter_substantive_rows(self._rows, self._row_numbers)
        issues = excel_validate.basic_issues(self._headers, rows, row_numbers)

        mapping = {}
        for key, combo in (
            ("giao_vien", self.cbo_map_gv), ("khoa_lop", self.cbo_map_khoa), ("bien_so", self.cbo_map_xe),
        ):
            idx = combo.currentIndex()
            if idx > 0:
                mapping[key] = idx - 1
        if mapping:
            issues += excel_validate.cross_reference_issues(self._headers, rows, row_numbers, self.model, mapping)

        severity_order = {"loi": 0, "canh_bao": 1}
        issues.sort(key=lambda it: (severity_order.get(it["severity"], 2), it["row"]))
        self._current_issues = issues
        self._fill_issues_table(issues)

        so_loi = sum(1 for it in issues if it["severity"] == "loi")
        so_canh_bao = len(issues) - so_loi
        bo_qua = f" (đã bỏ qua {skipped} dòng trống/mẫu không có dữ liệu thật)" if skipped else ""
        if issues:
            self.lbl_summary.setText(
                f"{len(rows)} dòng dữ liệu thật{bo_qua}  —  {so_loi} lỗi, {so_canh_bao} cảnh báo."
            )
        else:
            self.lbl_summary.setText(f"{len(rows)} dòng dữ liệu thật{bo_qua}  —  không phát hiện vấn đề gì.")
        self.btn_export.setEnabled(bool(issues))

    def _fill_issues_table(self, issues):
        self.tbl_issues.setSortingEnabled(False)
        self.tbl_issues.setRowCount(len(issues))
        for r, it in enumerate(issues):
            mau = _MAU_LOI if it["severity"] == "loi" else _MAU_CANH_BAO
            values = [
                str(it["row"]), it["column"] or "(cả dòng)",
                "Lỗi" if it["severity"] == "loi" else "Cảnh báo",
                it["problem"], it["fix"],
            ]
            for c, val in enumerate(values):
                item = QTableWidgetItem(val)
                item.setBackground(mau)
                if c == 0:
                    item.setData(Qt.ItemDataRole.UserRole, it["row"])
                self.tbl_issues.setItem(r, c, item)
        self.tbl_issues.setSortingEnabled(True)
        self.tbl_issues.resizeRowsToContents()

    # ---------------- xuất báo cáo ----------------
    def _on_export(self):
        if not self._current_issues:
            return
        sheet_name = self.cbo_sheet.currentText()
        default_name = f"XacThuc_{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất báo cáo lỗi", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            report = ReportWorkbook()
            report.add_sheet(
                "Loi va canh bao",
                ISSUE_COLUMNS,
                [
                    [it["row"], it["column"] or "(cả dòng)",
                     "Lỗi" if it["severity"] == "loi" else "Cảnh báo", it["problem"], it["fix"]]
                    for it in self._current_issues
                ],
                highlight=[it["severity"] == "loi" for it in self._current_issues],
            )
            report.save(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        QMessageBox.information(self, "Hoàn tất", f"Đã xuất {len(self._current_issues)} dòng ra:\n{path}")

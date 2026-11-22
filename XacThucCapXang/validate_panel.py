"""Panel dùng cho validate2.py (sheet PETROLIMEX + NLDACAP, mỗi sheet 1 tab,
mỗi tab 1 panel độc lập) - chọn file/sheet, tự dò + cho sửa lại 4 cột cần đối
chiếu (giáo viên/khóa/số phiếu/biển số), bấm Xác thực để chạy
validate1_logic.validate_rows() và xuất báo cáo lỗi ra Excel.

Tách riêng khỏi validate2.py để UI không dồn hết vào 1 file - validate2.py
giờ chỉ còn phần dựng cửa sổ + tab."""
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QBrush, QColor
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

import excel_validate
import validate1_logic
from excel_export import ReportWorkbook
from style import icon as _icon

_NO_MAPPING = "-- Không dùng --"
_MAU_LOI = QColor("#fde7e7")
_MAU_CANH_BAO = QColor("#fff1dc")
# Màu CHỮ cho ô dữ liệu bị cảnh báo/lỗi ngay trong bảng xem trước (khác với
# _MAU_LOI/_MAU_CANH_BAO ở trên là màu NỀN dùng cho bảng danh sách lỗi).
_MAU_CHU_LOI = QBrush(QColor("#c0392b"))
_MAU_CHU_CANH_BAO = QBrush(QColor("#b5651d"))

ISSUE_COLUMNS = ["Dòng", "Cột", "Mức độ", "Vấn đề", "Cách sửa"]
_MAP_KEYS = ("giao_vien", "khoa", "so_phieu", "bien_so", "ngay")
_MAP_LABELS = {
    "giao_vien": "Cột GV Quản lý:",
    "khoa": "Cột Khóa:",
    "so_phieu": "Cột Số phiếu:",
    "bien_so": "Cột Biển số (BSX):",
    "ngay": "Cột Ngày cấp:",
}


def _preview_text(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value)


class ValidateSheetPanel(QWidget):
    """1 panel = 1 file + 1 sheet đang chọn. `model` (CapXangModel) truyền
    vào từ ngoài để nhiều panel (vd 2 tab của validate2.py) dùng chung 1 lần
    đọc docs/cap_xang_data.json, không đọc lại mỗi tab. `default_data_dir`:
    thư mục quét sẵn danh sách file .xlsx/.xlsm (combo File Excel)."""

    # (so_loi, so_canh_bao) - phát mỗi khi self._current_issues thay đổi (xác
    # thực xong, hoặc đổi file/sheet làm mất kết quả cũ) - validate2.py nghe
    # tín hiệu này để hiện số lỗi ngay trên tên tab, không cần người dùng bấm
    # vào từng tab mới biết sheet nào còn lỗi.
    issues_changed = Signal(int, int)

    def __init__(self, model, default_data_dir, default_file, default_sheet, parent=None):
        super().__init__(parent)
        self.model = model
        self.default_data_dir = Path(default_data_dir)
        self.default_file = Path(default_file)
        self.default_sheet = default_sheet
        self._headers, self._rows, self._row_numbers = [], [], []
        self._current_issues = []
        self._validated_rows, self._validated_row_numbers = [], []
        self._row_to_table_row = {}
        self._issue_nav_index = -1
        self._map_combos = {}
        self._build_ui()
        self._scan_folder()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)

        file_box = QGroupBox(f"Chọn file & sheet cần xác thực (mặc định: sheet {self.default_sheet})")
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
        self.spin_header.valueChanged.connect(self._reload_preview)
        row2.addWidget(self.spin_header)
        file_layout.addLayout(row2)

        self.lbl_info = QLabel("Chưa chọn file nào.")
        self.lbl_info.setWordWrap(True)
        file_layout.addWidget(self.lbl_info)

        self.edit_search = QLineEdit()
        self.edit_search.setPlaceholderText(
            "Tìm trong bảng bên dưới (Mã tra cứu, Số HĐ, tên GV, biển số...)..."
        )
        self.edit_search.textChanged.connect(self._on_search_changed)
        file_layout.addWidget(self.edit_search)

        self.tbl_preview = QTableWidget()
        self.tbl_preview.setEditTriggers(QTableWidget.NoEditTriggers)
        self.tbl_preview.setMaximumHeight(320)
        self.tbl_preview.setAlternatingRowColors(True)
        self.tbl_preview.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tbl_preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.tbl_preview.cellClicked.connect(self._on_preview_row_clicked)
        self.tbl_preview.setToolTip(
            "Toàn bộ dữ liệu của sheet - ô có chữ ĐỎ là lỗi, CAM là cảnh báo (sau khi bấm "
            "\"Xác thực\"). Bấm vào 1 dòng để xem chi tiết lỗi/cảnh báo của riêng dòng đó."
        )
        file_layout.addWidget(self.tbl_preview)

        outer.addWidget(file_box)

        map_box = QGroupBox("Các cột gõ tay cần đối chiếu (app tự dò - sửa lại nếu dò sai)")
        map_layout = QHBoxLayout(map_box)
        for key in _MAP_KEYS:
            map_layout.addWidget(QLabel(_MAP_LABELS[key]))
            combo = QComboBox()
            combo.setMinimumWidth(140)
            map_layout.addWidget(combo, 1)
            self._map_combos[key] = combo
        map_box.setToolTip(
            "Đối chiếu bộ (GV Quản lý, Khóa, Biển số) với phân công khóa/giáo viên/xe trong "
            "docs/cap_xang_data.json, cộng Ngày cấp có nằm trong thời gian khóa không - để "
            "\"-- Không dùng --\" nếu file không có cột đó."
        )
        outer.addWidget(map_box)

        action_row = QHBoxLayout()
        self.btn_validate = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogApplyButton), "Xác thực")
        self.btn_validate.clicked.connect(self._on_validate)
        action_row.addWidget(self.btn_validate)
        self.btn_info = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_MessageBoxInformation), "Xác thực như thế nào?"
        )
        self.btn_info.setToolTip("Xem giải thích chi tiết cách app đối chiếu dữ liệu cho sheet đang mở.")
        self.btn_info.clicked.connect(self._on_show_info)
        action_row.addWidget(self.btn_info)
        self.btn_export = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất dữ liệu + báo cáo lỗi...")
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self._on_export)
        action_row.addWidget(self.btn_export)
        action_row.addSpacing(16)
        self.btn_prev_issue = QPushButton(_icon(self, QStyle.StandardPixmap.SP_ArrowBack), "Lỗi trước")
        self.btn_prev_issue.setEnabled(False)
        self.btn_prev_issue.clicked.connect(self._on_prev_issue)
        action_row.addWidget(self.btn_prev_issue)
        self.btn_next_issue = QPushButton(_icon(self, QStyle.StandardPixmap.SP_ArrowForward), "Lỗi sau")
        self.btn_next_issue.setEnabled(False)
        self.btn_next_issue.clicked.connect(self._on_next_issue)
        action_row.addWidget(self.btn_next_issue)
        self.lbl_nav = QLabel("")
        action_row.addWidget(self.lbl_nav)
        action_row.addStretch(1)
        outer.addLayout(action_row)

        summary_row = QHBoxLayout()
        self.lbl_summary = QLabel("")
        self.lbl_summary.setStyleSheet("font-weight: 700; font-size: 11pt;")
        summary_row.addWidget(self.lbl_summary, 1)
        self.lbl_row_filter = QLabel("")
        summary_row.addWidget(self.lbl_row_filter)
        self.btn_clear_filter = QPushButton("Xem tất cả lỗi")
        self.btn_clear_filter.setVisible(False)
        self.btn_clear_filter.clicked.connect(self._on_clear_row_filter)
        summary_row.addWidget(self.btn_clear_filter)
        outer.addLayout(summary_row)

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
        self.tbl_issues.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tbl_issues.cellClicked.connect(self._on_issue_row_clicked)
        outer.addWidget(self.tbl_issues, 1)

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
                self._clear_preview()
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
        self._row_to_table_row = {}
        self.edit_search.clear()
        self.tbl_preview.setRowCount(0)
        self.tbl_preview.setColumnCount(0)
        self.tbl_issues.setRowCount(0)
        self._current_issues = []
        self._validated_rows, self._validated_row_numbers = [], []
        self.btn_export.setEnabled(False)
        self.btn_prev_issue.setEnabled(False)
        self.btn_next_issue.setEnabled(False)
        self.lbl_summary.setText("")
        self._reset_row_filter()
        self._reset_issue_nav()
        self.issues_changed.emit(0, 0)

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
        self._validated_rows, self._validated_row_numbers = [], []
        self.btn_export.setEnabled(False)
        self.btn_prev_issue.setEnabled(False)
        self.btn_next_issue.setEnabled(False)
        self.lbl_summary.setText("")
        self._reset_row_filter()
        self._reset_issue_nav()
        self.issues_changed.emit(0, 0)

    def _fill_preview_table(self):
        self.tbl_preview.setColumnCount(len(self._headers))
        self.tbl_preview.setHorizontalHeaderLabels(self._headers)
        self.tbl_preview.setRowCount(len(self._rows))
        self._row_to_table_row = {}
        for r, (row, rn) in enumerate(zip(self._rows, self._row_numbers)):
            self._row_to_table_row[rn] = r
            self.tbl_preview.setRowHidden(r, False)
            for c, val in enumerate(row):
                self.tbl_preview.setItem(r, c, QTableWidgetItem(_preview_text(val)))
        self.edit_search.blockSignals(True)
        self.edit_search.clear()
        self.edit_search.blockSignals(False)

    def _fill_mapping_combos(self):
        guesses = validate1_logic.detect_columns(self._headers)
        for key in _MAP_KEYS:
            combo = self._map_combos[key]
            combo.blockSignals(True)
            combo.clear()
            combo.addItem(_NO_MAPPING)
            combo.addItems(self._headers)
            guess_idx = guesses.get(key)
            combo.setCurrentIndex((guess_idx + 1) if guess_idx is not None else 0)
            combo.blockSignals(False)

    # ---------------- thông tin cách xác thực ----------------
    def _on_show_info(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Xác thực như thế nào?")
        dialog.resize(640, 560)
        layout = QVBoxLayout(dialog)
        browser = QTextBrowser()
        browser.setHtml(self._build_info_html())
        layout.addWidget(browser)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        buttons.accepted.connect(dialog.accept)
        layout.addWidget(buttons)
        dialog.exec()

    def _build_info_html(self):
        """Giải thích CHUNG cách đối chiếu (áp dụng mọi sheet) + phần RIÊNG
        cho sheet đang mở (tên cột thực tế đang dùng cho từng phần đối chiếu
        - mỗi sheet đặt tên cột khác nhau nên app tự dò khác nhau, xem
        validate1_logic.detect_columns)."""
        sheet = self.cbo_sheet.currentText() or "(chưa chọn sheet)"
        file_name = Path(self.cbo_file.currentData()).name if self.cbo_file.currentData() else "(chưa chọn file)"

        def _mapped_header(key):
            combo = self._map_combos[key]
            text = combo.currentText()
            return text if text and text != _NO_MAPPING else None

        rows_html = []
        for key, mo_ta in (
            ("giao_vien", "Giáo viên quản lý"),
            ("khoa", "Khóa/lớp"),
            ("bien_so", "Biển số (BSX)"),
            ("so_phieu", "Số phiếu"),
            ("ngay", "Ngày cấp"),
        ):
            header = _mapped_header(key)
            cot_text = f"cột <b>&quot;{header}&quot;</b>" if header else "<i>-- không dùng --</i>"
            rows_html.append(f"<li>{mo_ta}: {cot_text}</li>")

        return f"""
        <h3>Đối chiếu với dữ liệu khóa/giáo viên/xe (docs/cap_xang_data.json)</h3>
        <ul>
          <li><b>Giáo viên quản lý</b>: phải có trong danh sách giáo viên hiện tại -
              không có thì báo <b>lỗi</b>, kèm gợi ý tên gần giống nhất.</li>
          <li><b>Khóa/lớp</b>: phải khớp mã khóa (vd "48002K262013") hoặc mã lớp
              (vd "BK0626") đang có - không khớp cái nào thì báo <b>lỗi</b>.</li>
          <li><b>Bộ 3 Giáo viên + Khóa + Biển số</b>: dù từng cột riêng lẻ hợp lệ,
              vẫn phải khớp ĐÚNG 1 phân công đang có (giáo viên đó thật sự dạy xe
              đó trong khóa đó) - khác đi thì báo <b>lỗi</b>, kèm gợi ý giáo viên/
              xe đúng theo phân công của khóa.</li>
          <li><b>Số phiếu</b>: nên chỉ gồm số (có thể nhiều số cách nhau dấu phẩy,
              vd "12,13,14") - có ký tự lạ (vd ký hiệu quyển "1/Q9") thì chỉ
              <b>cảnh báo</b>, vì không chắc là sai (1 số sổ dùng ký hiệu này).</li>
          <li><b>Ngày cấp</b>: nếu đã xác định đúng 1 khóa, ngày phải nằm trong
              khoảng từ ngày - đến ngày của khóa đó - ngoài khoảng chỉ <b>cảnh
              báo</b>, vì ngày cấp lố vài hôm trước/sau ngày khai/bế giảng vẫn có
              thể hợp lệ trong thực tế.</li>
        </ul>
        <p><i>Lỗi</i> = gần như chắc chắn sai, cần sửa. <i>Cảnh báo</i> = nên xem
        lại nhưng có thể vẫn đúng.</p>

        <h3>Mỗi sheet dò cột khác nhau</h3>
        <p>Các sheet không đặt tên cột giống nhau (vd sheet PETROLIMEX ghi
        "GV Quản lý* (Tự nhập)", sheet NLDACAP ghi "GV quản lý* (Tự nhập)" nhưng
        "Khoá học"/"Số Phiếu /Quyển" lại KHÔNG ghi "Tự nhập") - nên app tự dò cột
        nào đóng vai trò gì theo từ khóa trong tiêu đề mỗi khi đổi file/sheet.
        Dò sai thì sửa lại ở khung "Các cột gõ tay cần đối chiếu" phía trên bảng.</p>

        <h3>Đang áp dụng cho: {file_name} — sheet "{sheet}"</h3>
        <ul>
        {''.join(rows_html)}
        </ul>
        """

    # ---------------- xác thực ----------------
    def _current_mapping(self):
        mapping = {}
        for key, combo in self._map_combos.items():
            idx = combo.currentIndex()
            mapping[key] = (idx - 1) if idx > 0 else None
        return mapping

    def _on_validate(self):
        if not self._headers:
            QMessageBox.information(self, "Chưa có dữ liệu", "Chọn file/sheet có dữ liệu trước.")
            return
        mapping = self._current_mapping()
        if mapping.get("giao_vien") is None and mapping.get("khoa") is None:
            QMessageBox.information(
                self, "Chưa chọn cột",
                "Chọn ít nhất cột \"GV Quản lý\" hoặc \"Khóa\" để đối chiếu."
            )
            return

        rows, row_numbers, skipped = excel_validate.filter_substantive_rows(self._rows, self._row_numbers)
        issues = validate1_logic.validate_rows(self._headers, rows, row_numbers, self.model, mapping)
        # Giữ lại để _on_export xuất kèm TOÀN BỘ dữ liệu (không chỉ danh sách
        # lỗi) - đúng những dòng/cột vừa chạy xác thực, để tô đỏ đúng dòng.
        self._validated_rows = rows
        self._validated_row_numbers = row_numbers

        severity_order = {"loi": 0, "canh_bao": 1}
        issues.sort(key=lambda it: (severity_order.get(it["severity"], 2), it["row"]))
        self._current_issues = issues
        self._fill_issues_table(issues)
        self._apply_preview_colors(issues)
        self._reset_row_filter()
        self._reset_issue_nav()

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
        self.btn_prev_issue.setEnabled(bool(issues))
        self.btn_next_issue.setEnabled(bool(issues))
        self.issues_changed.emit(so_loi, so_canh_bao)

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

    def _apply_preview_colors(self, issues):
        """Tô chữ ĐỎ (lỗi)/CAM (cảnh báo) cho đúng Ô bị báo vấn đề trong bảng
        xem trước toàn bộ dữ liệu - issue["column"] có thể là tên 1 cột, 1
        tổ hợp nhiều cột cách nhau " / " (xem validate1_logic - lỗi đối chiếu
        chéo GV+Khóa), hoặc None (lỗi cả dòng, không gắn được vào ô nào, bỏ
        qua ở đây - vẫn hiện đủ trong bảng danh sách lỗi bên dưới)."""
        header_to_col = {h: i for i, h in enumerate(self._headers)}
        for r in range(self.tbl_preview.rowCount()):
            for c in range(self.tbl_preview.columnCount()):
                item = self.tbl_preview.item(r, c)
                if item is not None:
                    item.setData(Qt.ItemDataRole.ForegroundRole, None)
        # Tô cảnh báo trước, lỗi sau - để ô vừa lỗi vừa cảnh báo giữ màu lỗi
        # (nặng hơn) nếu 2 vấn đề cùng rơi vào 1 ô.
        for severity, brush in (("canh_bao", _MAU_CHU_CANH_BAO), ("loi", _MAU_CHU_LOI)):
            for it in issues:
                if it["severity"] != severity or not it.get("column"):
                    continue
                table_row = self._row_to_table_row.get(it["row"])
                if table_row is None:
                    continue
                for part in it["column"].split(" / "):
                    col = header_to_col.get(part.strip())
                    if col is None:
                        continue
                    item = self.tbl_preview.item(table_row, col)
                    if item is not None:
                        item.setForeground(brush)

    def _on_preview_row_clicked(self, row, _col):
        if row < 0 or row >= len(self._row_numbers):
            return
        self._show_row_detail(self._row_numbers[row])

    def _show_row_detail(self, row_number):
        if not self._current_issues:
            self.lbl_row_filter.setText(f"Dòng {row_number}: chưa xác thực - bấm \"Xác thực\" trước.")
            self.btn_clear_filter.setVisible(False)
            return
        matched = [it for it in self._current_issues if it["row"] == row_number]
        self._fill_issues_table(matched)
        if matched:
            so_loi = sum(1 for it in matched if it["severity"] == "loi")
            so_canh_bao = len(matched) - so_loi
            self.lbl_row_filter.setText(f"Dòng {row_number}: {so_loi} lỗi, {so_canh_bao} cảnh báo.")
        else:
            self.lbl_row_filter.setText(f"Dòng {row_number}: không có lỗi/cảnh báo.")
        self.btn_clear_filter.setVisible(True)

    def _on_issue_row_clicked(self, row, _col):
        """Bấm 1 dòng ở bảng lỗi bên dưới -> select + cuộn tới đúng dòng đó
        ở bảng xem trước bên trên, để xem đầy đủ dữ liệu cả dòng (không chỉ
        cột đang báo lỗi) - chiều ngược lại với _on_preview_row_clicked."""
        item = self.tbl_issues.item(row, 0)
        if item is None:
            return
        row_number = item.data(Qt.ItemDataRole.UserRole)
        if row_number is not None:
            self._sync_preview_to_row(row_number)

    def _sync_preview_to_row(self, row_number):
        table_row = self._row_to_table_row.get(row_number)
        if table_row is None:
            return
        if self.tbl_preview.isRowHidden(table_row):
            # Dòng đang bị ô tìm kiếm ẩn đi - bỏ lọc để chắc chắn thấy được,
            # đúng tinh thần "xem đầy đủ dữ liệu" của tính năng này.
            self.edit_search.clear()
        self.tbl_preview.selectRow(table_row)
        preview_item = self.tbl_preview.item(table_row, 0)
        if preview_item is not None:
            self.tbl_preview.scrollToItem(preview_item)

    def _on_clear_row_filter(self):
        self._fill_issues_table(self._current_issues)
        self._reset_row_filter()

    def _reset_row_filter(self):
        self.lbl_row_filter.setText("")
        self.btn_clear_filter.setVisible(False)

    # ---------------- tìm kiếm trong bảng xem trước ----------------
    def _on_search_changed(self, text):
        text = text.strip().lower()
        for r in range(self.tbl_preview.rowCount()):
            if not text:
                self.tbl_preview.setRowHidden(r, False)
                continue
            match = False
            for c in range(self.tbl_preview.columnCount()):
                item = self.tbl_preview.item(r, c)
                if item is not None and text in item.text().lower():
                    match = True
                    break
            self.tbl_preview.setRowHidden(r, not match)

    # ---------------- duyệt qua từng lỗi ----------------
    def _reset_issue_nav(self):
        self._issue_nav_index = -1
        self.lbl_nav.setText("")

    def _on_prev_issue(self):
        self._step_issue(-1)

    def _on_next_issue(self):
        self._step_issue(1)

    def _step_issue(self, direction):
        if not self._current_issues:
            return
        n = len(self._current_issues)
        if self._issue_nav_index < 0:
            self._issue_nav_index = 0 if direction > 0 else n - 1
        else:
            self._issue_nav_index = (self._issue_nav_index + direction) % n
        self._goto_issue(self._issue_nav_index)

    def _goto_issue(self, index):
        # Đang lọc bảng lỗi theo 1 dòng (xem _show_row_detail) thì bỏ lọc
        # trước, để chọn đúng theo chỉ mục trong self._current_issues.
        if self.tbl_issues.rowCount() != len(self._current_issues):
            self._fill_issues_table(self._current_issues)
            self._reset_row_filter()
        it = self._current_issues[index]
        visual_row = self._find_issue_visual_row(it)
        if visual_row is not None:
            self.tbl_issues.selectRow(visual_row)
            item = self.tbl_issues.item(visual_row, 0)
            if item is not None:
                self.tbl_issues.scrollToItem(item)
        self._sync_preview_to_row(it["row"])
        self.lbl_nav.setText(f"Lỗi {index + 1}/{len(self._current_issues)} — dòng {it['row']}")

    def _find_issue_visual_row(self, issue):
        """tbl_issues có thể đang sắp xếp khác thứ tự self._current_issues
        (người dùng bấm vào tiêu đề cột để sắp xếp) - dò theo đúng dòng + nội
        dung vấn đề thay vì dùng thẳng index, để luôn chọn đúng ô."""
        for r in range(self.tbl_issues.rowCount()):
            item = self.tbl_issues.item(r, 0)
            if item is None or item.data(Qt.ItemDataRole.UserRole) != issue["row"]:
                continue
            problem_item = self.tbl_issues.item(r, 3)
            if problem_item is not None and problem_item.text() == issue["problem"]:
                return r
        return None

    # ---------------- xuất báo cáo ----------------
    def _on_export(self):
        if not self._current_issues:
            return
        sheet_name = self.cbo_sheet.currentText()
        default_name = f"Validate_{sheet_name}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất dữ liệu + báo cáo lỗi", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        # Dòng nào có ÍT NHẤT 1 lỗi (không tính cảnh báo) thì tô đỏ ở sheet dữ
        # liệu gốc - đúng yêu cầu "dòng nào sai thì bôi đỏ".
        loi_rows = {it["row"] for it in self._current_issues if it["severity"] == "loi"}
        try:
            report = ReportWorkbook()
            report.add_sheet(
                "Du lieu",
                self._headers,
                self._validated_rows,
                highlight=[rn in loi_rows for rn in self._validated_row_numbers],
            )
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
        QMessageBox.information(
            self, "Hoàn tất",
            f"Đã xuất {len(self._validated_rows)} dòng dữ liệu ({len(loi_rows)} dòng tô đỏ vì có lỗi) + "
            f"{len(self._current_issues)} dòng báo lỗi/cảnh báo ra:\n{path}"
        )

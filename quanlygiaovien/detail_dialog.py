"""Hỗ trợ dùng chung cho dialog "Danh sách chi tiết" (tab DANH SÁCH GIÁO VIÊN
và tab DANH SÁCH XE) - lọc nhanh theo mọi cột, chọn cột ẩn/hiện (nhớ theo
từng loại danh sách) + xuất Excel đúng những dòng/cột đang hiện."""
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

_MAU_VANG = PatternFill(start_color="FFF3B0", end_color="FFF3B0", fill_type="solid")


def default_hidden_columns(columns, extra_dicts, always_visible=()):
    """Cột nào rỗng ở TẤT CẢ các dòng có dữ liệu (dict "extra" không rỗng) thì
    mặc định ẩn - để lọc bớt cột không có dữ liệu như người dùng yêu cầu.
    `always_visible` (vd cột tô vàng, không lấy từ XML) không bao giờ bị ẩn
    bởi mặc định này."""
    has_data = {col: False for col in columns}
    for extra in extra_dicts:
        if not extra:
            continue
        for col in columns:
            if (extra.get(col) or "").strip():
                has_data[col] = True
    return {col for col, found in has_data.items() if not found and col not in always_visible}


def apply_column_visibility(table, columns, hidden_columns):
    for col_idx, col in enumerate(columns):
        table.setColumnHidden(col_idx, col in hidden_columns)


def choose_visible_columns(parent, columns, hidden_columns):
    """Hiện dialog chọn cột nào hiện (checkbox). Trả về set cột bị ẩn mới,
    hoặc None nếu người dùng bấm Hủy."""
    dialog = QDialog(parent)
    dialog.setWindowTitle("Chọn cột hiện")
    dialog.resize(360, 480)
    layout = QVBoxLayout(dialog)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    inner = QWidget()
    inner_layout = QVBoxLayout(inner)
    checks = []
    for col in columns:
        cb = QCheckBox(col)
        cb.setChecked(col not in hidden_columns)
        inner_layout.addWidget(cb)
        checks.append(cb)
    inner_layout.addStretch(1)
    scroll.setWidget(inner)
    layout.addWidget(scroll, 1)

    quick_row = QHBoxLayout()
    btn_all = QPushButton("Chọn tất cả")
    btn_all.clicked.connect(lambda: [cb.setChecked(True) for cb in checks])
    quick_row.addWidget(btn_all)
    btn_none = QPushButton("Bỏ chọn tất cả")
    btn_none.clicked.connect(lambda: [cb.setChecked(False) for cb in checks])
    quick_row.addWidget(btn_none)
    quick_row.addStretch(1)
    layout.addLayout(quick_row)

    buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    buttons.accepted.connect(dialog.accept)
    buttons.rejected.connect(dialog.reject)
    layout.addWidget(buttons)

    if dialog.exec() != QDialog.Accepted:
        return None
    return {col for col, cb in zip(columns, checks) if not cb.isChecked()}


def filter_table_rows(table, text):
    """Ẩn các dòng không chứa `text` (không phân biệt hoa/thường) ở bất kỳ
    cột nào. Trả về số dòng còn hiện."""
    text = (text or "").strip().lower()
    visible = 0
    for row in range(table.rowCount()):
        match = not text or any(
            text in (table.item(row, col).text().lower() if table.item(row, col) else "")
            for col in range(table.columnCount())
        )
        table.setRowHidden(row, not match)
        if match:
            visible += 1
    return visible


def export_table_to_excel(parent, table, columns, highlight_col, default_name, sheet_title):
    """Xuất các dòng đang hiện (sau khi lọc) và các cột đang hiện (sau khi
    chọn cột) của `table` ra file Excel, tô vàng cột `highlight_col` (chỉ số
    0-based trong `columns`, None nếu không tô cột nào)."""
    visible_cols = [col for col in range(len(columns)) if not table.isColumnHidden(col)]
    if not visible_cols:
        QMessageBox.warning(parent, "Chưa có cột nào hiện", "Hãy chọn hiện ít nhất 1 cột trước khi xuất.")
        return
    path, _ = QFileDialog.getSaveFileName(
        parent, "Lưu file Excel", default_name, "Excel Files (*.xlsx)"
    )
    if not path:
        return
    highlight_xlsx_col = (
        visible_cols.index(highlight_col) + 1
        if highlight_col is not None and highlight_col in visible_cols
        else None
    )
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append([columns[col] for col in visible_cols])
    if highlight_xlsx_col is not None:
        ws.cell(row=1, column=highlight_xlsx_col).fill = _MAU_VANG
    for row in range(table.rowCount()):
        if table.isRowHidden(row):
            continue
        values = [
            table.item(row, col).text() if table.item(row, col) else ""
            for col in visible_cols
        ]
        ws.append(values)
        if highlight_xlsx_col is not None:
            ws.cell(row=ws.max_row, column=highlight_xlsx_col).fill = _MAU_VANG
    try:
        wb.save(path)
    except Exception as exc:  # báo lỗi rõ ràng, không để app crash
        QMessageBox.critical(parent, "Lỗi", str(exc))
        return
    QMessageBox.information(parent, "Hoàn tất", f"Đã xuất file:\n{path}")

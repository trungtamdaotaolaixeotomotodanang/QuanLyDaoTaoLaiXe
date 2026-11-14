"""Tiện ích xuất Excel dùng chung cho các tab (openpyxl): mỗi sheet là 1 Excel
Table (có nút lọc/sort sẵn trên từng cột khi mở trong Excel), tự giãn cột.

Các điểm BẮT BUỘC để Excel mở không báo lỗi "Repair" (đã vấp từng cái một ở
quanlygiaovien/tab_tra_cuu.py):
- id của Table phải duy nhất trong cả workbook (Table() mặc định id=1).
- Phải tự tạo tableColumns, tên khớp đúng text dòng tiêu đề (openpyxl không
  tự điền khi gọi ws.add_table).
- totalsRowShown phải đặt False rõ ràng (chuẩn OOXML mặc định là true ->
  Excel coi dòng dữ liệu cuối là dòng tổng -> lệch cấu trúc)."""
import re
from datetime import date
from itertools import repeat

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.filters import AutoFilter
from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

_SHEET_NAME_INVALID_RE = re.compile(r"[\\/*?:\[\]]")
DATE_FORMAT = "dd/mm/yyyy"
# Dòng "cần xem lại" (phiếu không khớp phân công / đã cấp vượt định mức) - tô
# đỏ nhạt cả dòng + chữ đỏ đậm. Định dạng ô trực tiếp đè lên kiểu sọc của
# Table nên vẫn nổi bật khi mở trong Excel.
_FILL_CAN_XEM = PatternFill(fill_type="solid", start_color="FFFDE7E7", end_color="FFFDE7E7")
_FONT_CAN_XEM = Font(color="FFB3261E", bold=True)
GHI_CHU_TO_DO = "Dòng tô đỏ trong file = cần xem lại (phiếu không khớp phân công hoặc đã cấp vượt định mức)."


class ReportWorkbook:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self._used_names = set()
        self._table_count = 0

    def _safe_sheet_name(self, name):
        name = _SHEET_NAME_INVALID_RE.sub("_", name or "").strip() or "Sheet"
        name = name[:31]
        base, i = name, 2
        while name in self._used_names:
            suffix = f" ({i})"
            name = base[: 31 - len(suffix)] + suffix
            i += 1
        self._used_names.add(name)
        return name

    def add_sheet(self, title, header, rows, highlight=None):
        """Thêm 1 sheet: dòng tiêu đề + các dòng dữ liệu (list/tuple cùng độ
        dài với header). Ô kiểu date được ghi là ngày thật (lọc/sort theo
        ngày được trong Excel), số được ghi là số (cộng được). `highlight`:
        list bool cùng độ dài rows - True = tô đỏ cả dòng (cần xem lại)."""
        ws = self.wb.create_sheet(self._safe_sheet_name(title))
        ws.append(list(header))
        for r, can_xem in zip(rows, highlight if highlight is not None else repeat(False)):
            ws.append(list(r))
            for cell in ws[ws.max_row]:
                if isinstance(cell.value, date):
                    cell.number_format = DATE_FORMAT
                if can_xem:
                    cell.fill = _FILL_CAN_XEM
                    cell.font = _FONT_CAN_XEM
        ws.freeze_panes = "A2"
        self._table_count += 1
        _add_filter_table(ws, f"Tbl{self._table_count}", self._table_count)
        _autofit_columns(ws)
        return ws

    def save(self, path):
        if not self.wb.sheetnames:
            raise ValueError("Không có dữ liệu nào để xuất.")
        self.wb.save(path)


def _add_filter_table(ws, table_name, table_id):
    if ws.max_row < 2:
        return  # Table cần ít nhất 1 dòng dữ liệu, chỉ có tiêu đề thì bỏ qua
    last_col_idx = ws.max_column
    last_col = ws.cell(row=1, column=last_col_idx).column_letter
    ref = f"A1:{last_col}{ws.max_row}"
    table = Table(id=table_id, displayName=table_name, ref=ref, totalsRowShown=False)
    table.tableColumns = [
        TableColumn(id=col, name=str(ws.cell(row=1, column=col).value))
        for col in range(1, last_col_idx + 1)
    ]
    table.autoFilter = AutoFilter(ref=ref)
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)


def _autofit_columns(ws, min_width=8, max_width=45):
    for col_idx in range(1, ws.max_column + 1):
        letter = ws.cell(row=1, column=col_idx).column_letter
        width = 0
        for c in ws[letter]:
            if c.value is None:
                continue
            text = c.value.strftime("%d/%m/%Y") if isinstance(c.value, date) else str(c.value)
            width = max(width, len(text))
        # +3 cho nút lọc của Table ở dòng tiêu đề khỏi đè lên chữ
        ws.column_dimensions[letter].width = min(max(width + 3, min_width), max_width)

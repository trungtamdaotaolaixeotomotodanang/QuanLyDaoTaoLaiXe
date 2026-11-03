"""Lớp bọc tối giản quanh xlrd, mô phỏng đúng phần API của openpyxl mà
convert_excel_to_json.py dùng (wb.worksheets, wb.sheetnames, wb[name],
ws.max_row, ws.iter_rows(...)). Dùng để đọc file Excel 97-2003 (.xls) -
openpyxl chỉ đọc được .xlsx.
"""
import xlrd
from xlrd.xldate import xldate_as_datetime


class _Cell:
    __slots__ = ("value",)

    def __init__(self, value):
        self.value = value


class _Sheet:
    def __init__(self, sheet, datemode):
        self._sheet = sheet
        self._datemode = datemode
        self.max_row = sheet.nrows
        self.max_col = sheet.ncols

    def _cell_value(self, row_idx, col_idx):
        cell = self._sheet.cell(row_idx, col_idx)
        if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            return None
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                dt = xldate_as_datetime(cell.value, self._datemode)
            except (xlrd.xldate.XLDateError, ValueError):
                return cell.value
            if dt.hour == dt.minute == dt.second == 0:
                return dt.date()
            return dt
        if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
            return int(cell.value)
        return cell.value

    def iter_rows(self, min_row=1, max_row=None, max_col=None, values_only=False):
        last_row = self.max_row if max_row is None else min(max_row, self.max_row)
        last_col = self.max_col if max_col is None else min(max_col, self.max_col)
        for r in range(min_row - 1, last_row):
            values = tuple(self._cell_value(r, c) for c in range(last_col))
            yield values if values_only else tuple(_Cell(v) for v in values)


class _Workbook:
    def __init__(self, book):
        self.sheetnames = book.sheet_names()
        self.worksheets = [
            _Sheet(book.sheet_by_index(i), book.datemode) for i in range(book.nsheets)
        ]
        self._by_name = dict(zip(self.sheetnames, self.worksheets))

    def __getitem__(self, name):
        return self._by_name[name]


def load_workbook(path):
    book = xlrd.open_workbook(str(path))
    return _Workbook(book)

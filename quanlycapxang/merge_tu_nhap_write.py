"""Ghi kết quả merge_tu_nhap_logic.compute_fill() vào sổ ĐÍCH thật (mở bằng
openpyxl KHÔNG read_only để giữ nguyên toàn bộ định dạng/màu/viền sẵn có của
sổ gốc - khác excel_export.ReportWorkbook vốn tạo sổ report MỚI từ đầu) -
chỉ thêm 3 cột mới vào CUỐI các cột đang có, không đụng tới ô nào khác."""
import copy

from openpyxl import load_workbook

from merge_tu_nhap_logic import NEW_COLUMNS, NEW_HEADERS


def _copy_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def write_filled_columns(target_path, sheet_name, header_row, n_existing_headers, fill_results, out_path):
    """n_existing_headers: số cột đã đọc được ở dòng tiêu đề (kể cả cột cuối
    rỗng, nếu có) - 3 cột mới chèn ngay sau đó, để chắc không ghi đè ô nào
    đang có dữ liệu/định dạng thật. Lưu ra out_path (có thể = target_path để
    ghi đè, hoặc 1 file khác)."""
    wb = load_workbook(target_path)
    ws = wb[sheet_name]
    start_col = n_existing_headers + 1

    # Cột mẫu lấy kiểu (font/màu/viền...) để áp cho 3 cột mới: cột có tiêu đề
    # THẬT gần cuối nhất (không lấy cột rỗng kiểu cột X trong BKeHDMST.xlsx -
    # rỗng nhưng vẫn có ô "has_style" do viền/định dạng kẻ sẵn, dễ chọn nhầm).
    style_src_col = n_existing_headers
    while style_src_col >= 1 and ws.cell(row=header_row, column=style_src_col).value in (None, ""):
        style_src_col -= 1
    style_src_col = style_src_col or 1

    for i, key in enumerate(NEW_COLUMNS):
        col = start_col + i
        header_cell = ws.cell(row=header_row, column=col, value=NEW_HEADERS[key])
        _copy_style(ws.cell(row=header_row, column=style_src_col), header_cell)

    for item in fill_results:
        rn = item["row"]
        for i, key in enumerate(NEW_COLUMNS):
            col = start_col + i
            value = item["values"].get(key)
            if value is None:
                continue
            cell = ws.cell(row=rn, column=col, value=value)
            _copy_style(ws.cell(row=rn, column=style_src_col), cell)

    wb.save(out_path)

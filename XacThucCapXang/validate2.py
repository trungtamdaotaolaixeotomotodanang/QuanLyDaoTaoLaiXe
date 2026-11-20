"""App validate2 - app RIÊNG (không nằm trong main.py) gồm 2 tab:

- "Xác thực dữ liệu" (xem validate_panel.py/validate1_logic.py): xác thực sổ
  cấp nhiên liệu đã có người gõ tay (mặc định sheet PETROLIMEX, đổi được sang
  NLDACAP hay sheet bất kỳ khác qua combo Sheet) - đối chiếu các cột gõ tay
  (giáo viên quản lý/khóa-lớp/biển số/ngày cấp) với dữ liệu khóa/giáo viên/xe
  trong docs/cap_xang_data.json, cho xuất ra Excel cả dữ liệu gốc (dòng sai tô
  đỏ) lẫn danh sách lỗi.
- "Điền dữ liệu" (xem dien_du_lieu_panel.py/dien_du_lieu_logic.py): sổ cái
  hóa đơn GỐC (mặc định quanlycapxang/data/BKeHDMST.xlsx) CHƯA có 3 cột Giáo
  viên quản lý/Khóa/Số phiếu - tự suy luận điền 2 cột đầu khi CHẮC CHẮN (dựa
  theo biển số + ngày hóa đơn khớp đúng 1 phân công trong cap_xang_data.json),
  để trống + cho chọn tay khi mơ hồ, Số phiếu luôn để gõ tay.

2 sheet PETROLIMEX/NLDACAP đặt tên cột khác nhau (PETROLIMEX: "GV Quản lý*/
Khóa*/Số phiếu*" đều ghi "(Tự nhập)"; NLDACAP: "GV quản lý*/Biển số*/Ngày
cấp*" ghi "(Tự nhập)" nhưng "Khoá học"/"Số Phiếu /Quyển" thì không) -
validate1_logic.detect_columns dò theo từ khóa trên toàn bộ tiêu đề nên tự
thích ứng được cả 2, không cần cấu hình riêng ở đây.

App này sống ở thư mục XacThucCapXang/ (TÁCH RIÊNG khỏi quanlycapxang/ cho đỡ
rối - xem yêu cầu dọn dẹp) nhưng vẫn cần dùng lại các module dùng chung của
app CẤP XĂNG (cap_xang_model/runtime_paths/style/excel_validate/attach_files/
excel_export) - nên phải tự thêm thư mục quanlycapxang/ (thư mục liền kề) vào
sys.path TRƯỚC khi import các module đó (xem khối "bootstrap" ngay dưới đây -
PHẢI đứng trước mọi import nội bộ khác).

Chạy: python validate2.py (từ thư mục XacThucCapXang/)."""
import sys
from pathlib import Path

_QUANLYCAPXANG_DIR = Path(__file__).resolve().parent.parent / "quanlycapxang"
if str(_QUANLYCAPXANG_DIR) not in sys.path:
    sys.path.insert(0, str(_QUANLYCAPXANG_DIR))

from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

from cap_xang_model import CapXangModel
from runtime_paths import app_dir
from style import apply_style
from dien_du_lieu_panel import DienDuLieuPanel
from validate_panel import ValidateSheetPanel

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent
CAP_XANG_DATA_JSON = REPO_ROOT / "docs" / "cap_xang_data.json"
DEFAULT_DATA_DIR = APP_DIR / "data"
DEFAULT_FILE = DEFAULT_DATA_DIR / "Re_ Cấp nhiên liệu cho xe tập lái Phương thực hiện và tổng hợp (1).xlsx"
DEFAULT_SHEET = "PETROLIMEX"
DEFAULT_FILL_FILE = DEFAULT_DATA_DIR / "BKeHDMST.xlsx"
DEFAULT_FILL_SHEET = "Sheet1"


class Validate2Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Xác thực & điền dữ liệu cấp xăng")
        self.resize(1300, 860)
        # Dùng CHUNG 1 CapXangModel cho cả 2 tab - không đọc cap_xang_data.json
        # 2 lần.
        model = CapXangModel(CAP_XANG_DATA_JSON)

        tabs = QTabWidget()
        self.validate_panel = ValidateSheetPanel(model, DEFAULT_DATA_DIR, DEFAULT_FILE, DEFAULT_SHEET)
        tabs.addTab(self.validate_panel, "Xác thực dữ liệu")
        self.fill_panel = DienDuLieuPanel(model, DEFAULT_DATA_DIR, DEFAULT_FILL_FILE, DEFAULT_FILL_SHEET)
        tabs.addTab(self.fill_panel, "Điền dữ liệu")
        self.setCentralWidget(tabs)


def main():
    app = QApplication(sys.argv)
    apply_style(app)
    window = Validate2Window()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

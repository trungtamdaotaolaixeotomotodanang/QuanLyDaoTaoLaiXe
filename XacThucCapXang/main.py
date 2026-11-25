"""App XÁC THỰC CẤP XĂNG - app RIÊNG (không nằm trong quanlycapxang/main.py)
gồm 3 tab:

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
- "Xuất phiếu cấp xăng (JSON)" (xem export_phieu_panel.py/export_phieu_logic.py):
  từ sổ đã gõ tay đầy đủ (mặc định sheet NLDACAP - ghi theo từng phiếu, khớp
  đúng độ chi tiết của phieu_cap_xang.json), xác thực rồi xuất JSON cho
  quanlycapxang NHẬP vào (nút "⬇ Nhập phiếu từ JSON..." trong quanlycapxang)
  - thay cho việc gõ tay lại trong lúc quanlycapxang còn đang phát triển.

2 sheet PETROLIMEX/NLDACAP đặt tên cột khác nhau (PETROLIMEX: "GV Quản lý*/
Khóa*/Số phiếu*" đều ghi "(Tự nhập)"; NLDACAP: "GV quản lý*/Biển số*/Ngày
cấp*" ghi "(Tự nhập)" nhưng "Khoá học"/"Số Phiếu /Quyển" thì không) -
validate1_logic.detect_columns dò theo từ khóa trên toàn bộ tiêu đề nên tự
thích ứng được cả 2, không cần cấu hình riêng ở đây.

App này sống ở thư mục XacThucCapXang/ (TÁCH RIÊNG khỏi quanlycapxang/ cho đỡ
rối) nhưng vẫn cần dùng lại các module dùng chung của app CẤP XĂNG
(cap_xang_model/runtime_paths/style/excel_validate/attach_files/excel_export)
- nên phải tự thêm thư mục quanlycapxang/ (thư mục liền kề) vào sys.path
TRƯỚC khi import các module đó (xem khối "bootstrap" ngay dưới đây - PHẢI
đứng trước mọi import nội bộ khác).

Chạy: python main.py (từ thư mục XacThucCapXang/)."""
import sys
from pathlib import Path


def _this_app_dir():
    """Thư mục thật chứa main.py nguồn, hoặc thư mục chứa main.exe khi đóng
    gói (__file__ trỏ sai vào thư mục giải nén tạm lúc PyInstaller --onefile
    chạy - cùng lý do/cách sửa với quanlycapxang/runtime_paths.app_dir)."""
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return exe_dir.parent if exe_dir.name == "dist" else exe_dir
    return Path(__file__).resolve().parent


_QUANLYCAPXANG_DIR = _this_app_dir().parent / "quanlycapxang"
if str(_QUANLYCAPXANG_DIR) not in sys.path:
    sys.path.insert(0, str(_QUANLYCAPXANG_DIR))

from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

from cap_xang_model import CapXangModel
from style import apply_style

from dien_du_lieu_panel import DienDuLieuPanel
from export_phieu_panel import ExportPhieuPanel
from validate_panel import ValidateSheetPanel

REPO_ROOT = _QUANLYCAPXANG_DIR.parent
CAP_XANG_DATA_JSON = REPO_ROOT / "docs" / "cap_xang_data.json"

# Sổ cấp xăng nhân viên gõ tay (mới nhất) nằm ở quanlycapxang/import/ (xem
# .gitignore) - KHÔNG phải quanlycapxang/data/ (đó là file tham chiếu/định
# mức ít đổi, xem BKeHDMST.xlsx bên dưới).
IMPORT_DIR = _QUANLYCAPXANG_DIR / "import"
LEDGER_FILE = IMPORT_DIR / "Re_ Cấp nhiên liệu cho xe tập lái Phương thực hiện và tổng hợp (1).xlsx"

DATA_DIR = _QUANLYCAPXANG_DIR / "data"
BKEHDMST_FILE = DATA_DIR / "BKeHDMST.xlsx"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Xác thực & xuất dữ liệu cấp xăng")
        self.resize(1300, 860)
        # Dùng CHUNG 1 CapXangModel cho cả 3 tab - không đọc cap_xang_data.json
        # nhiều lần.
        model = CapXangModel(CAP_XANG_DATA_JSON)

        tabs = QTabWidget()
        self.validate_panel = ValidateSheetPanel(model, IMPORT_DIR, LEDGER_FILE, "PETROLIMEX")
        tabs.addTab(self.validate_panel, "Xác thực dữ liệu")
        self.fill_panel = DienDuLieuPanel(model, DATA_DIR, BKEHDMST_FILE, "Sheet1")
        tabs.addTab(self.fill_panel, "Điền dữ liệu")
        self.export_panel = ExportPhieuPanel(model, IMPORT_DIR, LEDGER_FILE, "NLDACAP")
        tabs.addTab(self.export_panel, "Xuất phiếu cấp xăng (JSON)")
        self.setCentralWidget(tabs)


def main():
    app = QApplication(sys.argv)
    apply_style(app)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

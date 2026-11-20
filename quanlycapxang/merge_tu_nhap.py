"""App RIÊNG (không nằm trong main.py) - tự động bổ sung 3 cột "Tự nhập"
(GV Quản lý / Khóa / Số phiếu) vào BKeHDMST.xlsx (sổ cái hóa đơn gốc xuất ra
từ phần mềm Petrolimex, chưa ai gõ 3 cột này), lấy dữ liệu từ sổ "Re_ Cấp
nhiên liệu cho xe tập lái..." (đã có người gõ tay) - khớp theo "Mã tra cứu".

CHỈ điền vào các dòng chắc chắn khớp (xem merge_tu_nhap_logic.py), dòng chưa
chắc thì để trống, người dùng tự nhập sau - đúng yêu cầu gốc. Người dùng đổi
được sang file/sheet/cột khác nếu cần (combo + nút "Duyệt..." - xem
merge_tu_nhap_panel.py).

merge_tu_nhap_panel.py dùng lại validate1_logic.detect_columns() - module này
đã dọn sang thư mục XacThucCapXang/ (liền kề quanlycapxang/, xem yêu cầu dọn
dẹp) nên phải tự thêm thư mục đó vào sys.path TRƯỚC khi import
merge_tu_nhap_panel (import module đó sẽ kéo theo `import validate1_logic`).

Chạy: python merge_tu_nhap.py (từ thư mục quanlycapxang/)."""
import sys
from pathlib import Path

_XAC_THUC_CAP_XANG_DIR = Path(__file__).resolve().parent.parent / "XacThucCapXang"
if str(_XAC_THUC_CAP_XANG_DIR) not in sys.path:
    sys.path.insert(0, str(_XAC_THUC_CAP_XANG_DIR))

from PySide6.QtWidgets import QApplication, QMainWindow

from merge_tu_nhap_panel import MergeTuNhapPanel
from runtime_paths import app_dir
from style import apply_style

APP_DIR = app_dir()
DEFAULT_DATA_DIR = APP_DIR / "data"
DEFAULT_TARGET_FILE = DEFAULT_DATA_DIR / "BKeHDMST.xlsx"
DEFAULT_TARGET_SHEET = "Sheet1"
DEFAULT_SOURCE_FILE = DEFAULT_DATA_DIR / "Re_ Cấp nhiên liệu cho xe tập lái Phương thực hiện và tổng hợp (1).xlsx"
DEFAULT_SOURCE_SHEET = "PETROLIMEX"


class MergeTuNhapWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bổ sung 3 cột GV/Khóa/Số phiếu cho BKeHDMST")
        self.resize(1300, 850)
        self.panel = MergeTuNhapPanel(
            DEFAULT_DATA_DIR, DEFAULT_TARGET_FILE, DEFAULT_TARGET_SHEET,
            DEFAULT_SOURCE_FILE, DEFAULT_SOURCE_SHEET,
        )
        self.setCentralWidget(self.panel)


def main():
    app = QApplication(sys.argv)
    apply_style(app)
    window = MergeTuNhapWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

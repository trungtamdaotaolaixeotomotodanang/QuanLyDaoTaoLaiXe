import json
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

import publish_data
import roster_xml
from data_model import DataModel
from runtime_paths import app_dir
from style import apply_style
from tab_cap_nhat_web import CapNhatWebTab
from tab_danh_sach_gv import DanhSachGiaoVienTab
from tab_lich_hoc import LichHocTab
from tab_phan_xe import PhanXeTab
from tab_tra_cuu import TraCuuTab
from tab_tra_cuu_xe import TraCuuXeTab

APP_DIR = app_dir()
DATA_JSON = APP_DIR / "data.json"


class MainWindow(QMainWindow):
    def __init__(self, model):
        super().__init__()
        self.setWindowTitle("Quản lý đào tạo lái xe - Tra cứu giáo viên / xe / lịch học")
        self.resize(1200, 850)

        tabs = QTabWidget()
        tabs.addTab(CapNhatWebTab(), "CẬP NHẬT WEB")
        tabs.addTab(DanhSachGiaoVienTab(), "DANH SÁCH GIÁO VIÊN")
        tabs.addTab(TraCuuTab(model), "TRA CỨU")
        tabs.addTab(PhanXeTab(model), "PHÂN XE")
        tabs.addTab(TraCuuXeTab(model), "DANH SÁCH XE")
        tabs.addTab(LichHocTab(model), "LỊCH HỌC")
        self.setCentralWidget(tabs)


def main():
    app = QApplication(sys.argv)
    apply_style(app)

    if not DATA_JSON.exists():
        # Chưa đăng dữ liệu lần nào - tạo file rỗng để mở app được, người
        # dùng vào tab "CẬP NHẬT WEB" chọn nguồn dữ liệu rồi bấm đăng.
        DATA_JSON.write_text(
            json.dumps(
                {"generated_at": "", "source_files": [], "courses": []},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

    dsgv_file = publish_data.load_data_sources().get("dsgv_file") or ""
    dsgv = roster_xml.parse_dsgv(dsgv_file) if dsgv_file and Path(dsgv_file).exists() else {}

    model = DataModel(DATA_JSON, dsgv=dsgv)
    window = MainWindow(model)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

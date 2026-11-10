import json
import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

import git_publish
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
        self.model = model

        self.tab_cap_nhat_web = CapNhatWebTab()
        self.tab_danh_sach_gv = DanhSachGiaoVienTab()
        self.tab_tra_cuu = TraCuuTab(model)
        self.tab_phan_xe = PhanXeTab(model)
        self.tab_tra_cuu_xe = TraCuuXeTab(model)
        self.tab_lich_hoc = LichHocTab(model)

        tabs = QTabWidget()
        tabs.addTab(self.tab_cap_nhat_web, "CẬP NHẬT WEB")
        tabs.addTab(self.tab_danh_sach_gv, "DANH SÁCH GIÁO VIÊN")
        tabs.addTab(self.tab_tra_cuu, "TRA CỨU")
        tabs.addTab(self.tab_phan_xe, "PHÂN XE")
        tabs.addTab(self.tab_tra_cuu_xe, "DANH SÁCH XE")
        tabs.addTab(self.tab_lich_hoc, "LỊCH HỌC")
        self.setCentralWidget(tabs)

        # Trước đây phải tắt/bật lại app mới thấy dữ liệu mới sau khi bấm
        # "Cập nhật && Đăng lên web" (các tab khác chỉ đọc dữ liệu 1 lần lúc
        # mở app) - giờ nạp lại ngay cho mọi tab khi tab CẬP NHẬT WEB báo đã
        # ghi xong data.json mới.
        self.tab_cap_nhat_web.data_published.connect(self._on_data_published)
        self.tab_lich_hoc.lich_hoc_source_added.connect(
            self.tab_cap_nhat_web.pick_lich_hoc.add_path_if_missing
        )

    def _on_data_published(self):
        dsgv_file = publish_data.load_data_sources().get("dsgv_file") or ""
        dsgv = roster_xml.parse_dsgv(dsgv_file) if dsgv_file and Path(dsgv_file).exists() else {}
        self.model.reload(dsgv)
        publish_data.apply_ngay_hoc_overrides_fallback(self.model.courses)
        self.tab_tra_cuu.refresh()
        self.tab_phan_xe.refresh()
        self.tab_tra_cuu_xe.refresh()
        self.tab_lich_hoc.refresh()
        self.tab_danh_sach_gv.refresh()


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

    # Kéo về ngay_hoc_overrides.json (và các file công khai khác) mới nhất
    # từ máy khác trước khi nạp model - để ngày gán tay ở máy khác cũng
    # hiện ra ở đây, không chỉ ngày tự gán tay trên chính máy này. Chỉ thử
    # tốt nhất có thể (best-effort) - lỗi mạng/offline thì bỏ qua, vẫn mở
    # app bình thường với dữ liệu cục bộ hiện có.
    try:
        git_publish.sync_from_remote()
    except Exception:
        pass

    dsgv_file = publish_data.load_data_sources().get("dsgv_file") or ""
    dsgv = roster_xml.parse_dsgv(dsgv_file) if dsgv_file and Path(dsgv_file).exists() else {}

    model = DataModel(DATA_JSON, dsgv=dsgv)
    publish_data.apply_ngay_hoc_overrides_fallback(model.courses)
    window = MainWindow(model)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

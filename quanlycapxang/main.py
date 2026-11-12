import sys

from PySide6.QtCore import QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

import git_sync
from cap_xang_model import CapXangModel
from runtime_paths import app_dir
from style import apply_style
from tab_cap_xang import CapXangTab
from tab_dinh_muc_xang import DinhMucXangTab
from tab_tong_hop_cap_xang import TongHopCapXangTab

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent
CAP_XANG_DATA_JSON = REPO_ROOT / "docs" / "cap_xang_data.json"
SHARED_DATA_DIR = REPO_ROOT / "shared_data"


class _SyncWorker(QThread):
    """Chạy git_sync.sync_from_remote() ở thread riêng để không treo UI -
    xem MainWindow._on_sync_clicked."""

    finished_result = Signal(bool, str)

    def run(self):
        try:
            ok, message = git_sync.sync_from_remote()
        except Exception as exc:
            ok, message = False, str(exc)
        self.finished_result.emit(ok, message)


class _PushWorker(QThread):
    """Chạy git_sync.push_shared_data() ở thread riêng - xem
    MainWindow._on_push_clicked."""

    finished_result = Signal(bool, str)

    def run(self):
        try:
            ok, message = git_sync.push_shared_data("Cập nhật phiếu/định mức cấp xăng")
        except Exception as exc:
            ok, message = False, str(exc)
        self.finished_result.emit(ok, message)


class MainWindow(QMainWindow):
    def __init__(self, model):
        super().__init__()
        self.setWindowTitle("Quản lý cấp xăng")
        self.resize(1200, 800)
        self.model = model

        self.tab_dinh_muc_xang = DinhMucXangTab(model)
        self.tab_cap_xang = CapXangTab(model)
        self.tab_tong_hop_cap_xang = TongHopCapXangTab(model)

        tabs = QTabWidget()
        tabs.addTab(self.tab_dinh_muc_xang, "ĐỊNH MỨC XĂNG")
        tabs.addTab(self.tab_cap_xang, "CẤP XĂNG")
        tabs.addTab(self.tab_tong_hop_cap_xang, "TỔNG HỢP CẤP XĂNG")

        self.btn_sync = QPushButton("🔄 Lấy dữ liệu mới nhất")
        self.btn_sync.clicked.connect(self._on_sync_clicked)
        self.btn_push = QPushButton("⬆ Đẩy phiếu/định mức lên")
        self.btn_push.clicked.connect(self._on_push_clicked)
        top_row = QHBoxLayout()
        top_row.addStretch()
        top_row.addWidget(self.btn_push)
        top_row.addWidget(self.btn_sync)

        central = QWidget()
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(8, 6, 8, 8)
        central_layout.addLayout(top_row)
        central_layout.addWidget(tabs)
        self.setCentralWidget(central)

    def _on_sync_clicked(self):
        self.btn_sync.setEnabled(False)
        self.btn_sync.setText("Đang lấy dữ liệu...")
        self._sync_worker = _SyncWorker()
        self._sync_worker.finished_result.connect(self._on_sync_done)
        self._sync_worker.start()

    def _on_sync_done(self, ok, message):
        self.btn_sync.setEnabled(True)
        self.btn_sync.setText("🔄 Lấy dữ liệu mới nhất")
        if ok:
            self.model.reload()
            self.tab_dinh_muc_xang.refresh()
            self.tab_cap_xang.refresh()
            self.tab_tong_hop_cap_xang.refresh()
            QMessageBox.information(
                self, "Hoàn tất",
                "Đã lấy dữ liệu khóa/xe/giáo viên + phiếu/định mức (máy khác "
                "đã đẩy lên) mới nhất từ GitHub.\n\n"
                "File dữ liệu đã tải về:\n"
                f"{CAP_XANG_DATA_JSON}\n"
                f"{SHARED_DATA_DIR / 'phieu_cap_xang.json'}\n"
                f"{SHARED_DATA_DIR / 'dinh_muc_xang.json'}",
            )
        else:
            QMessageBox.warning(self, "Không lấy được dữ liệu mới", message)

    def _on_push_clicked(self):
        self.btn_push.setEnabled(False)
        self.btn_push.setText("Đang đẩy lên...")
        self._push_worker = _PushWorker()
        self._push_worker.finished_result.connect(self._on_push_done)
        self._push_worker.start()

    def _on_push_done(self, ok, message):
        self.btn_push.setEnabled(True)
        self.btn_push.setText("⬆ Đẩy phiếu/định mức lên")
        if ok:
            QMessageBox.information(
                self, "Hoàn tất",
                f"{message}\n\n"
                "File đã đẩy lên (máy khác bấm \"🔄 Lấy dữ liệu mới nhất\" sẽ thấy):\n"
                f"{SHARED_DATA_DIR / 'phieu_cap_xang.json'}\n"
                f"{SHARED_DATA_DIR / 'dinh_muc_xang.json'}",
            )
        else:
            QMessageBox.warning(self, "Không đẩy lên được", message)


def main():
    app = QApplication(sys.argv)
    apply_style(app)

    # Best-effort lúc mở app - lỗi mạng/offline thì bỏ qua, vẫn mở app bình
    # thường với dữ liệu cục bộ hiện có (xem nút "Lấy dữ liệu mới nhất" để
    # thử lại giữa chừng, có báo rõ kết quả hơn).
    try:
        git_sync.sync_from_remote()
    except Exception:
        pass

    model = CapXangModel(CAP_XANG_DATA_JSON)
    window = MainWindow(model)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

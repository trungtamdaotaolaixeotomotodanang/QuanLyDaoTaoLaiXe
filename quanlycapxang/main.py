import sys

from PySide6.QtCore import QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

import git_sync
import phieu_import
from cap_xang_model import CapXangModel
from runtime_paths import app_dir
from style import apply_style
from tab_cap_xang import CapXangTab
from tab_dinh_muc_xang import DinhMucXangTab
from tab_tong_hop_cap_xang import TongHopCapXangTab
from tab_tong_quan import TongQuanTab

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
        self.tab_tong_quan = TongQuanTab(
            model, tabs=tabs, tab_dinh_muc=self.tab_dinh_muc_xang, tab_tong_hop=self.tab_tong_hop_cap_xang,
        )
        tabs.addTab(self.tab_tong_quan, "TỔNG QUAN")
        tabs.addTab(self.tab_dinh_muc_xang, "ĐỊNH MỨC XĂNG")
        tabs.addTab(self.tab_cap_xang, "CẤP XĂNG")
        tabs.addTab(self.tab_tong_hop_cap_xang, "TỔNG HỢP CẤP XĂNG")

        self.lbl_unpushed = QLabel("")
        self.lbl_unpushed.setStyleSheet("color: #b3261e; font-weight: 600;")
        self.btn_sync = QPushButton("🔄 Lấy dữ liệu mới nhất")
        self.btn_sync.clicked.connect(self._on_sync_clicked)
        self.btn_import = QPushButton("⬇ Nhập phiếu từ JSON...")
        self.btn_import.clicked.connect(self._on_import_clicked)
        self.btn_push = QPushButton("⬆ Đẩy phiếu/định mức lên")
        self.btn_push.clicked.connect(self._on_push_clicked)
        top_row = QHBoxLayout()
        top_row.addWidget(self.lbl_unpushed)
        top_row.addStretch()
        top_row.addWidget(self.btn_import)
        top_row.addWidget(self.btn_push)
        top_row.addWidget(self.btn_sync)

        central = QWidget()
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(8, 6, 8, 8)
        central_layout.addLayout(top_row)
        central_layout.addWidget(tabs)
        self.setCentralWidget(central)

        # Nhắc "còn phiếu chưa đẩy lên" - chỉ đọc git status cục bộ (không
        # qua mạng) nên kiểm tra định kỳ vẫn nhẹ; tránh quên bấm "⬆ Đẩy..."
        # rồi tưởng máy khác đã thấy dữ liệu mới.
        self._unpushed_timer = QTimer(self)
        self._unpushed_timer.setInterval(20000)
        self._unpushed_timer.timeout.connect(self._check_unpushed)
        self._unpushed_timer.start()
        QTimer.singleShot(3000, self._check_unpushed)

    def _check_unpushed(self):
        try:
            has = git_sync.has_unpushed_changes()
        except Exception:
            return
        self.lbl_unpushed.setText("⚠ Có phiếu/định mức chưa đẩy lên GitHub" if has else "")

    def _on_sync_clicked(self):
        self.btn_sync.setEnabled(False)
        self.btn_sync.setText("Đang lấy dữ liệu...")
        self._sync_worker = _SyncWorker()
        self._sync_worker.finished_result.connect(self._on_sync_done)
        self._sync_worker.start()

    def start_background_sync(self):
        """Tự lấy dữ liệu mới nhất ngay lúc mở app - chạy NỀN (không chặn
        cửa sổ chính) và ÂM THẦM (không hiện hộp thoại dù thành công hay lỗi
        mạng/offline - đúng tinh thần "best-effort" lúc mở app, xem nút
        "🔄 Lấy dữ liệu mới nhất" nếu muốn thử lại có báo kết quả rõ ràng).
        Trước đây gọi git_sync.sync_from_remote() trực tiếp trong main() làm
        cả app bị treo tới khi xong (hoặc treo vô thời hạn nếu mạng chập
        chờn, do chưa có timeout) - giờ cửa sổ đã hiện ra ngay, sync chạy sau."""
        self.btn_sync.setEnabled(False)
        self.btn_sync.setText("Đang lấy dữ liệu...")
        self._startup_sync_worker = _SyncWorker()
        self._startup_sync_worker.finished_result.connect(self._on_startup_sync_done)
        self._startup_sync_worker.start()

    def _on_startup_sync_done(self, ok, _message):
        self.btn_sync.setEnabled(True)
        self.btn_sync.setText("🔄 Lấy dữ liệu mới nhất")
        if ok:
            self.model.reload()
            self.tab_dinh_muc_xang.refresh()
            self.tab_cap_xang.refresh()
            self.tab_tong_hop_cap_xang.refresh()
            self.tab_tong_quan.refresh()

    def _on_sync_done(self, ok, message):
        self.btn_sync.setEnabled(True)
        self.btn_sync.setText("🔄 Lấy dữ liệu mới nhất")
        if ok:
            self.model.reload()
            self.tab_dinh_muc_xang.refresh()
            self.tab_cap_xang.refresh()
            self.tab_tong_hop_cap_xang.refresh()
            self.tab_tong_quan.refresh()
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

    def _on_import_clicked(self):
        """Nhập phiếu từ file JSON do XacThucCapXang xuất (xem
        phieu_import.py) - chạy đồng bộ (chỉ đọc file cục bộ, không qua
        mạng/git nên không cần QThread như sync/push)."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file JSON phiếu cấp xăng", "", "JSON Files (*.json)"
        )
        if not path:
            return
        try:
            added, skipped_trung, errors = phieu_import.import_from_file(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không nhập được", str(exc))
            return

        if added:
            self.model.reload()
            self.tab_dinh_muc_xang.refresh()
            self.tab_cap_xang.refresh()
            self.tab_tong_hop_cap_xang.refresh()
            self.tab_tong_quan.refresh()

        msg = f"Đã nhập {added} phiếu mới, bỏ qua {skipped_trung} phiếu đã có."
        if errors:
            msg += f"\n\n{len(errors)} phiếu lỗi không nhập được:\n" + "\n".join(errors[:20])
        QMessageBox.information(self, "Hoàn tất", msg)

    def _on_push_clicked(self):
        self.btn_push.setEnabled(False)
        self.btn_push.setText("Đang đẩy lên...")
        self._push_worker = _PushWorker()
        self._push_worker.finished_result.connect(self._on_push_done)
        self._push_worker.start()

    def _on_push_done(self, ok, message):
        self.btn_push.setEnabled(True)
        self.btn_push.setText("⬆ Đẩy phiếu/định mức lên")
        self._check_unpushed()
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

    model = CapXangModel(CAP_XANG_DATA_JSON)
    window = MainWindow(model)
    window.show()
    # Hiện cửa sổ với dữ liệu cục bộ hiện có TRƯỚC, rồi mới âm thầm lấy dữ
    # liệu mới nhất ở nền (xem start_background_sync) - không chặn app chờ
    # mạng như trước đây.
    window.start_background_sync()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

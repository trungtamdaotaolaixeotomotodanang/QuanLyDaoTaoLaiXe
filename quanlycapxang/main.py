import sys

from PySide6.QtCore import QThread, QTimer, Signal
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QMainWindow,
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

# Không còn nút bấm tay - cứ 10 phút tự kiểm tra GitHub có gì mới 1 lần (xem
# MainWindow._queue_sync), cộng với lúc mở app (QTimer.singleShot(0, ...)).
_SYNC_INTERVAL_MS = 10 * 60 * 1000


class _SyncWorker(QThread):
    """Chạy git_sync.sync_from_remote() ở thread riêng để không treo UI -
    xem MainWindow._queue_sync."""

    finished_result = Signal(bool, str)

    def run(self):
        try:
            ok, message = git_sync.sync_from_remote()
        except Exception as exc:
            ok, message = False, str(exc)
        self.finished_result.emit(ok, message)


class _PushWorker(QThread):
    """Chạy git_sync.push_shared_data() ở thread riêng - xem
    MainWindow._queue_push."""

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
        tabs.addTab(self.tab_cap_xang, "CẤP XĂNG")
        tabs.addTab(self.tab_dinh_muc_xang, "ĐỊNH MỨC XĂNG")
        tabs.addTab(self.tab_tong_hop_cap_xang, "TỔNG HỢP CẤP XĂNG")

        # Không còn nút "🔄 Lấy dữ liệu mới nhất"/"⬆ Đẩy..." - cả 2 việc giờ
        # TỰ chạy (xem _queue_sync/_queue_push) - chỉ giữ 1 dòng chữ báo nếu
        # có sự cố (vd mất mạng khiến đẩy lên thất bại) để biết mà kiểm tra.
        self.lbl_trang_thai = QLabel("")
        self.lbl_trang_thai.setStyleSheet("color: #b3261e; font-weight: 600;")
        top_row = QHBoxLayout()
        top_row.addWidget(self.lbl_trang_thai)
        top_row.addStretch()

        central = QWidget()
        central_layout = QVBoxLayout(central)
        central_layout.setContentsMargins(8, 6, 8, 8)
        central_layout.addLayout(top_row)
        central_layout.addWidget(tabs)
        self.setCentralWidget(central)

        # Hàng đợi 1 lệnh git tại 1 thời điểm (push HOẶC sync, không chạy
        # song song - 2 lệnh git cùng lúc trên cùng 1 repo không an toàn).
        # Push được ưu tiên hơn (đẩy thay đổi của máy này lên trước), sync
        # định kỳ/lúc mở app xếp sau nếu đang có push chạy - xem _run_git_queue.
        self._git_worker = None
        self._dang_cho_sync = False
        self._dang_cho_push = False

        # Thêm/sửa/xóa phiếu hoặc định mức ở tab CẤP XĂNG/TỔNG HỢP CẤP XĂNG
        # -> tự đẩy lên GitHub NGAY (xem *.du_lieu_da_doi/_queue_push).
        self.tab_cap_xang.du_lieu_da_doi.connect(self._queue_push)
        self.tab_tong_hop_cap_xang.du_lieu_da_doi.connect(self._queue_push)

        # Lấy dữ liệu mới nhất ngay lúc mở app (hoãn 0ms để cửa sổ hiện ra
        # trước, không chặn chờ mạng) + định kỳ 10 phút/lần sau đó.
        QTimer.singleShot(0, self._queue_sync)
        self._sync_timer = QTimer(self)
        self._sync_timer.setInterval(_SYNC_INTERVAL_MS)
        self._sync_timer.timeout.connect(self._queue_sync)
        self._sync_timer.start()

    # ---------------- hàng đợi git (sync kéo về / push đẩy lên) ----------------
    def _queue_sync(self):
        self._dang_cho_sync = True
        self._run_git_queue()

    def _queue_push(self):
        self._dang_cho_push = True
        self._run_git_queue()

    def _run_git_queue(self):
        if self._git_worker is not None:
            return  # đang có lệnh git chạy - lượt này tự chạy tiếp lúc lệnh đó xong (xem _on_git_done)
        if self._dang_cho_push:
            self._dang_cho_push = False
            worker = _PushWorker()
            self._git_worker = worker
            worker.finished_result.connect(lambda ok, msg: self._on_git_done("push", ok, msg))
            worker.start()
        elif self._dang_cho_sync:
            self._dang_cho_sync = False
            worker = _SyncWorker()
            self._git_worker = worker
            worker.finished_result.connect(lambda ok, msg: self._on_git_done("sync", ok, msg))
            worker.start()

    def _on_git_done(self, kind, ok, message):
        self._git_worker = None
        if kind == "sync":
            if ok:
                self.model.reload()
                self.tab_dinh_muc_xang.refresh()
                self.tab_cap_xang.refresh()
                self.tab_tong_hop_cap_xang.refresh()
            # Lỗi lúc tự sync (vd mất mạng) - BỎ QUA ÂM THẦM, đúng kiểu
            # "best-effort": còn lượt định kỳ 10 phút sau tự thử lại.
        else:  # push
            if ok:
                self.lbl_trang_thai.setText("")
            else:
                self.lbl_trang_thai.setText(
                    "⚠ Không tự đẩy phiếu/định mức lên GitHub được (xem chi tiết) - "
                    "dữ liệu vẫn được lưu ở máy này, app sẽ tự thử lại."
                )
                self.lbl_trang_thai.setToolTip(message)
        self._run_git_queue()


def main():
    app = QApplication(sys.argv)
    apply_style(app)

    model = CapXangModel(CAP_XANG_DATA_JSON)
    window = MainWindow(model)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

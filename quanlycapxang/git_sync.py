"""Lấy docs/cap_xang_data.json mới nhất từ GitHub (app này CHỈ ĐỌC, không
bao giờ ghi/push gì lên git - phiếu cấp xăng/định mức là sổ riêng từng máy,
xem cap_xang_store.py/fuel_norm.py). Cùng cơ chế/lý do với
quanlygiaovien/git_publish.py.sync_from_remote() - xem file đó để biết chi
tiết (--autostash, vì sao cần _CREATE_NO_WINDOW khi đóng gói .exe...)."""
import subprocess
import sys

from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent  # git root thực tế là thư mục chứa cả quanlygiaovien/ lẫn quanlycapxang/

_CREATE_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0


def _run(args):
    return subprocess.run(
        ["git", *args],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=_CREATE_NO_WINDOW,
    )


def sync_from_remote():
    """Kéo về + rebase commit cục bộ (nếu có, bình thường không có vì app
    này không tự tạo commit) lên trên bản mới nhất trên GitHub. Trả về
    (ok, message)."""
    fetch = _run(["fetch", "origin"])
    if fetch.returncode != 0:
        return False, f"Lỗi khi 'git fetch':\n{fetch.stderr}"

    branch = _run(["rev-parse", "--abbrev-ref", "HEAD"]).stdout.strip()
    rebase = _run(["rebase", "--autostash", f"origin/{branch}"])
    if rebase.returncode != 0:
        _run(["rebase", "--abort"])
        return False, (
            "Có xung đột khi gộp (không tự gộp được) - đã hủy gộp. Chi "
            "tiết:\n" + rebase.stderr
        )
    return True, ""

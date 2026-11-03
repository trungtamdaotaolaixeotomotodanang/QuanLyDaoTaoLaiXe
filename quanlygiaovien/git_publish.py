"""Tự động commit + push dữ liệu công khai lên GitHub (để GitHub Pages cập nhật).

Người dùng không cần biết git: hàm push_public_data() tự làm mọi thứ,
chỉ commit/push đúng những file được phép public (xem .gitignore).
"""
import subprocess

from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent  # git root thực tế là D:\GTCC (docs/ nằm ở đây)

FILES_TO_PUBLISH = [
    "docs/public_data.json",
    "docs/files",
]


def _run(args):
    return subprocess.run(
        ["git", *args],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def push_public_data(commit_message):
    """Trả về (ok: bool, message: str)."""
    existing = [f for f in FILES_TO_PUBLISH if (REPO_ROOT / f).exists()]
    if not existing:
        return False, "Không có file dữ liệu công khai nào để cập nhật."

    add = _run(["add", "-A", "--", *existing])
    if add.returncode != 0:
        return False, f"Lỗi khi 'git add':\n{add.stderr}"

    status = _run(["status", "--porcelain", *existing])
    if not status.stdout.strip():
        return True, "Dữ liệu không có gì thay đổi so với lần cập nhật trước."

    commit = _run(["commit", "-m", commit_message])
    if commit.returncode != 0:
        return False, f"Lỗi khi 'git commit':\n{commit.stderr}"

    push = _run(["push"])
    if push.returncode != 0:
        return False, (
            "Đã lưu thay đổi ở máy này nhưng KHÔNG đẩy được lên GitHub "
            "(có thể do mất mạng hoặc cần đăng nhập lại). Chi tiết:\n"
            f"{push.stderr}"
        )

    return True, "Đã cập nhật và đăng lên web thành công."

"""Tự động commit + push dữ liệu công khai lên GitHub (để GitHub Pages cập nhật).

Người dùng không cần biết git: hàm push_public_data() tự làm mọi thứ,
chỉ commit/push đúng những file được phép public (xem .gitignore).
"""
import subprocess
import sys

from runtime_paths import app_dir

APP_DIR = app_dir()
REPO_ROOT = APP_DIR.parent  # git root thực tế là D:\GTCC (docs/ nằm ở đây)

# ten_gv_overrides.json/lich_hoc_overrides.json/ngay_hoc_overrides.json được
# publish_data.py lưu trong APP_DIR - không hardcode "quanlygiaovien/..." ở
# đây vì APP_DIR (tính từ runtime_paths.app_dir(), theo vị trí main.exe) có
# thể KHÁC thư mục "quanlygiaovien" nếu .exe không được đặt trong 1 thư mục
# tên đúng "dist" - lúc đó git add sẽ add nhầm 1 file đứng yên (không phải
# file thật đang được ghi), khiến máy khác không đồng bộ được ngày gán tay/
# tên GV override (từng gặp thực tế: file ngay_hoc_overrides.json không hề
# được commit mặc dù app báo lưu + đăng web thành công). Tính theo APP_DIR
# (dùng chung hàm app_dir() như publish_data.py) để luôn khớp đúng file thật.
FILES_TO_PUBLISH = [
    "docs/public_data.json",
    "docs/cap_xang_data.json",  # dữ liệu gọn cho app Quản lý cấp xăng (quanlycapxang/)
    "docs/files",
    str((APP_DIR / "ten_gv_overrides.json").relative_to(REPO_ROOT)),
    str((APP_DIR / "lich_hoc_overrides.json").relative_to(REPO_ROOT)),
    str((APP_DIR / "ngay_hoc_overrides.json").relative_to(REPO_ROOT)),
]

# Khi chạy từ file .exe đã build (không có console riêng - xem main.spec
# console=False), subprocess.run spawn "git" sẽ bị Windows tự mở 1 cửa sổ
# cmd đen mới cho MỖI lệnh (fetch/rebase/add/commit/push) vì không có
# console nào để gắn vào - cần cờ này để Windows không mở cửa sổ đó. Khi
# chạy bằng "python main.py" từ terminal thì không bị (đã có console sẵn)
# nên trước đây không phát hiện ra.
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
    """Kéo về + rebase commit cục bộ (nếu có) lên trên bản mới nhất trên
    GitHub - để biết/dùng dữ liệu mà máy khác vừa đăng, và để lần push sau
    không bị từ chối (non-fast-forward). --autostash để không vướng các file
    sửa dở khác (code...) không liên quan tới docs/.

    Gọi hàm này TRƯỚC KHI đọc/sửa docs/public_data.json (xem add_teacher_photo,
    remove_teacher_photo, run_pipeline trong publish_data.py) - để luôn sửa
    trên bản mới nhất, không ghi đè mất dữ liệu máy khác vừa thêm.
    Trả về (ok, message)."""
    fetch = _run(["fetch", "origin"])
    if fetch.returncode != 0:
        return False, f"Lỗi khi 'git fetch':\n{fetch.stderr}"

    branch = _run(["rev-parse", "--abbrev-ref", "HEAD"]).stdout.strip()
    rebase = _run(["rebase", "--autostash", f"origin/{branch}"])
    if rebase.returncode != 0:
        _run(["rebase", "--abort"])
        return False, (
            "Có xung đột dữ liệu với máy khác khi gộp (không tự gộp được) - "
            "đã hủy gộp, dữ liệu máy này vẫn còn nguyên, cần người biết git "
            "xử lý tay. Chi tiết:\n" + rebase.stderr
        )
    return True, ""


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

    # Kéo về + gộp trước khi đẩy lên, để tránh bị từ chối (non-fast-forward)
    # nếu máy khác vừa đăng dữ liệu trước đó (bình thường đã gọi
    # sync_from_remote() trước khi sửa file rồi nên bước này thường chỉ là
    # fast-forward, không có gì để gộp).
    ok, err = sync_from_remote()
    if not ok:
        return False, err

    push = _run(["push"])
    if push.returncode != 0:
        # Có thể máy khác vừa đăng xong trong lúc mình rebase/push - thử lại
        # một lần nữa cho chắc trước khi báo lỗi.
        ok, err = sync_from_remote()
        if not ok:
            return False, err
        push = _run(["push"])
        if push.returncode != 0:
            return False, (
                "Đã lưu thay đổi ở máy này nhưng KHÔNG đẩy được lên GitHub "
                "(có thể do mất mạng hoặc cần đăng nhập lại). Chi tiết:\n"
                f"{push.stderr}"
            )

    return True, "Đã cập nhật và đăng lên web thành công."

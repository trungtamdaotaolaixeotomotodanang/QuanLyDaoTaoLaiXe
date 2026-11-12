"""Đồng bộ qua GitHub cho app cấp xăng:
- sync_from_remote(): lấy docs/cap_xang_data.json (xe/giáo viên/khóa, từ
  app lái xe) + shared_data/phieu_cap_xang.json, dinh_muc_xang.json (máy
  khác đã đẩy lên) mới nhất về.
- push_shared_data(): đẩy phiếu cấp xăng/định mức của MÁY NÀY lên - chủ
  động (nút "⬆ Đẩy phiếu/định mức lên"), không tự động mỗi lần gõ 1 ô, vì
  ghi phiếu có thể rất nhiều lần/ngày.

Cùng cơ chế/lý do với quanlygiaovien/git_publish.py - xem file đó để biết
chi tiết (--autostash, vì sao cần _CREATE_NO_WINDOW khi đóng gói .exe...)."""
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


FILES_TO_PUSH = [
    "shared_data/phieu_cap_xang.json",
    "shared_data/dinh_muc_xang.json",
]


def push_shared_data(commit_message):
    """Đẩy phiếu cấp xăng/định mức nhiên liệu của máy này lên GitHub - trả
    về (ok: bool, message: str)."""
    existing = [f for f in FILES_TO_PUSH if (REPO_ROOT / f).exists()]
    if not existing:
        return False, "Chưa có phiếu/định mức nào để đẩy lên."

    add = _run(["add", "-A", "--", *existing])
    if add.returncode != 0:
        return False, f"Lỗi khi 'git add':\n{add.stderr}"

    status = _run(["status", "--porcelain", *existing])
    if not status.stdout.strip():
        return True, "Không có gì thay đổi so với lần đẩy lên trước."

    commit = _run(["commit", "-m", commit_message])
    if commit.returncode != 0:
        return False, f"Lỗi khi 'git commit':\n{commit.stderr}"

    # Kéo về + gộp trước khi đẩy lên, để tránh bị từ chối (non-fast-forward)
    # nếu máy khác vừa đẩy phiếu trước đó.
    ok, err = sync_from_remote()
    if not ok:
        return False, err

    push = _run(["push"])
    if push.returncode != 0:
        ok, err = sync_from_remote()
        if not ok:
            return False, err
        push = _run(["push"])
        if push.returncode != 0:
            return False, (
                "Đã lưu ở máy này nhưng KHÔNG đẩy lên GitHub được (có thể "
                "do mất mạng hoặc cần đăng nhập lại). Chi tiết:\n"
                f"{push.stderr}"
            )

    return True, "Đã đẩy phiếu/định mức lên GitHub thành công."

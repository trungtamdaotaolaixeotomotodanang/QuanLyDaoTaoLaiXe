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
# Mạng chậm/rớt (vd wifi yếu, VPN treo) từng làm "git fetch"/"git push" treo
# VÔ THỜI HẠN, kéo theo cả app bị đơ (sync_from_remote() gọi lúc mở app, chưa
# kịp hiện cửa sổ) - đặt timeout để tệ nhất cũng chỉ chờ từng ấy giây rồi báo
# lỗi mạng, không treo app mãi.
_GIT_TIMEOUT_S = 20


class _TimedOutResult:
    """Giả lập subprocess.CompletedProcess khi timeout - để chỗ gọi _run()
    xử lý y như lỗi git bình thường (chỉ cần .returncode/.stdout/.stderr)."""

    def __init__(self, stderr):
        self.returncode = 1
        self.stdout = ""
        self.stderr = stderr


_FRIENDLY_PATTERNS = (
    (("could not resolve host", "unable to access", "could not connect", "network is unreachable"),
     "Không kết nối được tới GitHub - kiểm tra lại mạng."),
    (("permission denied", "authentication failed", "could not read username", "could not read password"),
     "GitHub từ chối đăng nhập - có thể cần đăng nhập lại Git trên máy này."),
    (("timed out", "timeout"), "Kết nối tới GitHub quá chậm hoặc bị treo."),
)


def _friendly_message(action, stderr):
    """Câu dễ hiểu cho người dùng không rành kỹ thuật, kèm thông tin kỹ
    thuật gốc (để báo lại cho người hỗ trợ nếu cần) - xem _FRIENDLY_PATTERNS."""
    low = (stderr or "").lower()
    for keywords, friendly in _FRIENDLY_PATTERNS:
        if any(k in low for k in keywords):
            return f"{friendly}\n\nChi tiết kỹ thuật ('git {action}'):\n{stderr}"
    return f"Lỗi khi 'git {action}':\n{stderr}"


def _run(args):
    try:
        return subprocess.run(
            ["git", *args],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=_CREATE_NO_WINDOW,
            timeout=_GIT_TIMEOUT_S,
        )
    except subprocess.TimeoutExpired:
        return _TimedOutResult(
            f"Lệnh 'git {' '.join(args)}' không phản hồi sau {_GIT_TIMEOUT_S}s "
            "(mạng chậm hoặc mất kết nối)."
        )


def has_unpushed_changes():
    """True nếu phiếu/định mức (FILES_TO_PUSH) đang có thay đổi CHƯA đẩy lên
    GitHub (chưa bấm "⬆ Đẩy phiếu/định mức lên") - dùng nhắc người dùng ở
    main.py, tránh quên bấm nút rồi tưởng máy khác đã thấy dữ liệu mới. Chỉ
    đọc trạng thái cục bộ (không fetch mạng) nên gọi thường xuyên vẫn nhẹ."""
    existing = [f for f in FILES_TO_PUSH if (REPO_ROOT / f).exists()]
    if not existing:
        return False
    status = _run(["status", "--porcelain", *existing])
    return bool(status.stdout.strip())


def sync_from_remote():
    """Kéo về + rebase commit cục bộ (nếu có, bình thường không có vì app
    này không tự tạo commit) lên trên bản mới nhất trên GitHub. Trả về
    (ok, message)."""
    fetch = _run(["fetch", "origin"])
    if fetch.returncode != 0:
        return False, _friendly_message("fetch", fetch.stderr)

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
    "shared_data/dinh_muc_override.json",
]


def push_shared_data(commit_message):
    """Đẩy phiếu cấp xăng/định mức nhiên liệu của máy này lên GitHub - trả
    về (ok: bool, message: str)."""
    existing = [f for f in FILES_TO_PUSH if (REPO_ROOT / f).exists()]
    if not existing:
        return False, "Chưa có phiếu/định mức nào để đẩy lên."

    add = _run(["add", "-A", "--", *existing])
    if add.returncode != 0:
        return False, _friendly_message("add", add.stderr)

    status = _run(["status", "--porcelain", *existing])
    if not status.stdout.strip():
        return True, "Không có gì thay đổi so với lần đẩy lên trước."

    commit = _run(["commit", "-m", commit_message])
    if commit.returncode != 0:
        return False, _friendly_message("commit", commit.stderr)

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

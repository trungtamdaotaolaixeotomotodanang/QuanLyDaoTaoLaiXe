"""Đồng bộ TOÀN BỘ file dữ liệu cục bộ (data.json, các *_overrides.json,
app_settings.json, anh_gv/, qr_codes/, data/, icons/, login_logo.png,
app_icon.ico...) + main.exe mới build từ dist/ sang release/ trước khi nén
release.zip gửi cho khách.

Lý do cần script này: trước đây release/ được cập nhật bằng tay (chỉ copy
main.exe), dễ quên đồng bộ các file dữ liệu khác - đã từng gửi nhầm
release.zip với data.json/ngay_hoc_overrides.json cũ cả ngày, khách không
thấy các ngày bắt đầu/kết thúc đã gán tay trước đó. Chạy script này (sau khi
đã `pyinstaller main.spec --noconfirm`) để không bị quên/sót file nào nữa.

Dùng: python build_release.py
"""
import shutil
from pathlib import Path

from runtime_paths import app_dir

APP_DIR = app_dir()  # quanlygiaovien/ khi chạy bằng "python build_release.py"
REPO_ROOT = APP_DIR.parent
RELEASE_DIR = REPO_ROOT / "release"
DIST_EXE = APP_DIR / "dist" / "main.exe"

# Mọi file/thư mục mà app đọc qua APP_DIR lúc chạy (xem các "APP_DIR / ..."
# trong publish_data.py, login_dialog.py, teacher_card.py, style.py) - phải
# liệt kê đủ ở đây, thiếu file nào thì bản release sẽ thiếu đúng dữ liệu đó.
# Không còn liệt kê lich_hoc/ngay_hoc/ten_gv_overrides.json ở đây - các
# file này giờ nằm cố định ở REPO_ROOT/shared_data/ (không theo APP_DIR
# nữa, xem publish_data.py), nên release/ và quanlygiaovien/ tự dùng
# CHUNG 1 bản duy nhất, không cần đồng bộ qua lại.
FILES = [
    "data.json",
    "app_settings.json",
    "data_sources.json",
    "giao_vien_extra.json",
    "xe_extra.json",
    "login_logo.png",
    "app_icon.ico",
]
DIRS = [
    "data",
    "qr_codes",
    "anh_gv",
    "icons",
]


def main():
    if not DIST_EXE.exists():
        raise SystemExit(
            f"Chưa thấy {DIST_EXE} - chạy `pyinstaller main.spec --noconfirm` trước đã."
        )

    RELEASE_DIR.mkdir(exist_ok=True)
    shutil.copy2(DIST_EXE, RELEASE_DIR / "main.exe")
    print(f"OK main.exe ({DIST_EXE.stat().st_size} bytes)")

    for name in FILES:
        src = APP_DIR / name
        if not src.exists():
            print(f"BO QUA (khong co o nguon): {name}")
            continue
        shutil.copy2(src, RELEASE_DIR / name)
        print(f"OK {name}")

    for name in DIRS:
        src = APP_DIR / name
        if not src.exists():
            print(f"BO QUA (khong co o nguon): {name}/")
            continue
        shutil.copytree(src, RELEASE_DIR / name, dirs_exist_ok=True)
        print(f"OK {name}/")

    print("\nDa dong bo xong release/ - gio nen release/ thanh zip de gui.")


if __name__ == "__main__":
    main()

"""Thư mục thật trên đĩa của app (quanlygiaovien/) - dùng chung cho mọi
module cần tìm docs/, data/, qr_codes/...

Chạy từ source (python main.py): đó là thư mục chứa main.py (__file__).

Đóng gói bằng PyInstaller (--onefile, build ra quanlygiaovien/dist/<ten>.exe):
__file__ sẽ trỏ vào thư mục giải nén tạm bên trong file .exe (sai), nên phải
lấy theo vị trí thật của sys.executable. Nếu .exe đang nằm trong 1 thư mục
tên "dist" (cách build mặc định), lùi thêm 1 cấp để ra đúng quanlygiaovien/.
"""
import sys
from pathlib import Path


def app_dir():
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        return exe_dir.parent if exe_dir.name == "dist" else exe_dir
    return Path(__file__).resolve().parent

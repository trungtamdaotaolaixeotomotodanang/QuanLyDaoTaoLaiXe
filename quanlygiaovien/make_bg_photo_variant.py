"""Ép tạo lại ngay data/khunganhnen_co_anh.jpg (xem
teacher_card._ensure_bg_photo_variant - app tự tạo file này khi cần, BÌNH
THƯỜNG KHÔNG cần chạy tay). Chỉ chạy script này khi muốn thấy kết quả ngay,
ví dụ vừa đổi ảnh nền data/khunganhnen.jpg.
"""
from pathlib import Path

import teacher_card

if teacher_card.BG_PATH_PHOTO.exists():
    Path(teacher_card.BG_PATH_PHOTO).unlink()  # bỏ cache cũ để chắc chắn tạo lại

if teacher_card._ensure_bg_photo_variant():
    print(f"OK: {teacher_card.BG_PATH_PHOTO}")
else:
    print(f"Không tìm thấy {teacher_card.BG_PATH} - chưa có ảnh nền gốc.")

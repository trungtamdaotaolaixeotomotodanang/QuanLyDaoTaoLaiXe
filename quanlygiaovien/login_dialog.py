"""Màn hình đăng nhập - chặn trước khi vào app, hiện lúc khởi động (xem
main.py). Tài khoản cố định (xem _USERNAME/_PASSWORD) - app này chỉ dùng
nội bộ 1 trung tâm, không cần quản lý nhiều người dùng."""
import hashlib

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
)

from runtime_paths import app_dir

APP_DIR = app_dir()
_LOGO_FILE = APP_DIR / "login_logo.png"
_ICON_FILE = APP_DIR / "app_icon.ico"

_USERNAME = "LeLai"
# Chỉ lưu hash (sha256) của mật khẩu trong code, không lưu plaintext - cùng
# cách publish_data.py xử lý ho_so_password_hash, để không bị commit mật
# khẩu dạng chữ thường vào lịch sử git.
_PASSWORD_HASH = "d76ab39c6b8595a53b9d8bd31122f04bce3cc14e8b675298fd4417bc6cb0d963"

_CENTER_NAME = "TRUNG TÂM ĐÀO TẠO LÁI XE Ô TÔ - MÔ TÔ ĐÀ NẴNG"


class LoginDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Đăng nhập")
        self.setFixedSize(400, 500)
        if _ICON_FILE.exists():
            self.setWindowIcon(QIcon(str(_ICON_FILE)))

        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #122a6e, stop:1 #2f6fed
                );
            }
            QFrame#card {
                background: #ffffff;
                border-radius: 16px;
            }
            QLabel#centerName {
                color: #122a6e;
                font-weight: 700;
                font-size: 11pt;
            }
            QLabel#title {
                color: #1f2328;
                font-weight: 700;
                font-size: 16pt;
            }
            QLabel#error {
                color: #b3261e;
                font-size: 9.5pt;
            }
            QLineEdit {
                padding: 9px 10px;
                font-size: 10.5pt;
            }
        """)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(26, 26, 26, 26)

        card = QFrame()
        card.setObjectName("card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(30, 30, 30, 30)
        card_layout.setSpacing(8)

        logo = QLabel()
        logo.setAlignment(Qt.AlignCenter)
        if _LOGO_FILE.exists():
            pix = QPixmap(str(_LOGO_FILE)).scaled(
                92, 92, Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            logo.setPixmap(pix)
        card_layout.addWidget(logo)

        card_layout.addSpacing(4)
        center_name = QLabel(_CENTER_NAME)
        center_name.setObjectName("centerName")
        center_name.setAlignment(Qt.AlignCenter)
        center_name.setWordWrap(True)
        card_layout.addWidget(center_name)

        card_layout.addSpacing(10)
        title = QLabel("ĐĂNG NHẬP")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        card_layout.addWidget(title)
        card_layout.addSpacing(14)

        self.txt_user = QLineEdit()
        self.txt_user.setPlaceholderText("Tên đăng nhập")
        self.txt_user.textEdited.connect(self._clear_error)
        card_layout.addWidget(self.txt_user)

        self.txt_pass = QLineEdit()
        self.txt_pass.setPlaceholderText("Mật khẩu")
        self.txt_pass.setEchoMode(QLineEdit.Password)
        self.txt_pass.textEdited.connect(self._clear_error)
        self.txt_pass.returnPressed.connect(self._on_login)
        card_layout.addWidget(self.txt_pass)

        self.lbl_error = QLabel("")
        self.lbl_error.setObjectName("error")
        self.lbl_error.setAlignment(Qt.AlignCenter)
        self.lbl_error.setWordWrap(True)
        self.lbl_error.setFixedHeight(20)
        card_layout.addWidget(self.lbl_error)

        btn_login = QPushButton("Đăng nhập")
        btn_login.setObjectName("primary")
        btn_login.setCursor(Qt.PointingHandCursor)
        btn_login.clicked.connect(self._on_login)
        card_layout.addWidget(btn_login)

        outer.addWidget(card)

        self.txt_user.setFocus()

    def _clear_error(self):
        self.lbl_error.setText("")

    def _on_login(self):
        user = self.txt_user.text().strip()
        pw_hash = hashlib.sha256(self.txt_pass.text().encode("utf-8")).hexdigest()
        if user == _USERNAME and pw_hash == _PASSWORD_HASH:
            self.accept()
            return
        self.lbl_error.setText("Sai tên đăng nhập hoặc mật khẩu.")
        self.txt_pass.clear()
        self.txt_pass.setFocus()

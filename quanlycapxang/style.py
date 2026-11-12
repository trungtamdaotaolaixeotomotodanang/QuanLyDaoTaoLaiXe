"""Giao diện chung cho toàn app - 1 stylesheet (QSS) áp dụng 1 lần lúc khởi
động (xem main.py: apply_style(app)), không cần sửa từng tab."""
from runtime_paths import app_dir

APP_DIR = app_dir()
# Qt yêu cầu dấu "/" (kể cả trên Windows) trong url() của QSS.
_DROPDOWN_ARROW = (APP_DIR / "icons" / "dropdown_arrow.svg").as_posix()

APP_QSS = """
* {
    font-family: "Segoe UI", "Inter", sans-serif;
    font-size: 10.5pt;
}

QWidget {
    background: #f4f5f7;
    color: #1f2328;
}

QTabWidget::pane {
    border: 1px solid #d9dce1;
    border-radius: 6px;
    background: #ffffff;
    top: -1px;
}

QTabBar::tab {
    background: #e7e9ee;
    color: #4a5160;
    padding: 9px 20px;
    margin-right: 2px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    font-weight: 600;
}

QTabBar::tab:selected {
    background: #ffffff;
    color: #1f2328;
    border: 1px solid #d9dce1;
    border-bottom: none;
}

QTabBar::tab:hover:!selected {
    background: #dde1e8;
}

QGroupBox {
    background: #ffffff;
    border: 1px solid #dde1e8;
    border-radius: 8px;
    margin-top: 14px;
    padding: 14px 10px 10px 10px;
    font-weight: 600;
    color: #2c3340;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    padding: 0 6px;
    color: #3457d5;
}

QLabel {
    background: transparent;
}

QLineEdit, QTextEdit, QListWidget, QTableWidget {
    background: #ffffff;
    border: 1px solid #d0d4db;
    border-radius: 5px;
    padding: 5px 8px;
    selection-background-color: #3457d5;
    selection-color: #ffffff;
}

QLineEdit:focus, QTextEdit:focus {
    border: 1px solid #3457d5;
}

QLineEdit:read-only {
    background: #f4f5f7;
    color: #5b6372;
}

/* Dropdown (combo box) - khung + nút mũi tên rõ ràng, tách biệt hẳn khỏi
   QLineEdit để không bị nhầm là ô nhập chữ thường. */
QComboBox {
    background: #ffffff;
    border: 1px solid #b7bdc8;
    border-radius: 5px;
    padding: 5px 8px;
    padding-right: 30px;
    selection-background-color: #3457d5;
    selection-color: #ffffff;
}

QComboBox:hover {
    border-color: #3457d5;
}

QComboBox:focus {
    border: 1px solid #3457d5;
}

QComboBox:on {
    border: 1px solid #3457d5;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    border-left: 1px solid #b7bdc8;
    border-top-right-radius: 5px;
    border-bottom-right-radius: 5px;
    background: #eef1f8;
}

QComboBox::down-arrow {
    image: url(__DROPDOWN_ARROW__);
    width: 11px;
    height: 11px;
}


QComboBox QAbstractItemView {
    background: #ffffff;
    border: 1px solid #b7bdc8;
    border-radius: 4px;
    padding: 2px;
    outline: none;
    selection-background-color: #3457d5;
    selection-color: #ffffff;
}

QPushButton {
    background: #ffffff;
    border: 1px solid #c7ccd4;
    border-radius: 6px;
    padding: 6px 14px;
    color: #2c3340;
    font-weight: 600;
}

QPushButton:hover {
    background: #eef1f8;
    border-color: #a9b2c3;
}

QPushButton:pressed {
    background: #e2e6ef;
}

QPushButton:disabled {
    color: #a3a9b3;
    background: #f0f1f3;
    border-color: #e1e3e8;
}

QPushButton#primary {
    background: #2f6fed;
    border: 1px solid #2659c4;
    color: #ffffff;
    padding: 9px 18px;
    font-size: 11pt;
}

QPushButton#primary:hover {
    background: #4a82f0;
}

QPushButton#primary:pressed {
    background: #2659c4;
}

QPushButton#danger {
    color: #b3261e;
}

QTableWidget {
    gridline-color: #e7e9ee;
    alternate-background-color: #f7f8fa;
}

QHeaderView::section {
    background: #eef1f8;
    color: #2c3340;
    padding: 6px;
    border: none;
    border-bottom: 1px solid #d9dce1;
    font-weight: 600;
}

QScrollBar:vertical {
    width: 11px;
    background: transparent;
}

QScrollBar::handle:vertical {
    background: #c7ccd4;
    border-radius: 5px;
    min-height: 24px;
}

QScrollBar::handle:vertical:hover {
    background: #a9b2c3;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""


def apply_style(app):
    app.setStyleSheet(APP_QSS.replace("__DROPDOWN_ARROW__", _DROPDOWN_ARROW))


def icon(widget, standard_pixmap):
    """Icon chuẩn của Qt theo style hiện tại - dùng chung cho mọi tab thay vì
    định nghĩa lại hàm này ở từng file."""
    return widget.style().standardIcon(standard_pixmap)

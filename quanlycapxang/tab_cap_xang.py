"""Tab CẤP XĂNG: chọn 1 giáo viên -> bảng bên dưới tự liệt kê sẵn mọi khóa
học/biển số xe giáo viên đó từng được phân công, kèm các phiếu cấp xăng đã
có (sửa trực tiếp trên bảng được) và LUÔN có sẵn 1 dòng trống cuối mỗi nhóm
khóa/xe để cấp tiếp - người dùng chỉ cần gõ SỐ LƯỢNG (các ô khóa học/biển số
đã có sẵn, ngày cấp mặc định hôm nay) là lưu ngay, khỏi phải chọn/gõ lại từ
đầu như làm trên Excel trước đây.

Cùng 1 giáo viên/xe/ngày có thể cấp NHIỀU phiếu cùng lúc (vd đổ nhiều lần
trong ngày) - ô "Số lượng" cho gõ theo kiểu "SỐ PHIẾUxSỐ LÍT", nhiều phiếu
cách nhau bằng dấu phẩy, chỉ chứa số/x/dấu cách/dấu phẩy (xem parse_so_luong).
Vd: "2x30, 4x20, 20, 50" = 2 phiếu 30 lít + 4 phiếu 20 lít + 1 phiếu 20 lít +
1 phiếu 50 lít (8 phiếu, tổng 210 lít) - hiện lại đúng y chuỗi đã gõ, không
quy về 1 số duy nhất, để không mất thông tin từng phiếu.

Chỉ cho "Ngày cấp" nằm trong đúng khoảng thời gian đang học (tu_ngay ->
den_ngay) của khóa đó - sai sẽ báo/không cho gõ.

3 chế độ "Tổng theo khóa học/tuần/tháng" là MA TRẬN ngày có thể nhập liệu
luôn (không chỉ xem) - mỗi hàng là 1 khóa/xe, mỗi cột là 1 ngày trong khóa/
tuần/tháng đang chọn - dùng khi người dùng chỉ nhập dồn vào cuối tuần/tháng/
khóa thay vì nhập hằng ngày. Cột Khóa học/Lớp/Biển số cố định bên trái (không
trôi khi kéo thanh cuộn ngang qua các cột ngày) - xem _render_matrix."""
import calendar
import re
from datetime import date, datetime, timedelta

from PySide6.QtCore import QDate, QRegularExpression, Qt, QTimer
from PySide6.QtGui import QColor, QRegularExpressionValidator
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QCompleter,
    QDateEdit,
    QDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)

import attach_files
import cap_xang_store
from style import icon as _icon

VIEW_MODES = ["Chi tiết (nhập liệu)", "Tổng theo khóa học", "Tổng theo tuần", "Tổng theo tháng"]

GRID_COLUMNS = [
    "Khóa học", "Lớp", "Biển số", "Thời gian học",
    "Ngày cấp", "Số phiếu/Quyển", "Mặt hàng", "Số lượng", "Phân loại sổ", "Ghi chú", "",
]
MATRIX_FROZEN_COLUMNS = ["Khóa học", "Lớp", "Biển số"]
_WEEKDAY_VN = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

_MAU_DONG_TRONG = QColor("#eaf3ff")  # dòng trống cuối mỗi khóa/xe - chỗ gõ số cấp tiếp theo
_MAU_NGOAI_PHAM_VI = QColor("#e4e4e4")  # ô ngày ngoài thời gian học của khóa - không cho gõ
_MAU_CUOI_TUAN = QColor("#fff6e0")  # tô nhạt cột Thứ 7/CN cho dễ nhìn

_SO_LUONG_KY_TU_RE = QRegularExpression(r"^[0-9xX, ]*$")
_SO_LUONG_PHAN_RE = re.compile(r"^(\d+)(?:\s*[xX]\s*(\d+))?$")

_GHI_CHU_SO_LUONG = (
    "Nhập 1 số (vd 50) hoặc SỐ PHIẾUxSỐ LÍT (vd 2x30 = 2 phiếu, mỗi phiếu 30 "
    "lít). Nhiều phiếu cùng ngày cách nhau bằng dấu phẩy.\n"
    "Vd: 2x30, 4x20, 20, 50  =  2 phiếu 30 lít + 4 phiếu 20 lít + 1 phiếu 20 "
    "lít + 1 phiếu 50 lít (8 phiếu, tổng 210 lít)."
)

_GUIDE_CSS = """
<style>
  p { margin: 4px 0 10px 0; }
  b { color: #2b6cb0; }
  code { background: #eef2f7; padding: 1px 4px; border-radius: 3px; }
  ul { margin: 2px 0 10px 18px; padding: 0; }
  li { margin: 4px 0; }
</style>
"""

_HUONG_DAN_SO_LUONG_HTML = """
<p><b>Cách gõ ô "Số lượng"</b> (dùng chung cho mọi chế độ xem):</p>
<ul>
<li>Gõ 1 số, vd <code>50</code> = 1 phiếu 50 lít.</li>
<li>Gõ kiểu <code>SỐ PHIẾUxSỐ LÍT</code>, vd <code>2x30</code> = 2 phiếu, mỗi phiếu 30 lít.</li>
<li>Nhiều phiếu cùng ngày cách nhau bằng dấu phẩy, vd
<code>2x30, 4x20, 20, 50</code> = 2 phiếu 30L + 4 phiếu 20L + 1 phiếu 20L + 1 phiếu
50L (tổng 8 phiếu, 210 lít).</li>
<li>Ô này chỉ gõ được số, chữ <code>x</code>, dấu cách và dấu phẩy - gõ ký tự khác
sẽ không ăn.</li>
<li>Khi hiện lại trên bảng, ô vẫn hiện đúng y chuỗi đã gõ (không rút gọn về 1
số), để không mất thông tin từng phiếu riêng lẻ.</li>
</ul>
"""

_HUONG_DAN_CHI_TIET_HTML = _GUIDE_CSS + f"""
<p><b>Chế độ "Chi tiết (nhập liệu)"</b> - nhập/sửa từng phiếu cấp xăng, dùng khi
cấp xăng xong là ghi ngay trong ngày.</p>
<ul>
<li>Chọn 1 giáo viên ở trên - bảng tự liệt kê <b>mọi khóa học/biển số xe</b>
giáo viên đó từng được phân công, không cần chọn/gõ lại.</li>
<li>Mỗi khóa/xe luôn có sẵn <b>1 dòng nền xanh nhạt</b> ở cuối - đó là chỗ để
cấp phiếu mới: khóa học/lớp/biển số đã điền sẵn, ngày cấp mặc định là hôm nay
(sửa được bằng lịch).</li>
<li>Chỉ cần gõ <b>Số lượng</b> rồi bấm Tab/Enter ra khỏi ô là lưu ngay - không
cần điền các ô khác nếu không cần. Lưu xong, 1 dòng trống mới tự xuất hiện để
cấp tiếp.</li>
<li>Các phiếu đã lưu hiện thành dòng riêng (không tô màu) - sửa bất kỳ ô nào
(ngày, số lượng, mặt hàng, ghi chú...) rồi rời khỏi ô là lưu lại ngay, không
cần nút "Lưu" riêng.</li>
<li>Nút <b>"Xóa"</b> ở cuối mỗi dòng đã lưu để xóa hẳn phiếu đó (có hỏi xác
nhận trước).</li>
<li><b>Ngày cấp bắt buộc nằm trong khoảng thời gian học của khóa đó</b> (từ
ngày khai giảng đến ngày kết thúc) - gõ ngày ngoài khoảng này sẽ bị báo lỗi và
không lưu.</li>
</ul>
{_HUONG_DAN_SO_LUONG_HTML}
"""

_HUONG_DAN_KHOA_HTML = _GUIDE_CSS + f"""
<p><b>Chế độ "Tổng theo khóa học"</b> - bảng kiểu lịch, nhập dồn số liệu cho
<b>1 khóa học cụ thể</b> từ đầu tới cuối khóa trong 1 lần, dùng khi chỉ ghi sổ
lúc khóa học xong (hoặc gần xong).</p>
<ul>
<li>Chọn khóa học cần nhập ở ô <b>"Khóa học:"</b> phía trên bảng (chỉ liệt kê
khóa của giáo viên đang chọn).</li>
<li>Bảng hiện ra: mỗi <b>hàng</b> là 1 xe của khóa đó, mỗi <b>cột</b> là 1
<b>ngày</b> - chạy từ ngày khai giảng đến ngày kết thúc khóa.</li>
<li>Gõ số lượng thẳng vào ô ngày tương ứng - không cần chọn lại khóa/xe/ngày,
chỉ cần gõ số.</li>
<li>Cột <b>Khóa học / Lớp / Biển số</b> cố định ở khung bên trái, không trôi
khi kéo thanh cuộn ngang ở khung bên phải để xem các ngày khác (kéo thanh chia
đôi ở giữa 2 khung để chỉnh độ rộng khung trái nếu muốn).</li>
<li>Khóa <b>chưa có ngày khai giảng/kết thúc</b> sẽ không xem được kiểu lịch
này (báo rõ khi chọn) - dùng chế độ "Chi tiết (nhập liệu)" để nhập khóa đó.</li>
<li>Mặt hàng/Phân loại sổ/Ghi chú/Số phiếu-Quyển của phiếu tạo từ ô ngày sẽ
dùng giá trị gần nhất đã nhập (hoặc để trống) - muốn sửa chi tiết các trường
này thì quay lại chế độ "Chi tiết (nhập liệu)".</li>
</ul>
{_HUONG_DAN_SO_LUONG_HTML}
"""

_HUONG_DAN_TUAN_HTML = _GUIDE_CSS + f"""
<p><b>Chế độ "Tổng theo tuần"</b> - bảng kiểu lịch theo TUẦN, dùng khi chỉ ghi
sổ vào cuối tuần cho tất cả các khóa giáo viên đang dạy.</p>
<ul>
<li>Thanh điều hướng phía trên: <b>"◀ Tuần trước"</b> / <b>"Tuần sau ▶"</b> để
lùi/tới từng tuần, nút <b>"Tuần này"</b> để quay về tuần hiện tại. Mở tab mặc
định hiện tuần hiện tại.</li>
<li>Mỗi <b>hàng</b> là 1 khóa/xe của giáo viên đang chọn (chỉ hiện khóa còn
đang học trong tuần đó, hoặc khóa chưa rõ ngày), mỗi <b>cột</b> là 1 ngày
trong tuần (Thứ 2 → Chủ nhật).</li>
<li>Gõ số lượng thẳng vào ô ngày tương ứng cho từng khóa.</li>
<li>Ô <b>xám</b> = ngày đó ngoài thời gian học của khóa (chưa khai giảng/đã kết
thúc) - không gõ được. Cột <b>Thứ 7/Chủ nhật</b> tô vàng nhạt cho dễ nhìn.</li>
<li>Cột Khóa học/Lớp/Biển số cố định bên trái, giống chế độ "Tổng theo khóa
học".</li>
<li>Mặt hàng/Phân loại sổ/Ghi chú/Số phiếu-Quyển dùng giá trị gần nhất - sửa
chi tiết thì quay lại chế độ "Chi tiết (nhập liệu)".</li>
</ul>
{_HUONG_DAN_SO_LUONG_HTML}
"""

_HUONG_DAN_THANG_HTML = _GUIDE_CSS + f"""
<p><b>Chế độ "Tổng theo tháng"</b> - giống hệt "Tổng theo tuần" nhưng mỗi cột
là 1 ngày trong CẢ THÁNG (28-31 cột) thay vì chỉ 1 tuần - dùng khi chỉ ghi sổ
vào cuối tháng.</p>
<ul>
<li>Thanh điều hướng: <b>"◀ Tháng trước"</b> / <b>"Tháng sau ▶"</b> để lùi/tới
từng tháng, nút <b>"Tháng này"</b> để quay về tháng hiện tại.</li>
<li>Mỗi hàng là 1 khóa/xe còn liên quan tới tháng đó; ô xám = ngoài thời gian
học của khóa; cột cuối tuần tô vàng nhạt; cột Khóa học/Lớp/Biển số cố định bên
trái - y hệt chế độ "Tổng theo tuần".</li>
<li>Vì có nhiều cột hơn (cả tháng), kéo thanh cuộn ngang ở khung bên phải để
xem hết các ngày trong tháng.</li>
</ul>
{_HUONG_DAN_SO_LUONG_HTML}
"""


class _HuongDanDialog(QDialog):
    """Hộp thoại hướng dẫn sử dụng 1 chế độ xem - nội dung HTML cuộn được,
    dùng chung cho cả 4 chế độ (xem _HUONG_DAN_*_HTML)."""

    def __init__(self, title, html, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(640, 520)
        layout = QVBoxLayout(self)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(False)
        browser.setHtml(html)
        layout.addWidget(browser, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.setObjectName("primary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)


def _fmt_so(value):
    if value is None:
        return "—"
    return f"{value:g}"


def _parse_ngay(text):
    try:
        return datetime.strptime((text or "").strip(), "%d/%m/%Y").date()
    except ValueError:
        return None


def parse_so_luong(text):
    """Phân tích chuỗi số lượng kiểu "2x30, 4x20, 20, 50" (xem _GHI_CHU_SO_LUONG
    cho cú pháp) -> trả về (tổng_số_lít, tổng_số_phiếu). Chuỗi rỗng -> (0, 0).
    Có phần sai định dạng -> (None, None)."""
    text = (text or "").strip()
    if not text:
        return 0.0, 0
    tong_lit = 0.0
    tong_phieu = 0
    for phan in text.split(","):
        phan = phan.strip()
        if not phan:
            continue
        m = _SO_LUONG_PHAN_RE.match(phan)
        if not m:
            return None, None
        if m.group(2) is not None:
            so_phieu, lit = int(m.group(1)), int(m.group(2))
        else:
            so_phieu, lit = 1, int(m.group(1))
        tong_lit += so_phieu * lit
        tong_phieu += so_phieu
    return tong_lit, tong_phieu


def _overlaps_period(assignment, start, end):
    """True nếu khóa/xe `assignment` còn "liên quan" tới khoảng [start, end]
    (tuần/tháng đang xem) - khóa chưa có ngày bắt đầu/kết thúc thì luôn coi
    là liên quan (không đủ thông tin để loại)."""
    tu = _parse_ngay(assignment.get("tu_ngay", ""))
    den = _parse_ngay(assignment.get("den_ngay", ""))
    if not tu and not den:
        return True
    if tu and tu > end:
        return False
    if den and den < start:
        return False
    return True


class _EditableCombo(QComboBox):
    """Combobox gõ tự do + gợi ý từ các giá trị đã nhập trước đó (vd "Mặt
    hàng", "Phân loại sổ") - không ép phải chọn trong danh sách."""

    def __init__(self, suggestions, text, parent=None):
        super().__init__(parent)
        self.setEditable(True)
        self.setInsertPolicy(QComboBox.NoInsert)
        self.addItems(suggestions)
        self.setCurrentText(text)
        completer = QCompleter(suggestions, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.setCompleter(completer)


class CapXangTab(QWidget):
    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        today = datetime.now().date()
        self._week_start = today - timedelta(days=today.weekday())
        self._month = (today.year, today.month)
        self._build_ui()
        self._load_teachers()
        self._reload()

    # ---------------- UI ----------------
    def _build_ui(self):
        outer = QVBoxLayout(self)

        box = QGroupBox("Cấp xăng theo giáo viên")
        form_layout = QHBoxLayout(box)
        form_layout.addWidget(QLabel("Giáo viên:"))
        self.cbo_teacher = QComboBox()
        self.cbo_teacher.setEditable(True)
        self.cbo_teacher.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_teacher.lineEdit().setPlaceholderText("Nhập hoặc chọn tên giáo viên...")
        # Chỉ nạp lại bảng khi đã chọn xong tên (Enter/rời ô/chọn từ danh sách)
        # - không nạp lại theo từng ký tự gõ (currentTextChanged), vì bảng có
        # thể khá lớn, gõ theo từng ký tự sẽ giật/chậm.
        self.cbo_teacher.activated.connect(lambda _idx: self._reload())
        self.cbo_teacher.lineEdit().editingFinished.connect(self._reload)
        form_layout.addWidget(self.cbo_teacher, 1)

        form_layout.addSpacing(20)
        form_layout.addWidget(QLabel("Xem theo:"))
        self.cbo_view_mode = QComboBox()
        self.cbo_view_mode.addItems(VIEW_MODES)
        self.cbo_view_mode.currentIndexChanged.connect(lambda _: self._reload())
        form_layout.addWidget(self.cbo_view_mode)
        outer.addWidget(box)

        self.nav_stack = QStackedWidget()
        self.nav_stack.addWidget(self._build_nav_chi_tiet())
        self.nav_stack.addWidget(self._build_nav_khoa())
        self.nav_stack.addWidget(self._build_nav_tuan())
        self.nav_stack.addWidget(self._build_nav_thang())
        outer.addWidget(self.nav_stack)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setWordWrap(True)
        outer.addWidget(self.lbl_summary)

        self.stack = QStackedWidget()

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setSortingEnabled(False)
        self.stack.addWidget(self.table)

        self.stack.addWidget(self._build_matrix_widget())

        outer.addWidget(self.stack, 1)

        self.lbl_total = QLabel("")
        self.lbl_total.setStyleSheet("font-weight: 700; font-size: 12pt; color: #2659c4;")
        outer.addWidget(self.lbl_total)

    def _make_help_button(self, title, html):
        btn = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogHelpButton), "Hướng dẫn")
        btn.setToolTip(f"Xem hướng dẫn sử dụng chế độ \"{title}\"")
        btn.clicked.connect(lambda: _HuongDanDialog(f"Hướng dẫn - {title}", html, self).exec())
        return btn

    def _build_nav_chi_tiet(self):
        w = QWidget()
        layout = QHBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(QLabel("Nhập/sửa từng phiếu cấp xăng trực tiếp trên bảng bên dưới."))
        layout.addStretch(1)
        layout.addWidget(self._make_help_button(VIEW_MODES[0], _HUONG_DAN_CHI_TIET_HTML))
        return w

    def _build_nav_khoa(self):
        w = QWidget()
        layout = QHBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(QLabel("Khóa học:"))
        self.cbo_khoa_matrix = QComboBox()
        self.cbo_khoa_matrix.currentIndexChanged.connect(lambda _: self._reload())
        layout.addWidget(self.cbo_khoa_matrix, 1)
        layout.addWidget(self._make_help_button(VIEW_MODES[1], _HUONG_DAN_KHOA_HTML))
        return w

    def _build_nav_tuan(self):
        w = QWidget()
        layout = QHBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        btn_prev = QPushButton("◀ Tuần trước")
        btn_prev.clicked.connect(self._on_week_prev)
        layout.addWidget(btn_prev)
        self.lbl_tuan = QLabel("")
        self.lbl_tuan.setStyleSheet("font-weight: 600;")
        self.lbl_tuan.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_tuan, 1)
        btn_next = QPushButton("Tuần sau ▶")
        btn_next.clicked.connect(self._on_week_next)
        layout.addWidget(btn_next)
        btn_today = QPushButton("Tuần này")
        btn_today.clicked.connect(self._on_week_today)
        layout.addWidget(btn_today)
        layout.addWidget(self._make_help_button(VIEW_MODES[2], _HUONG_DAN_TUAN_HTML))
        return w

    def _build_nav_thang(self):
        w = QWidget()
        layout = QHBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        btn_prev = QPushButton("◀ Tháng trước")
        btn_prev.clicked.connect(self._on_month_prev)
        layout.addWidget(btn_prev)
        self.lbl_thang = QLabel("")
        self.lbl_thang.setStyleSheet("font-weight: 600;")
        self.lbl_thang.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.lbl_thang, 1)
        btn_next = QPushButton("Tháng sau ▶")
        btn_next.clicked.connect(self._on_month_next)
        layout.addWidget(btn_next)
        btn_today = QPushButton("Tháng này")
        btn_today.clicked.connect(self._on_month_today)
        layout.addWidget(btn_today)
        layout.addWidget(self._make_help_button(VIEW_MODES[3], _HUONG_DAN_THANG_HTML))
        return w

    def _build_matrix_widget(self):
        self.frozen_table = QTableWidget()
        self.frozen_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.frozen_table.verticalHeader().setVisible(False)
        self.frozen_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.frozen_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_table.setSelectionMode(QTableWidget.NoSelection)
        self.frozen_table.setSortingEnabled(False)

        self.date_table = QTableWidget()
        self.date_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.date_table.verticalHeader().setVisible(False)
        self.date_table.setSelectionMode(QTableWidget.NoSelection)
        self.date_table.setSortingEnabled(False)

        self.frozen_table.verticalScrollBar().valueChanged.connect(
            self.date_table.verticalScrollBar().setValue
        )
        self.date_table.verticalScrollBar().valueChanged.connect(
            self.frozen_table.verticalScrollBar().setValue
        )

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self.frozen_table)
        splitter.addWidget(self.date_table)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        return splitter

    def _load_teachers(self):
        teachers = self.model.all_teachers()
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.clear()
        self.cbo_teacher.addItem("")
        self.cbo_teacher.addItems(teachers)
        self.cbo_teacher.setCurrentIndex(0)
        self.cbo_teacher.blockSignals(False)

        completer = QCompleter(teachers, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.cbo_teacher.setCompleter(completer)

    def refresh(self):
        """Nạp lại danh sách giáo viên sau khi tab CẬP NHẬT DỮ LIỆU xử lý
        xong dữ liệu mới - giữ nguyên giáo viên đang chọn nếu còn hợp lệ."""
        teacher_text = self.cbo_teacher.currentText().strip()
        self._load_teachers()
        if teacher_text:
            idx = self.cbo_teacher.findText(teacher_text)
            if idx >= 0:
                self.cbo_teacher.blockSignals(True)
                self.cbo_teacher.setCurrentIndex(idx)
                self.cbo_teacher.blockSignals(False)
        self._reload()

    # ---------------- điều hướng kỳ (tuần/tháng) ----------------
    def _on_week_prev(self):
        self._week_start -= timedelta(days=7)
        self._reload()

    def _on_week_next(self):
        self._week_start += timedelta(days=7)
        self._reload()

    def _on_week_today(self):
        today = datetime.now().date()
        self._week_start = today - timedelta(days=today.weekday())
        self._reload()

    def _on_month_prev(self):
        y, m = self._month
        self._month = (y - 1, 12) if m == 1 else (y, m - 1)
        self._reload()

    def _on_month_next(self):
        y, m = self._month
        self._month = (y + 1, 1) if m == 12 else (y, m + 1)
        self._reload()

    def _on_month_today(self):
        today = datetime.now().date()
        self._month = (today.year, today.month)
        self._reload()

    # ---------------- dựng giao diện theo chế độ ----------------
    def _unique_assignments(self, name):
        """Mỗi khóa/xe giáo viên `name` từng được phân công -> 1 dòng (gộp
        các "ca" trùng khóa/xe lại làm 1, vì phiếu cấp xăng không phân biệt ca)."""
        assignments = {}
        for g in self.model.courses_of_teacher(name):
            key = (g["ma_khoa"], g["xe"])
            assignments.setdefault(key, g)
        return sorted(assignments.values(), key=lambda g: g["ma_khoa"])

    def _find_cell_of_widget(self, table, widget):
        """Tìm (hàng, cột) chứa `widget` (hoặc widget cha của nó, vd ô
        QComboBox chứa QLineEdit bên trong) trong `table` - dùng để nhớ ô
        đang gõ trước khi dựng lại bảng, rồi focus lại đúng ô đó sau khi dựng
        xong (xem _reload) - tránh mất vị trí đang nhập mỗi lần lưu."""
        if widget is None:
            return None
        for r in range(table.rowCount()):
            for c in range(table.columnCount()):
                w = table.cellWidget(r, c)
                if w is not None and (w is widget or w.isAncestorOf(widget)):
                    return r, c
        return None

    def _restore_cell_focus(self, table, pos):
        if not pos:
            return
        r, c = pos
        if r >= table.rowCount() or c >= table.columnCount():
            return
        w = table.cellWidget(r, c)
        if w is None:
            return
        target = w.lineEdit() if hasattr(w, "lineEdit") and w.lineEdit() else w
        target.setFocus()

    def _reload(self, *_args):
        name = self.cbo_teacher.currentText().strip()
        mode = self.cbo_view_mode.currentIndex()

        # Bảng bị dựng lại từ đầu mỗi lần nạp lại (sau khi lưu 1 ô, đổi
        # tuần/tháng...) - tự nhớ lại vị trí cuộn + ô đang gõ trước khi dựng,
        # rồi phục hồi lại sau, để người dùng không bị "nhảy" về đầu bảng sau
        # mỗi lần lưu một ô.
        active_table = self.table if mode == 0 else self.date_table
        focus_pos = self._find_cell_of_widget(active_table, QApplication.focusWidget())
        v_scroll = active_table.verticalScrollBar().value()
        h_scroll = active_table.horizontalScrollBar().value()

        self.nav_stack.setCurrentIndex(mode)
        self.stack.setCurrentIndex(0 if mode == 0 else 1)

        if not name:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self._clear_matrix()
            self.lbl_total.setText("")
            self.lbl_summary.setText("Chọn hoặc nhập tên giáo viên để cấp xăng.")
            return

        name_norm = attach_files.normalize_name(name)
        all_records = cap_xang_store.load()["records"]
        records = [r for r in all_records if attach_files.normalize_name(r.get("giao_vien", "")) == name_norm]

        if mode == 0:
            self._render_grid(name, records)
        elif mode == 1:
            self._render_matrix_khoa(name, records)
        elif mode == 2:
            self._render_matrix_tuan(name, records)
        else:
            self._render_matrix_thang(name, records)

        active_table = self.table if mode == 0 else self.date_table
        # Dựng lại cột/hàng còn vài việc Qt tự xếp lại kích thước ở vòng lặp
        # sự kiện kế tiếp (vd setDefaultSectionSize) có thể đè lại vị trí
        # cuộn vừa set - nên set ngay 1 lần rồi set lại lần 2 sau khi các việc
        # đó đã xử lý xong mới thật sự giữ đúng vị trí.
        active_table.verticalScrollBar().setValue(v_scroll)
        active_table.horizontalScrollBar().setValue(h_scroll)
        self._restore_cell_focus(active_table, focus_pos)
        QTimer.singleShot(0, lambda: self._restore_table_state(active_table, v_scroll, h_scroll, focus_pos))

    def _restore_table_state(self, table, v_scroll, h_scroll, focus_pos):
        table.verticalScrollBar().setValue(v_scroll)
        table.horizontalScrollBar().setValue(h_scroll)
        if QApplication.focusWidget() is None:
            self._restore_cell_focus(table, focus_pos)

    def _render_grid(self, name, records):
        groups = self._unique_assignments(name)

        records_by_key = {}
        for r in records:
            key = (r.get("ma_khoa", ""), r.get("xe", ""))
            records_by_key.setdefault(key, []).append(r)
        for lst in records_by_key.values():
            lst.sort(key=lambda r: (_parse_ngay(r.get("ngay_cap", "")) or datetime.min.date()))

        mat_hang_goi_y = cap_xang_store.distinct_values("mat_hang")
        phan_loai_goi_y = cap_xang_store.distinct_values("phan_loai_so")
        mat_hang_mac_dinh = mat_hang_goi_y[0] if mat_hang_goi_y else ""
        phan_loai_mac_dinh = phan_loai_goi_y[0] if phan_loai_goi_y else ""

        self.table.setRowCount(0)
        self.table.setColumnCount(len(GRID_COLUMNS))
        self.table.setHorizontalHeaderLabels(GRID_COLUMNS)

        if not groups:
            self.table.setRowCount(0)
            self.lbl_total.setText("")
            self.lbl_summary.setText(
                f"Giáo viên '{name}' chưa được phân công khóa học/xe nào trong dữ liệu."
            )
            return
        self.lbl_summary.setText(
            "Gõ số lượng vào dòng trống (nền xanh nhạt) cuối mỗi khóa/xe để cấp xăng - "
            "khóa học/biển số/ngày cấp đã có sẵn, không cần chọn lại. " + _GHI_CHU_SO_LUONG
        )

        row_specs = []  # list[(assignment, record_or_None)]
        for g in groups:
            key = (g["ma_khoa"], g["xe"])
            for r in records_by_key.get(key, []):
                row_specs.append((g, r))
            row_specs.append((g, None))

        self.table.setRowCount(len(row_specs))
        self.table.horizontalHeaderItem(7).setToolTip(_GHI_CHU_SO_LUONG)
        tong_so_luong = 0.0
        for row_idx, (g, r) in enumerate(row_specs):
            thoi_gian = f"{g.get('tu_ngay') or '?'} → {g.get('den_ngay') or '?'}"
            for col, text in enumerate([g["ma_khoa"], g["lop"], g["xe"], thoi_gian]):
                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                if r is None:
                    item.setBackground(_MAU_DONG_TRONG)
                self.table.setItem(row_idx, col, item)

            date_edit = QDateEdit()
            date_edit.setCalendarPopup(True)
            date_edit.setDisplayFormat("dd/MM/yyyy")
            ngay = _parse_ngay(r["ngay_cap"]) if r else None
            date_edit.setDate(QDate(ngay.year, ngay.month, ngay.day) if ngay else QDate.currentDate())
            self.table.setCellWidget(row_idx, 4, date_edit)

            txt_so_phieu = QLineEdit(r.get("so_phieu", "") if r else "")
            self.table.setCellWidget(row_idx, 5, txt_so_phieu)

            cbo_mat_hang = _EditableCombo(
                mat_hang_goi_y, r.get("mat_hang", "") if r else mat_hang_mac_dinh
            )
            self.table.setCellWidget(row_idx, 6, cbo_mat_hang)

            txt_so_luong = QLineEdit()
            txt_so_luong.setValidator(QRegularExpressionValidator(_SO_LUONG_KY_TU_RE))
            txt_so_luong.setPlaceholderText("vd: 2x30, 4x20, 20, 50")
            txt_so_luong.setToolTip(_GHI_CHU_SO_LUONG)
            if r:
                txt_so_luong.setText(r.get("so_luong_text") or _fmt_so(r.get("so_luong")))
                tong_so_luong += r.get("so_luong") or 0
            self.table.setCellWidget(row_idx, 7, txt_so_luong)

            cbo_phan_loai = _EditableCombo(
                phan_loai_goi_y, r.get("phan_loai_so", "") if r else phan_loai_mac_dinh
            )
            self.table.setCellWidget(row_idx, 8, cbo_phan_loai)

            txt_ghi_chu = QLineEdit(r.get("ghi_chu", "") if r else "")
            self.table.setCellWidget(row_idx, 9, txt_ghi_chu)

            widgets = {
                "date": date_edit,
                "so_phieu": txt_so_phieu,
                "mat_hang": cbo_mat_hang,
                "so_luong": txt_so_luong,
                "phan_loai_so": cbo_phan_loai,
                "ghi_chu": txt_ghi_chu,
            }
            commit = self._make_commit_handler(name, g, r, widgets)
            date_edit.editingFinished.connect(commit)
            txt_so_phieu.editingFinished.connect(commit)
            cbo_mat_hang.lineEdit().editingFinished.connect(commit)
            txt_so_luong.editingFinished.connect(commit)
            cbo_phan_loai.lineEdit().editingFinished.connect(commit)
            txt_ghi_chu.editingFinished.connect(commit)

            if r is not None:
                btn_xoa = QPushButton("Xóa")
                btn_xoa.clicked.connect(lambda _checked=False, rid=r["id"]: self._on_delete(rid))
                self.table.setCellWidget(row_idx, 10, btn_xoa)
            else:
                placeholder = QTableWidgetItem("")
                placeholder.setFlags(placeholder.flags() & ~Qt.ItemIsEditable)
                placeholder.setBackground(_MAU_DONG_TRONG)
                self.table.setItem(row_idx, 10, placeholder)

        self.lbl_total.setText(f"Tổng cộng đã cấp: {_fmt_so(tong_so_luong)} lít")

    def _make_commit_handler(self, name, assignment, record, widgets):
        def commit():
            so_luong_text = widgets["so_luong"].text().strip()
            if record is None and not so_luong_text:
                return  # dòng trống chưa gõ gì - bỏ qua, không tạo phiếu rỗng

            tong_lit, tong_phieu = parse_so_luong(so_luong_text)
            if tong_lit is None:
                QMessageBox.warning(self, "Số lượng không đúng định dạng", _GHI_CHU_SO_LUONG)
                return
            if record is None and tong_lit <= 0:
                return  # gõ toàn dấu phẩy/khoảng trắng - chưa có số thật, bỏ qua

            ngay_cap = widgets["date"].date().toPython()
            tu = _parse_ngay(assignment.get("tu_ngay", ""))
            den = _parse_ngay(assignment.get("den_ngay", ""))
            if tu and ngay_cap < tu:
                QMessageBox.warning(
                    self, "Ngày cấp không hợp lệ",
                    f"Ngày cấp ({ngay_cap.strftime('%d/%m/%Y')}) trước ngày khóa "
                    f"'{assignment['ma_khoa']}' bắt đầu học ({assignment['tu_ngay']}).",
                )
                return
            if den and ngay_cap > den:
                QMessageBox.warning(
                    self, "Ngày cấp không hợp lệ",
                    f"Ngày cấp ({ngay_cap.strftime('%d/%m/%Y')}) sau ngày khóa "
                    f"'{assignment['ma_khoa']}' kết thúc học ({assignment['den_ngay']}).",
                )
                return

            fields = {
                "giao_vien": name,
                "ma_khoa": assignment["ma_khoa"],
                "lop": assignment["lop"],
                "xe": assignment["xe"],
                "ngay_cap": ngay_cap.strftime("%d/%m/%Y"),
                "so_phieu": widgets["so_phieu"].text().strip(),
                "mat_hang": widgets["mat_hang"].currentText().strip(),
                "so_luong": tong_lit,
                "so_luong_text": so_luong_text,
                "so_phieu_count": tong_phieu,
                "ghi_chu": widgets["ghi_chu"].text().strip(),
                "phan_loai_so": widgets["phan_loai_so"].currentText().strip(),
            }
            if record is None:
                cap_xang_store.add_record(fields)
            else:
                cap_xang_store.update_record(record["id"], fields)
            # Dựng lại bảng ngay sau khi tín hiệu hiện tại xử lý xong - tránh
            # xóa widget đang phát tín hiệu ngay trong lúc nó chưa chạy xong.
            QTimer.singleShot(0, self._reload)

        return commit

    def _on_delete(self, record_id):
        confirm = QMessageBox.question(self, "Xác nhận", "Xóa phiếu cấp xăng này?")
        if confirm != QMessageBox.StandardButton.Yes:
            return
        cap_xang_store.delete_record(record_id)
        self._reload()

    # ---------------- ma trận ngày (theo khóa/tuần/tháng) ----------------
    def _clear_matrix(self):
        self.frozen_table.setRowCount(0)
        self.frozen_table.setColumnCount(0)
        self.date_table.setRowCount(0)
        self.date_table.setColumnCount(0)

    def _populate_khoa_selector(self, name):
        current = self.cbo_khoa_matrix.currentData()
        by_khoa = {}
        for g in self.model.courses_of_teacher(name):
            by_khoa.setdefault(g["ma_khoa"], g)
        items = sorted(by_khoa.values(), key=lambda g: g["ma_khoa"])

        self.cbo_khoa_matrix.blockSignals(True)
        self.cbo_khoa_matrix.clear()
        for g in items:
            label = f"{g['ma_khoa']} ({g['lop']})"
            if g.get("tu_ngay") and g.get("den_ngay"):
                label += f" [{g['tu_ngay']} → {g['den_ngay']}]"
            else:
                label += " [chưa có ngày bắt đầu/kết thúc]"
            self.cbo_khoa_matrix.addItem(label, g["ma_khoa"])
        if current is not None:
            idx = self.cbo_khoa_matrix.findData(current)
            if idx >= 0:
                self.cbo_khoa_matrix.setCurrentIndex(idx)
        self.cbo_khoa_matrix.blockSignals(False)

    def _render_matrix_khoa(self, name, records):
        self._populate_khoa_selector(name)
        ma_khoa = self.cbo_khoa_matrix.currentData()
        if not ma_khoa:
            self._clear_matrix()
            self.lbl_total.setText("")
            self.lbl_summary.setText(
                f"Giáo viên '{name}' chưa được phân công khóa học/xe nào trong dữ liệu."
            )
            return

        rows_spec = [g for g in self._unique_assignments(name) if g["ma_khoa"] == ma_khoa]
        tu = _parse_ngay(rows_spec[0].get("tu_ngay", "")) if rows_spec else None
        den = _parse_ngay(rows_spec[0].get("den_ngay", "")) if rows_spec else None
        if not tu or not den:
            self._clear_matrix()
            self.lbl_total.setText("")
            self.lbl_summary.setText(
                f"Khóa '{ma_khoa}' chưa có ngày bắt đầu/kết thúc nên không hiển thị được "
                "theo ngày - dùng chế độ 'Chi tiết (nhập liệu)' để nhập."
            )
            return

        dates = [tu + timedelta(days=i) for i in range((den - tu).days + 1)]
        self._render_matrix(name, records, rows_spec, dates)

    def _render_matrix_tuan(self, name, records):
        monday = self._week_start
        sunday = monday + timedelta(days=6)
        iso_year, iso_week, _ = monday.isocalendar()
        self.lbl_tuan.setText(
            f"Tuần {iso_week}/{iso_year}  ({monday.strftime('%d/%m/%Y')} → {sunday.strftime('%d/%m/%Y')})"
        )
        dates = [monday + timedelta(days=i) for i in range(7)]
        rows_spec = [g for g in self._unique_assignments(name) if _overlaps_period(g, monday, sunday)]
        self._render_matrix(name, records, rows_spec, dates)

    def _render_matrix_thang(self, name, records):
        y, m = self._month
        self.lbl_thang.setText(f"Tháng {m:02d}/{y}")
        n_days = calendar.monthrange(y, m)[1]
        start, end = date(y, m, 1), date(y, m, n_days)
        dates = [date(y, m, d) for d in range(1, n_days + 1)]
        rows_spec = [g for g in self._unique_assignments(name) if _overlaps_period(g, start, end)]
        self._render_matrix(name, records, rows_spec, dates)

    def _render_matrix(self, name, records, rows_spec, dates):
        if not rows_spec or not dates:
            self._clear_matrix()
            self.lbl_total.setText("")
            if not rows_spec:
                self.lbl_summary.setText(
                    f"Giáo viên '{name}' không có khóa/xe nào liên quan tới kỳ đang xem."
                )
            return
        self.lbl_summary.setText(
            "Gõ số lượng vào từng ô ngày để cấp xăng cho khóa/xe đó (để trống nếu ngày đó "
            "không cấp) - ô xám là ngoài thời gian học của khóa, không gõ được. "
            "Mặt hàng/Phân loại sổ dùng giá trị gần nhất, sửa chi tiết ở chế độ 'Chi tiết "
            "(nhập liệu)'. " + _GHI_CHU_SO_LUONG
        )

        rec_by_key = {}
        for r in records:
            key = (r.get("ma_khoa", ""), r.get("xe", ""), r.get("ngay_cap", ""))
            rec_by_key.setdefault(key, []).append(r)

        mat_hang_goi_y = cap_xang_store.distinct_values("mat_hang")
        phan_loai_goi_y = cap_xang_store.distinct_values("phan_loai_so")
        mat_hang_mac_dinh = mat_hang_goi_y[0] if mat_hang_goi_y else ""
        phan_loai_mac_dinh = phan_loai_goi_y[0] if phan_loai_goi_y else ""

        self.frozen_table.setRowCount(0)
        self.frozen_table.setColumnCount(len(MATRIX_FROZEN_COLUMNS))
        self.frozen_table.setHorizontalHeaderLabels(MATRIX_FROZEN_COLUMNS)
        self.frozen_table.setRowCount(len(rows_spec))

        self.date_table.setRowCount(0)
        self.date_table.setColumnCount(len(dates))
        self.date_table.setHorizontalHeaderLabels(
            [f"{_WEEKDAY_VN[d.weekday()]}\n{d.strftime('%d/%m')}" for d in dates]
        )

        # Cột cố định chỉ có tiêu đề 1 dòng ("Khóa học"...) còn cột ngày có
        # tiêu đề 2 dòng ("T2\n27/11") -> Qt tự tính chiều cao mỗi tiêu đề
        # khác nhau, làm 2 bảng lệch nhau ở hàng đầu. Ép cả 2 header dùng
        # chung 1 chiều cao (lấy chiều cao lớn nhất cần) để luôn thẳng hàng.
        header_h = max(
            self.frozen_table.horizontalHeader().sizeHint().height(),
            self.date_table.horizontalHeader().sizeHint().height(),
        )
        self.frozen_table.horizontalHeader().setFixedHeight(header_h)
        self.date_table.horizontalHeader().setFixedHeight(header_h)
        self.date_table.setRowCount(len(rows_spec))

        tong_so_luong = 0.0
        tong_so_phieu = 0
        for row_idx, g in enumerate(rows_spec):
            for col, text in enumerate([g["ma_khoa"], g["lop"], g["xe"]]):
                item = QTableWidgetItem(text)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.frozen_table.setItem(row_idx, col, item)

            tu = _parse_ngay(g.get("tu_ngay", ""))
            den = _parse_ngay(g.get("den_ngay", ""))
            for col, d in enumerate(dates):
                trong_pham_vi = (not tu or d >= tu) and (not den or d <= den)
                if not trong_pham_vi:
                    item = QTableWidgetItem("")
                    item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                    item.setBackground(_MAU_NGOAI_PHAM_VI)
                    self.date_table.setItem(row_idx, col, item)
                    continue

                key = (g["ma_khoa"], g["xe"], d.strftime("%d/%m/%Y"))
                matched = rec_by_key.get(key, [])
                cell_text = ", ".join(
                    r.get("so_luong_text") or _fmt_so(r.get("so_luong")) for r in matched
                )
                cell = QLineEdit(cell_text)
                cell.setValidator(QRegularExpressionValidator(_SO_LUONG_KY_TU_RE))
                cell.setToolTip(_GHI_CHU_SO_LUONG)
                cell.setAlignment(Qt.AlignCenter)
                if d.weekday() >= 5:
                    cell.setStyleSheet("background-color: #fff6e0;")
                matched_ids = [r["id"] for r in matched]
                cell.editingFinished.connect(
                    self._make_matrix_commit_handler(
                        name, g, d, matched_ids, cell, mat_hang_mac_dinh, phan_loai_mac_dinh
                    )
                )
                self.date_table.setCellWidget(row_idx, col, cell)
                for r in matched:
                    tong_so_luong += r.get("so_luong") or 0
                    tong_so_phieu += r.get("so_phieu_count") or 1

            h = max(self.date_table.rowHeight(row_idx), self.frozen_table.rowHeight(row_idx))
            self.date_table.setRowHeight(row_idx, h)
            self.frozen_table.setRowHeight(row_idx, h)

        self.frozen_table.resizeColumnsToContents()
        self.date_table.horizontalHeader().setDefaultSectionSize(76)
        self.lbl_total.setText(f"Tổng cộng: {_fmt_so(tong_so_luong)} lít ({tong_so_phieu} phiếu)")

    def _make_matrix_commit_handler(self, name, assignment, ngay, matched_ids, widget,
                                     mat_hang_mac_dinh, phan_loai_mac_dinh):
        def commit():
            text = widget.text().strip()
            tong_lit, tong_phieu = parse_so_luong(text)
            if tong_lit is None:
                QMessageBox.warning(self, "Số lượng không đúng định dạng", _GHI_CHU_SO_LUONG)
                return

            if not text:
                if matched_ids:
                    for rid in matched_ids:
                        cap_xang_store.delete_record(rid)
                    QTimer.singleShot(0, self._reload)
                return

            if len(matched_ids) == 1:
                cap_xang_store.update_record(matched_ids[0], {
                    "so_luong": tong_lit,
                    "so_luong_text": text,
                    "so_phieu_count": tong_phieu,
                })
            else:
                for rid in matched_ids:
                    cap_xang_store.delete_record(rid)
                cap_xang_store.add_record({
                    "giao_vien": name,
                    "ma_khoa": assignment["ma_khoa"],
                    "lop": assignment["lop"],
                    "xe": assignment["xe"],
                    "ngay_cap": ngay.strftime("%d/%m/%Y"),
                    "so_phieu": "",
                    "mat_hang": mat_hang_mac_dinh,
                    "so_luong": tong_lit,
                    "so_luong_text": text,
                    "so_phieu_count": tong_phieu,
                    "ghi_chu": "",
                    "phan_loai_so": phan_loai_mac_dinh,
                })
            # Dựng lại ma trận ngay sau khi tín hiệu hiện tại xử lý xong - tránh
            # xóa widget đang phát tín hiệu ngay trong lúc nó chưa chạy xong.
            QTimer.singleShot(0, self._reload)

        return commit

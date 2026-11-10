"""Tab NGUỒN DỮ LIỆU & CẬP NHẬT WEB: khai báo tất cả nguồn dữ liệu rồi bấm 1
nút để đẩy lên trang web (GitHub Pages).

Nguồn dữ liệu gồm:
- File Excel "BIÊN CHẾ XE KHÓA xxxx.xlsx/.xls" (mỗi khóa 1 file, thêm/xóa được)
- File PDF "LICH HOC THUC HANH xxxx.pdf" (mỗi khóa 1 file, thêm/xóa được)
- File DSGV_*.xml: danh sách TOÀN BỘ giáo viên (nguồn mã QR chính thức)
- File DSXTL_*.xml: danh sách xe tập lái (đối chiếu biển số)
- Thư mục hồ sơ cá nhân giáo viên (.pdf theo tên)

Ảnh đại diện, hạng lái xe dạy, xuất thẻ giáo viên: xem tab DANH SÁCH GIÁO
VIÊN (tab_danh_sach_gv.py). Link giấy tờ xe: xem tab DANH SÁCH XE
(tab_tra_cuu_xe.py).

Người dùng không cần biết git: toàn bộ convert -> JSON công khai ->
git commit/push chạy ngầm khi bấm nút "Cập nhật && Đăng lên web".
"""
from pathlib import Path

from PySide6.QtCore import Qt, QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QCompleter,
    QDialog,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextBrowser,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import publish_data
import git_publish
from style import icon as _icon

# Mô tả định dạng từng loại file nguồn - hiện ngay trong tab để người dùng
# khác (hoặc máy mới lấy code về) biết cần chuẩn bị file đúng cấu trúc nào
# trước khi chọn, khỏi phải đọc code mới hiểu được.
_FORMAT_GUIDE_HTML = """
<style>
  p { margin: 4px 0 10px 0; }
  b { color: #2b6cb0; }
  code { background: #eef2f7; padding: 1px 4px; border-radius: 3px; }
  ul { margin: 2px 0 10px 18px; padding: 0; }
  li { margin: 2px 0; }
</style>

<p><b>1. File Excel "BIÊN CHẾ XE KHÓA xxxx.xlsx/.xls"</b> (mỗi khóa 1 file) - gồm 2 sheet app đọc tới:</p>
<ul>
<li>1 sheet có dòng tiêu đề <code>BIÊN CHẾ XE KHÓA &lt;mã khóa&gt; (&lt;lớp&gt;)</code> ở cột A
(trong 10 dòng đầu). Bảng học viên bắt đầu từ dòng có cột A = <code>Stt</code>, các cột:
A=Stt, B=Họ và tên, C=Ngày sinh (dd/mm/yyyy), D=Địa chỉ thường trú,
E=Xe/GV/SĐT (1 ô gộp 3 dòng: biển số xe [có thể thêm hậu tố
<code>(Ca 1)</code>/<code>(Ca 2)</code> nếu 1 xe chạy nhiều ca] / tên giáo viên / số điện thoại),
F=Ghi chú (vd "Bảo lưu").</li>
<li>Sheet tên <code>"Sổ phân công GV"</code>: dòng tiêu đề có cột A = <code>Số TT</code>,
cột C = "Thời gian" dạng <code>Từ dd/mm/yyyy ... đến dd/mm/yyyy</code>, cột D = mã khóa
(phải khớp mã khóa ở sheet trên) - dùng lấy ngày bắt đầu/kết thúc khóa.</li>
</ul>

<p><b>2. File PDF "LICH HOC THUC HANH xxxx.pdf"</b> (mỗi khóa 1 file) - không yêu cầu
nội dung bên trong, chỉ cần <u>tên file</u> có chứa mã lớp (vd "BK0726") để app tự khớp
với khóa tương ứng.</p>

<p><b>3. File DSGV_*.xml</b> - XML gốc xuất từ hệ thống của Sở GTVT (không tự tạo tay),
cấu trúc <code>&lt;GIAOVIENS&gt;&lt;GIAOVIEN&gt;&lt;MaGV&gt;&lt;HoTenDem&gt;&lt;TenGV&gt;
&lt;DienThoai&gt;&lt;NgaySinh&gt;&lt;NoiCT&gt;&lt;TrangThai&gt;...</code> - là nguồn mã QR
chính thức cho từng giáo viên, <u>bắt buộc phải có</u>.</p>

<p><b>4. File DSXTL_*.xml</b> - XML gốc từ Sở GTVT, cấu trúc
<code>&lt;XETAPS&gt;&lt;XETAP&gt;&lt;BIENSOXE&gt;&lt;NHANHIEU&gt;&lt;LOAIXE&gt;
&lt;TRANGTHAI&gt;...</code> - dùng đối chiếu biển số Excel có khớp danh sách chính thức
không (không bắt buộc, nhưng thiếu thì mất cảnh báo biển số sai).</p>

<p><b>5. Thư mục hồ sơ cá nhân giáo viên</b> - chứa file PDF, chỉ cần <u>tên file</u> chứa
tên giáo viên (không phân biệt hoa/thường, có thể có số thứ tự ở đầu tên file) để app
tự khớp, không yêu cầu đặt tên file theo quy tắc cố định.</p>

<p><b>6. File kết quả thi tốt nghiệp</b> (vd "SỐ LƯỢNG THI ĐẠT TỐT NGHIỆP ...xlsx/.xls",
<u>không bắt buộc</u>) - đọc sheet tên <code>"TỔNGHỢP"</code> (không có thì lấy sheet đầu
tiên), mỗi dòng 1 học viên ĐÃ ĐẠT: A=Stt, B=Họ và tên, C=Ngày sinh, D=Số CMND
(KHÔNG đọc/không đưa lên web), E=Địa chỉ (KHÔNG đọc/không đưa lên web), F=Khóa (mã lớp,
vd "BK0626", có thể kèm <code>, Hạng</code>), G=Hạng, H=Ngày tốt nghiệp (dd/mm/yyyy).
Đối chiếu với học viên bên file biên chế xe theo (tên đã chuẩn hoá + ngày sinh + mã lớp) -
khớp được thì ghi "Đạt" + ngày tốt nghiệp; không khớp coi là <i>chưa có kết quả</i>
(KHÔNG phải trượt).</p>
"""


class _FileListPicker(QGroupBox):
    """Khung: danh sách nhiều file + nút Thêm / Xóa mục chọn."""

    def __init__(self, title, file_filter, parent=None):
        super().__init__(title, parent)
        self.file_filter = file_filter
        layout = QVBoxLayout(self)

        self.lst = QListWidget()
        self.lst.setFixedHeight(90)
        self.lst.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.lst.setToolTip("Bấm đúp vào 1 file để mở lên kiểm tra.")
        self.lst.itemDoubleClicked.connect(self._on_open_item)
        layout.addWidget(self.lst)

        btn_row = QHBoxLayout()
        btn_add = QPushButton(_icon(self, QStyle.StandardPixmap.SP_FileIcon), "Thêm...")
        btn_add.clicked.connect(self._on_add)
        btn_row.addWidget(btn_add)
        btn_remove = QPushButton(_icon(self, QStyle.StandardPixmap.SP_TrashIcon), "Xóa")
        btn_remove.clicked.connect(self._on_remove)
        btn_row.addWidget(btn_remove)
        btn_detail = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Xem chi tiết..."
        )
        btn_detail.setToolTip(
            "Mở bảng danh sách đầy đủ, dễ xem/sort/xóa/mở từng file hơn khung nhỏ này."
        )
        btn_detail.clicked.connect(self._on_detail)
        btn_row.addWidget(btn_detail)
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

    def _on_detail(self):
        dlg = _FileListDetailDialog(self, self)
        dlg.exec()

    def remove_path(self, path):
        for i in range(self.lst.count()):
            if self.lst.item(i).text() == path:
                self.lst.takeItem(i)
                return

    def add_path_if_missing(self, path):
        """Thêm 1 file vào danh sách nếu chưa có - dùng khi nơi khác (vd
        dialog "Thông tin chi tiết" ở tab LỊCH HỌC) vừa đăng ký thêm file
        nguồn, để khung này không bị lệch rồi ghi đè mất lúc người dùng bấm
        "Cập nhật && Đăng lên web" (xem LichHocTab.lich_hoc_source_added)."""
        if path not in self.paths():
            self.lst.addItem(path)

    def _on_add(self):
        existing = self.paths()
        start_dir = str(Path(existing[0]).parent) if existing else ""
        files, _ = QFileDialog.getOpenFileNames(
            self, f"Chọn file - {self.title()}", start_dir, self.file_filter
        )
        if not files:
            return
        existing_set = set(existing)
        for f in files:
            if f not in existing_set:
                self.lst.addItem(f)

    def _on_open_item(self, item):
        path = Path(item.text())
        if not path.exists():
            QMessageBox.warning(self, "Không tìm thấy file", f"File không còn tồn tại:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _on_remove(self):
        items = self.lst.selectedItems()
        if not items:
            return
        ten_file = "\n".join(f"- {Path(i.text()).name}" for i in items)
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Bỏ {len(items)} file khỏi danh sách này?\n{ten_file}\n\n"
            "(Chỉ bỏ khỏi danh sách đang chọn trong app, không xóa file gốc trên máy.)",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for item in items:
            self.lst.takeItem(self.lst.row(item))

    def paths(self):
        return [self.lst.item(i).text() for i in range(self.lst.count())]

    def set_paths(self, paths):
        self.lst.clear()
        for p in paths:
            self.lst.addItem(str(p))


class _FileListDetailDialog(QDialog):
    """Hộp thoại xem đầy đủ danh sách file của 1 _FileListPicker - bảng lớn,
    có scroll, sort được theo cột, mỗi dòng có nút Mở/Xóa riêng, có checkbox
    chọn nhiều dòng để xóa nhanh (hoặc xóa tất cả) - dễ quản lý hơn khung
    danh sách nhỏ ở ngoài khi có nhiều file."""

    def __init__(self, picker, parent=None):
        super().__init__(parent)
        self.picker = picker
        self.setWindowTitle(f"Chi tiết - {picker.title()}")
        self.resize(760, 480)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)

        layout = QVBoxLayout(self)

        select_row = QHBoxLayout()
        self.chk_all = QCheckBox("Chọn tất cả")
        self.chk_all.toggled.connect(self._on_toggle_all)
        select_row.addWidget(self.chk_all)
        select_row.addStretch(1)
        btn_delete_selected = QPushButton("Xóa các mục đã chọn")
        btn_delete_selected.setObjectName("danger")
        btn_delete_selected.clicked.connect(self._on_delete_selected)
        select_row.addWidget(btn_delete_selected)
        btn_delete_all = QPushButton("Xóa tất cả")
        btn_delete_all.setObjectName("danger")
        btn_delete_all.clicked.connect(self._on_delete_all)
        select_row.addWidget(btn_delete_all)
        layout.addLayout(select_row)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["", "Tên file", "Thư mục", ""])
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.table, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.setObjectName("primary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        self._reload()
        self.table.setSortingEnabled(True)

    def _reload(self):
        self.table.setSortingEnabled(False)
        self.chk_all.blockSignals(True)
        self.chk_all.setChecked(False)
        self.chk_all.blockSignals(False)
        paths = self.picker.paths()
        self.table.setRowCount(len(paths))
        for row, p in enumerate(paths):
            path_obj = Path(p)
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            chk_item.setCheckState(Qt.CheckState.Unchecked)
            chk_item.setData(Qt.ItemDataRole.UserRole, p)
            self.table.setItem(row, 0, chk_item)
            self.table.setItem(row, 1, QTableWidgetItem(path_obj.name))
            self.table.setItem(row, 2, QTableWidgetItem(str(path_obj.parent)))
            self.table.setCellWidget(row, 3, self._build_actions(p))
        self.table.setSortingEnabled(True)

    def _build_actions(self, path):
        widget = QWidget()
        h = QHBoxLayout(widget)
        h.setContentsMargins(4, 2, 4, 2)
        btn_open = QPushButton("Mở")
        btn_open.clicked.connect(lambda: self._on_open(path))
        h.addWidget(btn_open)
        btn_del = QPushButton("Xóa")
        btn_del.clicked.connect(lambda: self._on_delete(path))
        h.addWidget(btn_del)
        return widget

    def _checked_paths(self):
        paths = []
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item.checkState() == Qt.CheckState.Checked:
                paths.append(item.data(Qt.ItemDataRole.UserRole))
        return paths

    def _on_toggle_all(self, checked):
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for row in range(self.table.rowCount()):
            self.table.item(row, 0).setCheckState(state)

    def _on_open(self, path):
        p = Path(path)
        if not p.exists():
            QMessageBox.warning(self, "Không tìm thấy file", f"File không còn tồn tại:\n{p}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(p)))

    def _on_delete(self, path):
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Bỏ file khỏi danh sách này?\n{Path(path).name}\n\n"
            "(Chỉ bỏ khỏi danh sách đang chọn trong app, không xóa file gốc trên máy.)",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.picker.remove_path(path)
        self._reload()

    def _on_delete_selected(self):
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "Chưa chọn file nào", "Hãy tick chọn ít nhất 1 file ở cột đầu bảng.")
            return
        ten_file = "\n".join(f"- {Path(p).name}" for p in paths)
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Bỏ {len(paths)} file đã chọn khỏi danh sách này?\n{ten_file}\n\n"
            "(Chỉ bỏ khỏi danh sách đang chọn trong app, không xóa file gốc trên máy.)",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        for p in paths:
            self.picker.remove_path(p)
        self._reload()

    def _on_delete_all(self):
        total = self.table.rowCount()
        if not total:
            return
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Bỏ TẤT CẢ {total} file khỏi danh sách này?\n\n"
            "(Chỉ bỏ khỏi danh sách đang chọn trong app, không xóa file gốc trên máy.)",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self.picker.set_paths([])
        self._reload()


class _SinglePathPicker(QGroupBox):
    """Khung: 1 đường dẫn (file hoặc thư mục) + nút Chọn."""

    def __init__(self, title, mode="file", file_filter="", parent=None):
        super().__init__(title, parent)
        self.mode = mode
        self.file_filter = file_filter
        layout = QHBoxLayout(self)
        self.txt = QLineEdit()
        self.txt.setReadOnly(True)
        layout.addWidget(self.txt, 1)
        btn_open = QPushButton(_icon(self, QStyle.StandardPixmap.SP_FileIcon), "Mở")
        btn_open.setToolTip("Mở file/thư mục này lên để kiểm tra.")
        btn_open.clicked.connect(self._on_open)
        layout.addWidget(btn_open)
        btn = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DirOpenIcon), "Chọn...")
        btn.clicked.connect(self._on_choose)
        layout.addWidget(btn)

    def _on_choose(self):
        if self.mode == "file":
            start_dir = str(Path(self.txt.text()).parent) if self.txt.text() else ""
            path, _ = QFileDialog.getOpenFileName(
                self, f"Chọn file - {self.title()}", start_dir, self.file_filter
            )
        else:
            start_dir = self.txt.text() or ""
            path = QFileDialog.getExistingDirectory(
                self, f"Chọn thư mục - {self.title()}", start_dir
            )
        if path:
            self.txt.setText(path)

    def _on_open(self):
        path_str = self.txt.text().strip()
        if not path_str:
            QMessageBox.information(self, "Chưa chọn", "Chưa chọn file/thư mục nào.")
            return
        path = Path(path_str)
        if not path.exists():
            QMessageBox.warning(self, "Không tìm thấy", f"Không tồn tại:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def path(self):
        return self.txt.text().strip()

    def set_path(self, path):
        self.txt.setText(str(path) if path else "")


class _ActionWorker(QThread):
    """Chạy 1 hành động ngắn (lưu xử lý tên GV - có git push) ở thread riêng
    để không treo UI. fn không nhận tham số, trả về (ok: bool, message: str)."""

    finished_result = Signal(bool, str)

    def __init__(self, fn, parent=None):
        super().__init__(parent)
        self.fn = fn

    def run(self):
        try:
            ok, message = self.fn()
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            ok, message = False, str(exc)
        self.finished_result.emit(ok, message)


class _TeacherPickerDialog(QDialog):
    """Hộp thoại nhỏ: chọn 1 giáo viên trong danh sách DSGV - dùng khi người
    dùng xác nhận 1 tên không khớp là viết tắt/ghi sai của giáo viên nào."""

    def __init__(self, dsgv_teachers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chọn giáo viên đúng")
        self.resize(420, 120)
        self._dsgv_teachers = dsgv_teachers

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Tên này thực ra là của giáo viên nào trong DSGV?"))

        self.cbo = QComboBox()
        self.cbo.setEditable(True)
        self.cbo.setInsertPolicy(QComboBox.NoInsert)
        for ma_gv, ten in dsgv_teachers:
            self.cbo.addItem(f"{ten} ({ma_gv})", ma_gv)
        names = [f"{ten} ({ma_gv})" for ma_gv, ten in dsgv_teachers]
        completer = QCompleter(names, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        self.cbo.setCompleter(completer)
        layout.addWidget(self.cbo)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("Hủy")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_ok = QPushButton("Chọn")
        btn_ok.setObjectName("primary")
        btn_ok.clicked.connect(self._on_accept)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    def _on_accept(self):
        idx = self.cbo.currentIndex()
        if idx < 0 or self.cbo.itemData(idx) is None:
            QMessageBox.warning(self, "Chưa chọn", "Hãy chọn 1 giáo viên trong danh sách.")
            return
        self.accept()

    def selected(self):
        """Trả về (ma_gv, ten) đã chọn, hoặc None nếu không hợp lệ."""
        idx = self.cbo.currentIndex()
        ma_gv = self.cbo.itemData(idx) if idx >= 0 else None
        if not ma_gv:
            return None
        return ma_gv, self.cbo.itemText(idx).rsplit(" (", 1)[0]


class _PublishResultDialog(QDialog):
    """Hộp thoại kết quả "Cập nhật && Đăng lên web" - tách cảnh báo theo
    từng loại thành 1 tab riêng (bảng File/Vị trí/Nội dung, cuộn được) thay
    cho 1 khối text dài trong QMessageBox, để không bị tràn màn hình khi có
    nhiều cảnh báo khác nhau. Với tên giáo viên không khớp/mơ hồ, cho xử lý
    tay ngay (giáo viên đã nghỉ, hoặc chọn đúng tên viết tắt của ai) - xem
    _ACTIONABLE_LABELS."""

    _ACTIONABLE_LABELS = {"Tên giáo viên không khớp", "Tên giáo viên khớp mơ hồ"}

    def __init__(self, result, parent=None):
        super().__init__(parent)
        self.setWindowTitle(
            "Kết quả cập nhật" if result.get("git_ok") else "Hoàn tất (chưa đăng lên web)"
        )
        self.resize(820, 560)
        self._dsgv_teachers = result.get("dsgv_teachers", [])
        self._workers = []
        layout = QVBoxLayout(self)

        summary = QLabel(self._build_summary(result))
        summary.setWordWrap(True)
        layout.addWidget(summary)

        tabs = QTabWidget()
        self._add_overview_tab(tabs, result)
        if result.get("ma_khoa_trung"):
            self._add_ma_khoa_trung_tab(tabs, result["ma_khoa_trung"])
        for label, rows in self._group_warnings(result.get("canh_bao", [])).items():
            self._add_warning_tab(tabs, label, rows)
        layout.addWidget(tabs, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.setObjectName("primary")
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

    @staticmethod
    def _build_summary(result):
        lines = [
            f"Đã xử lý xong lúc {result['generated_at']} "
            f"(từ {result['so_file_excel']} file Excel, {result['so_giao_vien']} giáo viên)."
        ]
        if result.get("git_ok"):
            lines.append("Đã đăng lên trang web thành công.")
        else:
            lines.append("CHƯA đăng lên web được - " + result.get("git_message", ""))
        n_trung = len(result.get("ma_khoa_trung", []))
        if n_trung:
            lines.append(f"Có {n_trung} mã khóa trùng giữa các file (xem tab \"Mã khóa trùng\").")
        n_warn = len(result.get("canh_bao", []))
        if n_warn:
            lines.append(f"Có {n_warn} cảnh báo dữ liệu cần kiểm tra lại (xem các tab dưới).")
        return "\n".join(lines)

    @staticmethod
    def _group_warnings(warnings):
        """Gom cảnh báo theo label -> mỗi label là 1 tab, giữ thứ tự xuất hiện."""
        grouped = {}
        for w in warnings:
            grouped.setdefault(w.get("label", "Khác"), []).append(w)
        return grouped

    def _add_overview_tab(self, tabs, result):
        txt = QTextEdit()
        txt.setReadOnly(True)
        lines = []
        if result.get("giao_vien_moi"):
            lines.append("Giáo viên MỚI (đã tạo mã QR):")
            for name, tid in result["giao_vien_moi"]:
                lines.append(f"  - {name} ({tid}.png)")
        if result.get("giao_vien_doi_ten"):
            lines.append("Đã cập nhật tên in trên mã QR (link/mã QR không đổi):")
            for ten_cu, ten_moi, tid in result["giao_vien_doi_ten"]:
                lines.append(f"  - {tid}: '{ten_cu}' → '{ten_moi}'")
        txt.setPlainText("\n".join(lines) if lines else "Không có thông tin khác.")
        tabs.addTab(txt, "Tổng quan")

    def _add_warning_tab(self, tabs, label, rows):
        actionable = label in self._ACTIONABLE_LABELS and self._dsgv_teachers
        table = QTableWidget()
        table.setColumnCount(4 if actionable else 3)
        headers = ["File (bấm đúp để mở)", "Vị trí", "Nội dung"]
        if actionable:
            headers.append("Xử lý")
        table.setHorizontalHeaderLabels(headers)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(True)
        table.setWordWrap(True)
        table.setToolTip("Bấm đúp vào tên file để mở file Excel gốc lên xem/sửa.")
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        if actionable:
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        table.setRowCount(len(rows))
        for row_idx, w in enumerate(rows):
            file_item = QTableWidgetItem(w.get("file", ""))
            file_path = w.get("file_path", "")
            if file_path:
                file_item.setData(Qt.ItemDataRole.UserRole, file_path)
                file_item.setToolTip(file_path)
            table.setItem(row_idx, 0, file_item)
            table.setItem(row_idx, 1, QTableWidgetItem(w.get("vi_tri", "")))
            table.setItem(row_idx, 2, QTableWidgetItem(w.get("noi_dung", "")))
            if actionable:
                ten_goc = w.get("ten_goc", "")
                table.setCellWidget(row_idx, 3, self._build_action_widget(table, row_idx, ten_goc))
        table.resizeRowsToContents()
        table.setSortingEnabled(True)
        table.cellDoubleClicked.connect(lambda row, col: self._open_cell_file(table, row, 0))
        tabs.addTab(table, f"{label} ({len(rows)})")

    def _add_ma_khoa_trung_tab(self, tabs, rows):
        """Tab riêng cho mã khóa trùng giữa 2 file Excel (xem
        convert_excel_to_json.build_data) - chỉ rõ đúng file/sheet/ô dòng
        tiêu đề của bản được GIỮ LẠI và bản bị BỎ QUA, để người dùng tìm và
        xử lý (xóa file sai/cũ khỏi danh sách "Biên chế xe khóa")."""
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(
            [
                "Mã khóa",
                "File GIỮ LẠI (bấm đúp để mở)", "Vị trí (giữ lại)",
                "File BỊ BỎ QUA (bấm đúp để mở)", "Vị trí (bỏ qua)",
            ]
        )
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setAlternatingRowColors(True)
        table.setToolTip("Bấm đúp vào tên file để mở file Excel gốc lên xem/sửa.")
        header = table.horizontalHeader()
        for col in range(5):
            header.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
        table.setRowCount(len(rows))
        for row_idx, w in enumerate(rows):
            table.setItem(row_idx, 0, QTableWidgetItem(w.get("ma_khoa", "")))

            giu_lai_item = QTableWidgetItem(w.get("giu_lai_file", ""))
            giu_lai_path = w.get("giu_lai_file_path", "")
            if giu_lai_path:
                giu_lai_item.setData(Qt.ItemDataRole.UserRole, giu_lai_path)
                giu_lai_item.setToolTip(giu_lai_path)
            table.setItem(row_idx, 1, giu_lai_item)
            table.setItem(row_idx, 2, QTableWidgetItem(w.get("giu_lai_vi_tri", "")))

            bo_qua_item = QTableWidgetItem(w.get("bo_qua_file", ""))
            bo_qua_path = w.get("bo_qua_file_path", "")
            if bo_qua_path:
                bo_qua_item.setData(Qt.ItemDataRole.UserRole, bo_qua_path)
                bo_qua_item.setToolTip(bo_qua_path)
            table.setItem(row_idx, 3, bo_qua_item)
            table.setItem(row_idx, 4, QTableWidgetItem(w.get("bo_qua_vi_tri", "")))
        table.resizeRowsToContents()
        table.setSortingEnabled(True)
        table.cellDoubleClicked.connect(
            lambda row, col: self._open_cell_file(table, row, 3 if col in (3, 4) else 1)
        )

        note = QLabel(
            "Ghi chú: bấm đúp vào tên file ở cột \"File GIỮ LẠI\" hoặc \"File BỊ "
            "BỎ QUA\" để mở đúng file Excel đó lên xem/sửa."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #6c757d; font-style: italic; padding-top: 4px;")

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.addWidget(table)
        container_layout.addWidget(note)
        tabs.addTab(container, f"Mã khóa trùng ({len(rows)})")

    def _build_action_widget(self, table, row_idx, ten_goc):
        widget = QWidget()
        hl = QHBoxLayout(widget)
        hl.setContentsMargins(2, 2, 2, 2)
        hl.setSpacing(4)
        btn_ignore = QPushButton("GV đã nghỉ")
        btn_ignore.setToolTip("Giáo viên đã nghỉ, không còn trong DSGV - bỏ qua, không báo lại.")
        btn_ignore.clicked.connect(lambda: self._on_mark_ignore(table, row_idx, ten_goc, widget))
        hl.addWidget(btn_ignore)
        btn_pick = QPushButton("Chọn tên đúng...")
        btn_pick.setToolTip("Tên này là viết tắt/ghi sai của giáo viên nào trong DSGV.")
        btn_pick.clicked.connect(lambda: self._on_pick_teacher(table, row_idx, ten_goc, widget))
        hl.addWidget(btn_pick)
        return widget

    def _on_mark_ignore(self, table, row_idx, ten_goc, widget):
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Đánh dấu '{ten_goc}' là giáo viên đã nghỉ (không còn trong DSGV)?\n"
            "Từ lần 'Cập nhật && Đăng lên web' kế tiếp, tên này sẽ không bị báo lỗi nữa.",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._run_override_action(
            widget, lambda: publish_data.set_ten_gv_override(ten_goc, "ignore")
        )

    def _on_pick_teacher(self, table, row_idx, ten_goc, widget):
        dlg = _TeacherPickerDialog(self._dsgv_teachers, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        picked = dlg.selected()
        if not picked:
            return
        ma_gv, ten = picked
        confirm = QMessageBox.question(
            self, "Xác nhận",
            f"Ghi nhận '{ten_goc}' là tên viết tắt/ghi sai của '{ten}'?\n"
            "Từ lần 'Cập nhật && Đăng lên web' kế tiếp, tên này sẽ tự khớp với giáo viên này.",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
        self._run_override_action(
            widget, lambda: publish_data.set_ten_gv_override(ten_goc, ma_gv)
        )

    def _run_override_action(self, widget, fn):
        widget.setEnabled(False)
        worker = _ActionWorker(fn)
        worker.finished_result.connect(lambda ok, msg: self._on_override_done(widget, ok, msg))
        self._workers.append(worker)
        worker.start()

    def _on_override_done(self, widget, ok, message):
        if ok:
            layout = widget.layout()
            while layout.count():
                item = layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()
            layout.addWidget(QLabel("Đã lưu ✓"))
        else:
            widget.setEnabled(True)
            QMessageBox.warning(self, "Chưa lưu được", message)

    def _open_cell_file(self, table, row, col):
        item = table.item(row, col)
        file_path = item.data(Qt.ItemDataRole.UserRole) if item else None
        if not file_path:
            return
        path = Path(file_path)
        if not path.exists():
            QMessageBox.warning(self, "Không tìm thấy file", f"File không còn tồn tại:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))


class _PublishWorker(QThread):
    finished_ok = Signal(dict)
    finished_err = Signal(str)

    def __init__(self, sources, parent=None):
        super().__init__(parent)
        self.sources = sources

    def run(self):
        try:
            result = publish_data.run_pipeline(self.sources)
            ok, message = git_publish.push_public_data(
                f"Cập nhật số liệu giáo viên ({result['generated_at']})"
            )
            result["git_ok"] = ok
            result["git_message"] = message
            self.finished_ok.emit(result)
        except Exception as exc:  # báo lỗi rõ ràng, không để app crash
            self.finished_err.emit(str(exc))


class CapNhatWebTab(QWidget):
    # Phát ra ngay sau khi "Cập nhật && Đăng lên web" ghi lại xong data.json
    # cục bộ (kể cả khi đăng lên web/git thất bại - data.json vẫn đã đổi) -
    # MainWindow lắng nghe để nạp lại dữ liệu cho các tab khác ngay, khỏi
    # phải khởi động lại app.
    data_published = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker = None
        self.last_result = None
        self._build_ui()
        self._load_sources()

    def _build_ui(self):
        # Nội dung tab ngày càng dài (nhiều khung cài đặt) nên bọc trong
        # QScrollArea - tránh bị tràn/đè lên nhau khi cửa sổ không đủ cao.
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        root.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        outer = QVBoxLayout(content)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(10)

        guide_box = QGroupBox("Hướng dẫn định dạng từng file nguồn (bấm để xem/ẩn)")
        guide_box.setCheckable(True)
        guide_box.setChecked(False)
        guide_layout = QVBoxLayout(guide_box)
        self.txt_guide = QTextBrowser()
        self.txt_guide.setOpenExternalLinks(False)
        self.txt_guide.setHtml(_FORMAT_GUIDE_HTML)
        self.txt_guide.setFixedHeight(320)
        self.txt_guide.setVisible(False)
        guide_layout.addWidget(self.txt_guide)
        guide_box.toggled.connect(self.txt_guide.setVisible)
        outer.addWidget(guide_box)

        self.pick_bien_che = _FileListPicker(
            'File Excel "BIÊN CHẾ XE KHÓA ..." (mỗi khóa 1 file)',
            "Excel Files (*.xlsx *.xls)",
        )
        outer.addWidget(self.pick_bien_che)

        self.pick_lich_hoc = _FileListPicker(
            'File PDF "LICH HOC THUC HANH ..." (mỗi khóa 1 file)',
            "PDF Files (*.pdf)",
        )
        outer.addWidget(self.pick_lich_hoc)

        row1 = QHBoxLayout()
        self.pick_dsgv = _SinglePathPicker(
            "Danh sách giáo viên (DSGV_*.xml)", mode="file", file_filter="XML Files (*.xml)"
        )
        row1.addWidget(self.pick_dsgv)
        self.pick_dsxtl = _SinglePathPicker(
            "Danh sách xe tập lái (DSXTL_*.xml)", mode="file", file_filter="XML Files (*.xml)"
        )
        row1.addWidget(self.pick_dsxtl)
        outer.addLayout(row1)

        self.pick_ho_so_folder = _SinglePathPicker(
            "Thư mục hồ sơ cá nhân giáo viên", mode="folder"
        )
        outer.addWidget(self.pick_ho_so_folder)

        self.pick_ket_qua_thi = _SinglePathPicker(
            "File kết quả thi tốt nghiệp (vd \"SỐ LƯỢNG THI ĐẠT TỐT NGHIỆP ...xlsx\") "
            "- không bắt buộc",
            mode="file",
            file_filter="Excel Files (*.xlsx *.xls)",
        )
        outer.addWidget(self.pick_ket_qua_thi)

        pass_box = QGroupBox(
            "Mật khẩu xem hồ sơ cá nhân giáo viên (áp dụng cho mọi giáo viên)"
        )
        pass_layout = QHBoxLayout(pass_box)
        self.txt_password = QLineEdit()
        self.txt_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_password.setPlaceholderText(
            "Để trống = ẩn hoàn toàn mục hồ sơ cá nhân trên trang web"
        )
        settings = publish_data.load_app_settings()
        self.txt_password.setText(settings.get("ho_so_password", ""))
        pass_layout.addWidget(self.txt_password)
        self.chk_show_password = QCheckBox("Hiện mật khẩu")
        self.chk_show_password.setChecked(False)
        self.chk_show_password.toggled.connect(self._on_toggle_show_password)
        pass_layout.addWidget(self.chk_show_password)
        btn_save_password = QPushButton(_icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Lưu")
        btn_save_password.clicked.connect(self._on_save_password)
        pass_layout.addWidget(btn_save_password)
        outer.addWidget(pass_box)

        box = QGroupBox("Đăng số liệu giáo viên lên trang web")
        box_layout = QVBoxLayout(box)
        box_layout.addWidget(
            QLabel(
                "Kiểm tra lại các nguồn dữ liệu ở trên rồi bấm nút dưới đây để\n"
                "đăng số liệu lên trang web. Mã QR của mỗi giáo viên cố định\n"
                "theo mã GV chính thức (DSGV), không đổi giữa các lần cập nhật."
            )
        )
        publish_row = QHBoxLayout()
        self.btn_publish = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_ArrowUp), "Cập nhật && Đăng lên web"
        )
        self.btn_publish.setObjectName("primary")
        self.btn_publish.clicked.connect(self._on_publish_clicked)
        publish_row.addWidget(self.btn_publish, 1)
        self.btn_view_last_result = QPushButton("Xem lại kết quả/cảnh báo...")
        self.btn_view_last_result.setToolTip(
            "Mở lại hộp thoại kết quả + cảnh báo của lần \"Cập nhật && Đăng lên web\" gần "
            "nhất trong phiên làm việc này."
        )
        self.btn_view_last_result.setEnabled(False)
        self.btn_view_last_result.clicked.connect(self._on_view_last_result)
        publish_row.addWidget(self.btn_view_last_result)
        box_layout.addLayout(publish_row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setFixedHeight(160)
        box_layout.addWidget(self.log)

        outer.addWidget(box)
        outer.addStretch(1)

    def _load_sources(self):
        sources = publish_data.load_data_sources()
        self.pick_bien_che.set_paths(sources.get("bien_che_files", []))
        self.pick_lich_hoc.set_paths(sources.get("lich_hoc_files", []))
        self.pick_dsgv.set_path(sources.get("dsgv_file", ""))
        self.pick_dsxtl.set_path(sources.get("dsxtl_file", ""))
        self.pick_ho_so_folder.set_path(sources.get("ho_so_folder", ""))
        self.pick_ket_qua_thi.set_path(sources.get("ket_qua_thi_file", ""))

    def _current_sources(self):
        return {
            "bien_che_files": self.pick_bien_che.paths(),
            "lich_hoc_files": self.pick_lich_hoc.paths(),
            "dsgv_file": self.pick_dsgv.path(),
            "dsxtl_file": self.pick_dsxtl.path(),
            "ho_so_folder": self.pick_ho_so_folder.path(),
            "ket_qua_thi_file": self.pick_ket_qua_thi.path(),
        }

    def _on_toggle_show_password(self, checked):
        self.txt_password.setEchoMode(
            QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
        )

    def _on_save_password(self):
        settings = publish_data.load_app_settings()
        settings["ho_so_password"] = self.txt_password.text().strip()
        publish_data.save_app_settings(settings)
        QMessageBox.information(
            self,
            "Đã lưu",
            "Đã lưu mật khẩu. Bấm 'Cập nhật && Đăng lên web' để áp dụng lên trang web.",
        )

    def _on_publish_clicked(self):
        sources = self._current_sources()
        if not sources["bien_che_files"]:
            QMessageBox.warning(
                self, "Thiếu dữ liệu", "Hãy thêm ít nhất 1 file 'BIÊN CHẾ XE KHÓA ...'."
            )
            return
        if not sources["dsgv_file"]:
            QMessageBox.warning(
                self, "Thiếu dữ liệu", "Hãy chọn file danh sách giáo viên (DSGV_*.xml)."
            )
            return

        publish_data.save_data_sources(sources)

        self.btn_publish.setEnabled(False)
        self.btn_publish.setText("Đang cập nhật...")
        self.log.append("Đang xử lý, vui lòng chờ...")

        self.worker = _PublishWorker(sources)
        self.worker.finished_ok.connect(self._on_done)
        self.worker.finished_err.connect(self._on_error)
        self.worker.start()

    def _on_done(self, result):
        self.btn_publish.setEnabled(True)
        self.btn_publish.setText("Cập nhật && Đăng lên web")
        self.last_result = result
        self.btn_view_last_result.setEnabled(True)

        log_lines = [
            f"Đã xử lý xong lúc {result['generated_at']} "
            f"(từ {result['so_file_excel']} file Excel, {result['so_giao_vien']} giáo viên)."
        ]
        if result.get("canh_bao"):
            log_lines.append(f"Có {len(result['canh_bao'])} cảnh báo - xem chi tiết ở hộp thoại kết quả.")
        log_lines.append(
            "Đã đăng lên trang web thành công." if result["git_ok"] else result["git_message"]
        )
        self.log.append("\n".join(log_lines))

        self.data_published.emit()

        _PublishResultDialog(result, self).exec()

    def _on_view_last_result(self):
        if self.last_result:
            _PublishResultDialog(self.last_result, self).exec()

    def _on_error(self, message):
        self.btn_publish.setEnabled(True)
        self.btn_publish.setText("Cập nhật && Đăng lên web")
        self.log.append(f"LỖI: {message}")
        QMessageBox.critical(self, "Lỗi", message)

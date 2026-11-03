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

from PySide6.QtCore import QThread, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QStyle,
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
        btn_row.addStretch(1)
        layout.addLayout(btn_row)

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
    def __init__(self, parent=None):
        super().__init__(parent)
        self.worker = None
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
        self.btn_publish = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_ArrowUp), "Cập nhật && Đăng lên web"
        )
        self.btn_publish.setObjectName("primary")
        self.btn_publish.clicked.connect(self._on_publish_clicked)
        box_layout.addWidget(self.btn_publish)

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

        lines = [
            f"Đã xử lý xong lúc {result['generated_at']} "
            f"(từ {result['so_file_excel']} file Excel)."
        ]
        lines.append(f"Tổng số giáo viên: {result['so_giao_vien']}.")
        if result["ma_khoa_trung"]:
            lines.append(
                "CẢNH BÁO: các mã khóa bị trùng giữa các file (lấy dữ liệu file sau cùng): "
                + ", ".join(result["ma_khoa_trung"])
            )
        if result.get("canh_bao"):
            lines.append("CẢNH BÁO dữ liệu (kiểm tra lại nguồn):")
            for w in result["canh_bao"]:
                lines.append("  - " + w)
        if result["giao_vien_moi"]:
            lines.append("Giáo viên MỚI (đã tạo mã QR trong thư mục qr_codes/):")
            for name, tid in result["giao_vien_moi"]:
                lines.append(f"  - {name} ({tid}.png)")
        if result.get("giao_vien_doi_ten"):
            lines.append("Đã cập nhật lại tên in trên mã QR (link/mã QR không đổi):")
            for ten_cu, ten_moi, tid in result["giao_vien_doi_ten"]:
                lines.append(f"  - {tid}: '{ten_cu}' -> '{ten_moi}'")
        lines.append(result["git_message"])
        self.log.append("\n".join(lines))

        if result["git_ok"]:
            QMessageBox.information(self, "Hoàn tất", "\n".join(lines))
        else:
            QMessageBox.warning(self, "Chưa đăng lên web được", result["git_message"])

    def _on_error(self, message):
        self.btn_publish.setEnabled(True)
        self.btn_publish.setText("Cập nhật && Đăng lên web")
        self.log.append(f"LỖI: {message}")
        QMessageBox.critical(self, "Lỗi", message)

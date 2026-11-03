import re
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QColor, QDesktopServices
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

import attach_files
import detail_dialog
import publish_data
import roster_xml
from style import icon as _icon

HISTORY_COLUMNS = ["Mã khóa", "Lớp", "Giáo viên", "SĐT", "Thời gian học", "Số HV", "Tình trạng"]

# Màu tô nền cột "Link giấy tờ xe" trong dialog Danh sách chi tiết - để phân
# biệt rõ đây là dữ liệu nhập/cập nhật trong app, không phải lấy từ file XML.
_MAU_COT_LINK = QColor("#fff3b0")


class TraCuuXeTab(QWidget):
    """Tab DANH SÁCH XE: danh sách toàn bộ xe tập lái (từ DSXTL) hiện sẵn
    bên trái để chọn nhanh; bấm vào 1 xe là xem nhãn hiệu/loại xe, toàn bộ
    lịch sử giáo viên/khóa đã dùng xe này, và nhập/sửa + mở link giấy tờ xe
    hiện công khai trên trang web."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        sources = publish_data.load_data_sources()
        dsxtl_file = sources.get("dsxtl_file", "")
        self.dsxtl = (
            roster_xml.parse_dsxtl(dsxtl_file)
            if dsxtl_file and Path(dsxtl_file).exists()
            else {}
        )
        self.vehicles = self._build_vehicle_list()
        self.current_xe_link = ""
        self._build_ui()
        self._load_list()

    def _build_vehicle_list(self):
        """Gộp danh sách xe từ DSXTL (nguồn chính thức) với mọi biển số xuất
        hiện trong các khóa (phòng khi biển số chưa kịp cập nhật vào DSXTL)."""
        vehicles = {}
        for bien_so, info in self.dsxtl.items():
            plate = attach_files.normalize_plate(bien_so)
            vehicles[plate] = {
                "bien_so": bien_so,
                "nhan_hieu": info.get("nhan_hieu", ""),
                "loai_xe": info.get("loai_xe", ""),
                "extra": info.get("extra", {}),
                "in_dsxtl": True,
            }
        for c in self.model.all_courses():
            for g in c.get("groups", []):
                xe = g.get("xe", "")
                if not xe:
                    continue
                plate = attach_files.normalize_plate(xe)
                if plate and plate not in vehicles:
                    vehicles[plate] = {
                        "bien_so": xe,
                        "nhan_hieu": "",
                        "loai_xe": "",
                        "extra": {},
                        "in_dsxtl": False,
                    }
        return dict(sorted(vehicles.items(), key=lambda kv: kv[1]["bien_so"]))

    def _build_ui(self):
        layout = QVBoxLayout(self)
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(main_splitter, 1)

        # Danh sách xe hiện sẵn toàn bộ bên trái (thay vì giấu trong dropdown)
        # để chọn nhanh - giống cách làm ở tab DANH SÁCH GIÁO VIÊN.
        left = QWidget()
        left.setMinimumWidth(220)
        left.setMaximumWidth(320)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_detail_list = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogDetailedView), "Danh sách chi tiết"
        )
        self.btn_detail_list.clicked.connect(self._on_show_detail_list)
        left_layout.addWidget(self.btn_detail_list)

        self.txt_search = QLineEdit()
        self.txt_search.setPlaceholderText("Tìm kiếm theo biển số...")
        self.txt_search.textChanged.connect(self._on_search_changed)
        left_layout.addWidget(self.txt_search)

        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self._on_xe_changed)
        left_layout.addWidget(self.list_widget, 1)

        main_splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)

        top = QHBoxLayout()
        self.lbl_info = QLabel("")
        top.addWidget(self.lbl_info, 1)
        self.btn_open_doc = QPushButton("Mở giấy tờ xe...")
        self.btn_open_doc.setEnabled(False)
        self.btn_open_doc.clicked.connect(self._open_doc_external)
        top.addWidget(self.btn_open_doc)
        right_layout.addLayout(top)

        link_row = QHBoxLayout()
        link_row.addWidget(QLabel("Link giấy tờ xe (hiện trên trang web):"))
        self.txt_xe_link = QLineEdit()
        self.txt_xe_link.setPlaceholderText("https://... (vd trang tra cứu xe của Sở)")
        link_row.addWidget(self.txt_xe_link, 1)
        self.btn_save_xe_link = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Lưu link"
        )
        self.btn_save_xe_link.clicked.connect(self._on_save_xe_link)
        link_row.addWidget(self.btn_save_xe_link)
        right_layout.addLayout(link_row)
        self.txt_xe_link.textChanged.connect(self._show_doc)

        splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget()
        self.table.setColumnCount(len(HISTORY_COLUMNS))
        self.table.setHorizontalHeaderLabels(HISTORY_COLUMNS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        splitter.addWidget(self.table)

        xml_info_box = QWidget()
        xml_info_layout = QVBoxLayout(xml_info_box)
        xml_info_layout.setContentsMargins(0, 0, 0, 0)
        xml_info_layout.addWidget(QLabel("Toàn bộ thông tin từ file DSXTL_*.xml:"))
        self.txt_xml_info = QTextEdit()
        self.txt_xml_info.setReadOnly(True)
        xml_info_layout.addWidget(self.txt_xml_info)
        splitter.addWidget(xml_info_box)

        self.doc_stack = QStackedWidget()

        self.lbl_no_doc = QLabel("Chưa có link giấy tờ xe cho biển số này.")
        self.lbl_no_doc.setAlignment(Qt.AlignCenter)
        self.doc_stack.addWidget(self.lbl_no_doc)

        self.lbl_doc_external = QLabel("")
        self.lbl_doc_external.setAlignment(Qt.AlignCenter)
        self.doc_stack.addWidget(self.lbl_doc_external)

        splitter.addWidget(self.doc_stack)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 2)
        right_layout.addWidget(splitter, 1)

        main_splitter.addWidget(right)
        main_splitter.setStretchFactor(0, 1)
        main_splitter.setStretchFactor(1, 3)

    def _load_list(self, filter_text=""):
        self.list_widget.clear()
        items = list(self.vehicles.items())
        filter_text = (filter_text or "").strip().lower()
        if filter_text:
            items = [
                (plate, v) for plate, v in items
                if filter_text in v["bien_so"].lower() or filter_text in plate.lower()
            ]
        for stt, (plate, v) in enumerate(items, start=1):
            label = v["bien_so"] if v["in_dsxtl"] else f"{v['bien_so']} (chưa có trong DSXTL)"
            item = QListWidgetItem(f"{stt}. {label}")
            item.setData(Qt.ItemDataRole.UserRole, plate)
            self.list_widget.addItem(item)
        if not self.vehicles:
            self.lbl_info.setText(
                "Chưa có danh sách xe - chọn file DSXTL_*.xml ở tab CẬP NHẬT WEB "
                "rồi khởi động lại app."
            )
        elif not items:
            empty_item = QListWidgetItem("Không tìm thấy biển số phù hợp.")
            empty_item.setFlags(Qt.ItemFlag.NoItemFlags)
            self.list_widget.addItem(empty_item)

    def _on_search_changed(self, text):
        self._load_list(text)

    # ---- dialog danh sách chi tiết (toàn bộ xe + toàn bộ thông tin XML) ----
    def _on_show_detail_list(self):
        if not self.vehicles:
            QMessageBox.warning(
                self, "Chưa có DS xe",
                "Hãy chọn file DSXTL_*.xml ở tab CẬP NHẬT WEB rồi khởi động lại app.",
            )
            return
        vehicles = list(self.vehicles.items())

        columns = []
        seen = set()
        for _, v in vehicles:
            for tag in v.get("extra", {}):
                if tag == "DUONGDANANH":
                    continue
                if tag not in seen:
                    seen.add(tag)
                    columns.append(tag)
        columns.append("Link giấy tờ xe")
        link_col = len(columns) - 1

        dialog = QDialog(self)
        dialog.setWindowTitle("Danh sách chi tiết xe")
        dialog.setWindowFlags(
            dialog.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
        )
        dialog.resize(1100, 600)
        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel(
            "Toàn bộ thông tin từ file DSXTL_*.xml. Cột \"Link giấy tờ xe\" (tô vàng) "
            "là dữ liệu đã cập nhật trong app, không nằm trong file XML."
        ))

        filter_row = QHBoxLayout()
        lbl_total = QLabel()
        filter_row.addWidget(lbl_total)
        filter_row.addStretch(1)
        filter_row.addWidget(QLabel("Lọc:"))
        txt_filter = QLineEdit()
        txt_filter.setPlaceholderText("Lọc theo bất kỳ cột...")
        filter_row.addWidget(txt_filter, 1)
        layout.addLayout(filter_row)

        table = QTableWidget()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.setRowCount(len(vehicles))
        for row, (plate, v) in enumerate(vehicles):
            extra = v.get("extra", {})
            for col, tag in enumerate(columns[:-1]):
                table.setItem(row, col, QTableWidgetItem(extra.get(tag, "")))
            link_item = QTableWidgetItem(publish_data.get_xe_link(plate))
            link_item.setBackground(_MAU_COT_LINK)
            table.setItem(row, link_col, link_item)
        table.setSortingEnabled(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addWidget(table, 1)

        col_key = "xe"
        saved_hidden = publish_data.get_detail_dialog_hidden_columns(col_key)
        if saved_hidden is None:
            hidden_columns = detail_dialog.default_hidden_columns(
                columns[:-1], [v.get("extra", {}) for _, v in vehicles],
                always_visible=("Link giấy tờ xe",),
            )
        else:
            hidden_columns = set(saved_hidden) & set(columns)
        detail_dialog.apply_column_visibility(table, columns, hidden_columns)

        def _update_total():
            visible = detail_dialog.filter_table_rows(table, txt_filter.text())
            lbl_total.setText(f"Tổng: {visible}/{len(vehicles)} xe")

        txt_filter.textChanged.connect(_update_total)
        _update_total()

        def _on_choose_columns():
            new_hidden = detail_dialog.choose_visible_columns(dialog, columns, hidden_columns)
            if new_hidden is None:
                return
            hidden_columns.clear()
            hidden_columns.update(new_hidden)
            detail_dialog.apply_column_visibility(table, columns, hidden_columns)
            publish_data.set_detail_dialog_hidden_columns(col_key, sorted(hidden_columns))

        bottom_row = QHBoxLayout()
        btn_columns = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_FileDialogListView), "Chọn cột..."
        )
        btn_columns.clicked.connect(_on_choose_columns)
        bottom_row.addWidget(btn_columns)
        btn_export = QPushButton(
            _icon(self, QStyle.StandardPixmap.SP_DialogSaveButton), "Xuất Excel"
        )
        btn_export.clicked.connect(
            lambda: detail_dialog.export_table_to_excel(
                dialog, table, columns, link_col, "danh_sach_chi_tiet_xe.xlsx", "Xe"
            )
        )
        bottom_row.addWidget(btn_export)
        bottom_row.addStretch(1)
        btn_close = QPushButton("Đóng")
        btn_close.clicked.connect(dialog.accept)
        bottom_row.addWidget(btn_close)
        layout.addLayout(bottom_row)

        dialog.exec()

    def _on_xe_changed(self, current, _previous=None):
        plate = current.data(Qt.ItemDataRole.UserRole) if current else None
        v = self.vehicles.get(plate)
        if not v:
            self.lbl_info.setText("")
            self.txt_xe_link.setText("")
            self.btn_save_xe_link.setEnabled(False)
            self._render_history(None)
            self._show_doc("")
            self.txt_xml_info.clear()
            return

        if v["in_dsxtl"]:
            self.lbl_info.setText(f"Nhãn hiệu: {v['nhan_hieu']}  |  Loại xe: {v['loai_xe']}")
        else:
            self.lbl_info.setText(
                "CẢNH BÁO: biển số này không có trong DSXTL "
                "(kiểm tra lại chính tả biển số trong Excel)."
            )

        self.btn_save_xe_link.setEnabled(True)
        self.txt_xe_link.setText(publish_data.get_xe_link(plate))

        self._render_history(plate)
        self._show_doc(self.txt_xe_link.text())
        self._refresh_xml_info(v)

    def _refresh_xml_info(self, v):
        extra = v.get("extra") or {}
        if not extra:
            self.txt_xml_info.setPlainText("(Biển số này chưa có trong DSXTL_*.xml.)")
            return
        lines = [f"{tag}: {value}" for tag, value in extra.items()]
        self.txt_xml_info.setPlainText("\n".join(lines))

    def _on_save_xe_link(self):
        current = self.list_widget.currentItem()
        plate = current.data(Qt.ItemDataRole.UserRole) if current else None
        if not plate:
            return
        link = self.txt_xe_link.text().strip()
        if link and not re.match(r"^https?://", link, re.IGNORECASE):
            QMessageBox.warning(
                self, "Link không hợp lệ",
                "Link giấy tờ xe phải bắt đầu bằng http:// hoặc https://\n"
                "(để trống nếu muốn xóa link).",
            )
            return
        publish_data.set_xe_link(plate, link)
        QMessageBox.information(
            self, "Đã lưu",
            "Đã lưu link giấy tờ xe. Bấm 'Cập nhật && Đăng lên web' ở tab CẬP NHẬT WEB "
            "để áp dụng lên trang web.",
        )

    def _render_history(self, plate):
        rows = []
        if plate:
            for c in self.model.all_courses():
                for g in c.get("groups", []):
                    if attach_files.normalize_plate(g.get("xe", "")) != plate:
                        continue
                    rows.append(
                        {
                            "ma_khoa": c["ma_khoa"],
                            "lop": c.get("lop", ""),
                            "ca": g.get("ca", ""),
                            "giao_vien": g.get("giao_vien", ""),
                            "sdt": g.get("sdt", ""),
                            "thoi_gian": f"{c.get('tu_ngay', '')} → {c.get('den_ngay', '')}",
                            "so_hoc_vien": str(len(g.get("hoc_vien", []))),
                            "status": self.model.status_of(c),
                        }
                    )

        self.table.setRowCount(len(rows))
        for row_idx, r in enumerate(rows):
            values = [
                r["ma_khoa"],
                f"{r['lop']} ({r['ca']})" if r.get("ca") else r["lop"],
                r["giao_vien"],
                r["sdt"],
                r["thoi_gian"],
                r["so_hoc_vien"],
                r["status"],
            ]
            for col, val in enumerate(values):
                self.table.setItem(row_idx, col, QTableWidgetItem(val))
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(True)

    def _show_doc(self, link):
        link = (link or "").strip()
        self.current_xe_link = link
        is_valid = bool(re.match(r"^https?://", link, re.IGNORECASE))
        self.btn_open_doc.setEnabled(is_valid)
        if not link:
            self.lbl_no_doc.setText("Chưa có link giấy tờ xe cho biển số này.")
            self.doc_stack.setCurrentWidget(self.lbl_no_doc)
        elif not is_valid:
            self.lbl_no_doc.setText(
                "Link giấy tờ xe không hợp lệ (phải bắt đầu bằng http:// hoặc https://)."
            )
            self.doc_stack.setCurrentWidget(self.lbl_no_doc)
        else:
            self.lbl_doc_external.setText(
                'Đã có link giấy tờ xe - bấm "Mở giấy tờ xe..." để xem.'
            )
            self.doc_stack.setCurrentWidget(self.lbl_doc_external)

    def _open_doc_external(self):
        if self.current_xe_link:
            QDesktopServices.openUrl(QUrl(self.current_xe_link))

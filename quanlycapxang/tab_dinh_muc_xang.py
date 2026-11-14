"""Tab ĐỊNH MỨC XĂNG: mặc định hiện TẤT CẢ giáo viên đang dạy, hoặc chọn/gõ
1 giáo viên để lọc riêng - hiện danh sách khóa ĐANG DẠY (ô "Khóa" đổi sang
"Đã đào tạo"/"Tất cả trạng thái" để quyết toán cả khóa đã kết thúc), định mức xăng
(lít/học viên) của xe từng khóa (tra theo file "BẢNG ĐỊNH MỨC NL.xlsx" đã
chọn ở nút "Chọn file định mức..." - xem fuel_norm.py), số học viên tính
xăng (KHÔNG tính học viên "Bảo lưu"), tổng số lít xăng của từng khóa, số lít
ĐÃ CẤP thực tế (cộng dồn các phiếu đã ghi ở tab CẤP XĂNG, khớp theo khóa+xe)
và CÒN LẠI = tổng định mức - đã cấp (tô đỏ nếu âm, tức đã cấp vượt định
mức) - đúng cột "CÒN LẠI" trước đây phải tự tính tay trên Excel, nay app tự
tính + luôn cập nhật theo phiếu mới nhất. + tổng cộng toàn bộ các khóa đang
hiện trong bảng.

Khóa đã cấp VƯỢT định mức: tô cả dòng màu đỏ nhạt + dòng cảnh báo tổng hợp
phía trên bảng liệt kê số khóa vượt/tổng số lít vượt; tick "Chỉ hiện khóa
vượt định mức" để bảng chỉ còn các khóa đó (xử lý nhanh, không phải dò từng
dòng). Nút "Xuất Excel..." xuất đúng bảng đang hiện."""
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QCompleter,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import attach_files
import fuel_norm
from excel_export import ReportWorkbook

RESULT_COLUMNS = [
    "Giáo viên", "Mã khóa", "Lớp", "Xe", "Thời gian học",
    "Số HV tính xăng", "Định mức/HV (lít)", "Tổng (lít)", "Đã cấp (lít)", "Còn lại (lít)",
]
EXPORT_COLUMNS = [
    "Giáo viên", "Mã khóa", "Lớp", "Xe", "Ca", "Từ ngày", "Đến ngày",
    "Số HV tính xăng", "Định mức/HV (lít)", "Tổng định mức (lít)", "Đã cấp (lít)",
    "Còn lại (lít)", "Vượt định mức",
]

ALL_TEACHERS_LABEL = "-- Tất cả giáo viên --"
# Lọc theo trạng thái khóa - mặc định chỉ khóa đang dạy; "Đã đào tạo" để
# quyết toán xăng từng khóa đã kết thúc (trước đây không xem được ở đâu).
STATUS_OPTIONS = ["Đang đào tạo", "Đã đào tạo", "Chưa đào tạo", "Tất cả trạng thái"]
_MAU_DONG_VUOT = QColor("#fde7e7")  # nền cả dòng khi đã cấp vượt định mức
_MAU_CHU_VUOT = QColor("#b3261e")


def _is_vuot(r):
    return r["con_lai"] is not None and r["con_lai"] < 0


class _TeacherNameItem(QTableWidgetItem):
    """Ô cột "Giáo viên" - bấm vào tiêu đề cột để sắp xếp theo TÊN (từ cuối
    cùng), không theo HỌ (mặc định QTableWidgetItem sắp theo nguyên chuỗi
    "Họ Tên", tức theo họ)."""

    def __lt__(self, other):
        return attach_files.given_name_sort_key(self.text()) < attach_files.given_name_sort_key(other.text())


def _fmt_so(value):
    """Hiện số không có số lẻ .0 thừa (vd 112.0 -> "112", 106.5 giữ nguyên)."""
    if value is None:
        return "—"
    return f"{value:g}"


class DinhMucXangTab(QWidget):
    """Tab ĐỊNH MỨC XĂNG: lấy dữ liệu từ DataModel (data.json + dinh_muc_xang.json)."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._build_ui()
        self._load_filters()
        self._update_dinh_muc_file_label()
        self._search_by_teacher(ALL_TEACHERS_LABEL)

    def _update_dinh_muc_file_label(self):
        src = self.model.dinh_muc_source_file
        self.lbl_dinh_muc_file.setText(
            f"Đang dùng: {src}" if src else "Chưa chọn file định mức nhiên liệu nào."
        )

    def _on_pick_dinh_muc_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Chọn file định mức nhiên liệu", "", "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            fuel_norm.load_and_cache(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không đọc được file", str(exc))
            return
        self.model.reload()
        self._update_dinh_muc_file_label()
        self._search_by_teacher()
        QMessageBox.information(
            self, "Hoàn tất",
            "Đã đọc xong file định mức. Bấm \"⬆ Đẩy phiếu/định mức lên\" ở "
            "góc trên để máy khác cũng thấy file này.",
        )

    def _build_ui(self):
        outer = QVBoxLayout(self)

        file_box = QGroupBox("File định mức nhiên liệu")
        file_layout = QHBoxLayout(file_box)
        self.lbl_dinh_muc_file = QLabel("")
        self.lbl_dinh_muc_file.setWordWrap(True)
        file_layout.addWidget(self.lbl_dinh_muc_file, 1)
        btn_pick_dinh_muc = QPushButton("Chọn file định mức...")
        btn_pick_dinh_muc.clicked.connect(self._on_pick_dinh_muc_file)
        file_layout.addWidget(btn_pick_dinh_muc)
        outer.addWidget(file_box)

        box = QGroupBox("Tra cứu định mức xăng theo giáo viên")
        form_layout = QHBoxLayout(box)
        form_layout.addWidget(QLabel("Tên giáo viên:"))
        self.cbo_teacher = QComboBox()
        self.cbo_teacher.setEditable(True)
        self.cbo_teacher.setInsertPolicy(QComboBox.NoInsert)
        self.cbo_teacher.lineEdit().setPlaceholderText("Nhập hoặc chọn tên giáo viên...")
        self.cbo_teacher.currentTextChanged.connect(self._search_by_teacher)
        form_layout.addWidget(self.cbo_teacher, 1)
        form_layout.addSpacing(12)
        form_layout.addWidget(QLabel("Khóa:"))
        self.cbo_status = QComboBox()
        self.cbo_status.addItems(STATUS_OPTIONS)
        self.cbo_status.setToolTip("Chọn \"Đã đào tạo\" để xem/quyết toán xăng các khóa đã kết thúc.")
        self.cbo_status.currentIndexChanged.connect(lambda _: self._search_by_teacher())
        form_layout.addWidget(self.cbo_status)
        form_layout.addSpacing(12)
        self.chk_vuot = QCheckBox("Chỉ hiện khóa vượt định mức")
        self.chk_vuot.toggled.connect(lambda _: self._search_by_teacher())
        form_layout.addWidget(self.chk_vuot)
        self.btn_export = QPushButton("Xuất Excel...")
        self.btn_export.setToolTip("Xuất đúng bảng đang hiện (theo giáo viên/bộ lọc đang chọn) ra file Excel.")
        self.btn_export.clicked.connect(self._on_export)
        form_layout.addWidget(self.btn_export)
        outer.addWidget(box)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setWordWrap(True)
        outer.addWidget(self.lbl_summary)

        self.lbl_vuot = QLabel("")
        self.lbl_vuot.setWordWrap(True)
        self.lbl_vuot.setStyleSheet(
            "color: #b3261e; font-weight: 700; background: #fde7e7; "
            "border: 1px solid #f3b7b7; border-radius: 4px; padding: 6px;"
        )
        self.lbl_vuot.hide()
        outer.addWidget(self.lbl_vuot)

        self.table = QTableWidget()
        self.table.setColumnCount(len(RESULT_COLUMNS))
        self.table.setHorizontalHeaderLabels(RESULT_COLUMNS)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.setSortingEnabled(True)
        outer.addWidget(self.table, 1)

        self.lbl_total = QLabel("")
        self.lbl_total.setStyleSheet("font-weight: 700; font-size: 12pt; color: #2659c4;")
        outer.addWidget(self.lbl_total)

        self.lbl_missing = QLabel("")
        self.lbl_missing.setWordWrap(True)
        self.lbl_missing.setStyleSheet("color: #b3261e;")
        outer.addWidget(self.lbl_missing)

    def _load_filters(self):
        teachers = self.model.all_teachers()
        self.cbo_teacher.blockSignals(True)
        self.cbo_teacher.clear()
        self.cbo_teacher.addItem(ALL_TEACHERS_LABEL)
        self.cbo_teacher.addItems(teachers)
        self.cbo_teacher.setCurrentIndex(0)
        self.cbo_teacher.blockSignals(False)

        completer = QCompleter(teachers, self)
        completer.setCaseSensitivity(Qt.CaseInsensitive)
        completer.setFilterMode(Qt.MatchContains)
        completer.setCompletionMode(QCompleter.PopupCompletion)
        self.cbo_teacher.setCompleter(completer)

    def refresh(self):
        """Nạp lại danh sách giáo viên + tra cứu lại lựa chọn đang chọn (nếu
        có) - gọi sau khi bấm "Lấy dữ liệu mới nhất" ở main.py."""
        teacher_text = self.cbo_teacher.currentText().strip()
        self._load_filters()
        self._update_dinh_muc_file_label()
        if teacher_text and teacher_text != ALL_TEACHERS_LABEL:
            idx = self.cbo_teacher.findText(teacher_text)
            if idx >= 0:
                self.cbo_teacher.blockSignals(True)
                self.cbo_teacher.setCurrentIndex(idx)
                self.cbo_teacher.blockSignals(False)
        self._search_by_teacher(teacher_text or ALL_TEACHERS_LABEL)

    def _search_by_teacher(self, name=None):
        if name is None:
            name = self.cbo_teacher.currentText()
        name = name.strip()
        status = self.cbo_status.currentText()
        statuses = None if status == STATUS_OPTIONS[-1] else (status,)
        status_text = "mọi trạng thái" if statuses is None else status.lower()
        if not name or name == ALL_TEACHERS_LABEL:
            rows = self.model.fuel_rows_for_teacher(None, statuses)
            summary = f"Tất cả giáo viên, khóa {status_text}: {len(rows)} khóa"
        else:
            rows = self.model.fuel_rows_for_teacher(name, statuses)
            summary = f"Giáo viên '{name}', khóa {status_text}: {len(rows)} khóa"

        vuot = [r for r in rows if _is_vuot(r)]
        self._update_vuot_banner(vuot)
        if self.chk_vuot.isChecked():
            rows = vuot
            summary += f" - đang chỉ hiện {len(rows)} khóa vượt định mức"
        self._show_rows(rows, summary)

    def _update_vuot_banner(self, vuot):
        if not vuot:
            self.lbl_vuot.hide()
            return
        tong_vuot = sum(-r["con_lai"] for r in vuot)
        ds = ", ".join(
            f"{r['giao_vien']} - {r['ma_khoa']} ({r['xe']}): vượt {_fmt_so(-r['con_lai'])} lít"
            for r in vuot
        )
        self.lbl_vuot.setText(
            f"⚠ {len(vuot)} khóa đã cấp VƯỢT định mức, tổng vượt {_fmt_so(tong_vuot)} lít: {ds}"
        )
        self.lbl_vuot.show()

    def _show_rows(self, rows, summary_text):
        self._current_rows = rows
        self.table.setSortingEnabled(False)
        self.table.setRowCount(len(rows))
        for row_idx, r in enumerate(rows):
            thoi_gian = f"{r['tu_ngay']} → {r['den_ngay']}"
            values = [
                r["giao_vien"],
                r["ma_khoa"],
                r["lop"],
                f"{r['xe']} ({r['ca']})" if r.get("ca") else r["xe"],
                thoi_gian,
                str(r["so_hoc_vien"]),
                _fmt_so(r["dinh_muc"]),
                _fmt_so(r["tong"]),
                _fmt_so(r["da_cap"]),
                _fmt_so(r["con_lai"]),
            ]
            vuot = _is_vuot(r)
            for col, val in enumerate(values):
                item = _TeacherNameItem(val) if col == 0 else QTableWidgetItem(val)
                if col >= 5:
                    item.setTextAlignment(Qt.AlignCenter)
                if vuot:
                    item.setBackground(_MAU_DONG_VUOT)
                    if col == 9:
                        item.setForeground(_MAU_CHU_VUOT)
                self.table.setItem(row_idx, col, item)
        self.table.setSortingEnabled(True)
        self.lbl_summary.setText(summary_text)

        tong_cong = sum(r["tong"] for r in rows if r["tong"] is not None)
        tong_da_cap = sum(r["da_cap"] for r in rows)
        tong_con_lai = sum(r["con_lai"] for r in rows if r["con_lai"] is not None)
        self.lbl_total.setText(
            f"Tổng cộng: {_fmt_so(tong_cong)} lít xăng ({len(rows)} khóa) - "
            f"đã cấp {_fmt_so(tong_da_cap)} lít - còn lại {_fmt_so(tong_con_lai)} lít"
        )

        thieu = [r for r in rows if r["dinh_muc"] is None]
        if thieu:
            ds = ", ".join(f"{r['xe']} ({r['ma_khoa']})" for r in thieu)
            self.lbl_missing.setText(
                f"Chưa tra được định mức cho {len(thieu)} xe/khóa - cần bổ sung vào "
                f"file định mức nhiên liệu (nút \"Chọn file định mức...\" ở trên): {ds}"
            )
        else:
            self.lbl_missing.setText("")

    # ---------------- xuất Excel ----------------
    def _on_export(self):
        rows = getattr(self, "_current_rows", [])
        if not rows:
            QMessageBox.information(self, "Không có dữ liệu", "Bảng đang trống, không có gì để xuất.")
            return
        default_name = f"DinhMucXang_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất Excel", default_name, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            report = ReportWorkbook()
            report.add_sheet("Dinh muc xang", EXPORT_COLUMNS, [dinh_muc_export_row(r) for r in rows])
            report.save(path)
        except Exception as exc:
            QMessageBox.warning(self, "Không xuất được", str(exc))
            return
        QMessageBox.information(self, "Hoàn tất", f"Đã xuất {len(rows)} khóa ra:\n{path}")


def dinh_muc_export_row(r):
    """1 dòng Excel cho 1 khóa/xe (xem CapXangModel.fuel_rows_for_teacher) -
    dùng chung với sheet "Định mức" trong báo cáo ở tab TỔNG HỢP CẤP XĂNG."""
    return [
        r["giao_vien"], r["ma_khoa"], r["lop"], r["xe"], r.get("ca", ""),
        r["tu_ngay"], r["den_ngay"], r["so_hoc_vien"], r["dinh_muc"], r["tong"],
        r["da_cap"], r["con_lai"], "Có" if _is_vuot(r) else "",
    ]

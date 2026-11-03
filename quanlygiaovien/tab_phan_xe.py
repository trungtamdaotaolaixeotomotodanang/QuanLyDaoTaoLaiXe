from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QSplitter,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)


class PhanXeTab(QWidget):
    """Tab PHÂN XE: chọn khóa, xem danh sách xe/giáo viên và học viên theo từng xe."""

    def __init__(self, model, parent=None):
        super().__init__(parent)
        self.model = model
        self._build_ui()
        self._load_courses()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        top = QHBoxLayout()
        top.addWidget(QLabel("Chọn khóa:"))
        self.cbo_course = QComboBox()
        self.cbo_course.currentIndexChanged.connect(self._on_course_changed)
        top.addWidget(self.cbo_course, 1)
        self.lbl_summary = QLabel("")
        top.addWidget(self.lbl_summary, 2)
        layout.addLayout(top)

        self.tree = QTreeWidget()
        self.tree.setColumnCount(7)
        self.tree.setHeaderLabels(
            [
                "Họ và tên / Xe - GV - SĐT",
                "Ngày sinh",
                "Địa chỉ",
                "Ghi chú",
                "Kết quả thi",
                "Ngày tốt nghiệp",
                "Stt",
            ]
        )
        self.tree.setColumnHidden(6, True)
        self.tree.setAlternatingRowColors(True)
        layout.addWidget(self.tree)

    def _load_courses(self):
        self.cbo_course.clear()
        for c in self.model.all_courses():
            label = f"{c['ma_khoa']} ({c['lop']})"
            self.cbo_course.addItem(label, c["ma_khoa"])

    def _on_course_changed(self, _index):
        ma_khoa = self.cbo_course.currentData()
        self._render_course(ma_khoa)

    def _render_course(self, ma_khoa):
        self.tree.clear()
        course = self.model.find_course(ma_khoa)
        if not course:
            self.lbl_summary.setText("")
            return

        self.lbl_summary.setText(
            f"Lớp {course['lop']} | Thời gian: {course['tu_ngay']} → {course['den_ngay']} | "
            f"Tổng học viên: {course['so_hoc_vien']} | Số xe: {len(course['groups'])}"
        )

        for g in course.get("groups", []):
            header_text = f"{g['giao_vien']}  |  Xe {g['xe']}  |  SĐT {g.get('sdt','')}  ({len(g['hoc_vien'])} HV)"
            top_item = QTreeWidgetItem([header_text, "", "", "", "", "", ""])
            top_item.setFirstColumnSpanned(False)
            font = top_item.font(0)
            font.setBold(True)
            top_item.setFont(0, font)
            self.tree.addTopLevelItem(top_item)

            for hv in g.get("hoc_vien", []):
                child = QTreeWidgetItem(
                    [
                        f"  {hv['stt']}. {hv['ho_ten']}",
                        hv.get("ngay_sinh", ""),
                        hv.get("dia_chi", ""),
                        hv.get("ghi_chu", ""),
                        hv.get("ket_qua_thi", ""),
                        hv.get("ngay_tot_nghiep", ""),
                        str(hv.get("stt", "")),
                    ]
                )
                top_item.addChild(child)

            top_item.setExpanded(True)

        for col in range(6):
            self.tree.resizeColumnToContents(col)

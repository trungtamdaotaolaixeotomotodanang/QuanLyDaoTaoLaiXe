"""Dialog chỉnh sửa ảnh thẻ: trước đây app tự crop giữa ảnh gốc theo tỉ lệ
3x4 rồi kéo giãn vừa khít khung trên thẻ - khi khung không đúng hẳn tỉ lệ
3x4 (hoặc chủ thể không nằm giữa ảnh gốc) thì ảnh bị méo/lệch. Dialog này
cho người dùng tự kéo để chọn vùng + cuộn chuột (hoặc kéo thanh trượt) để
phóng to/thu nhỏ, khung crop luôn cố định đúng tỉ lệ 3:4, rồi xuất ra ảnh đã
crop đúng tỉ lệ để lưu làm ảnh đại diện giáo viên."""
from PIL import Image
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSlider,
    QVBoxLayout,
    QWidget,
)

CROP_RATIO = 3 / 4  # rộng:cao - chuẩn ảnh thẻ 3x4
FRAME_W, FRAME_H = 300, 400  # kích thước khung crop hiển thị trên dialog (px)
OUTPUT_W, OUTPUT_H = 600, 800  # kích thước ảnh xuất ra sau khi crop (giữ nét khi in)
_WORKING_MAX_DIM = 1600  # giới hạn ảnh làm việc để thao tác kéo/phóng mượt


def _pil_to_qpixmap(img):
    img = img.convert("RGB")
    data = img.tobytes("raw", "RGB")
    qimg = QImage(data, img.width, img.height, img.width * 3, QImage.Format.Format_RGB888)
    return QPixmap.fromImage(qimg.copy())


class _CropCanvas(QWidget):
    """Khung crop cố định (FRAME_W x FRAME_H, tỉ lệ 3:4) - ảnh nằm bên dưới,
    kéo chuột để di chuyển, cuộn chuột để phóng to/thu nhỏ. Ảnh luôn được
    giữ phủ kín khung (không bao giờ để lộ khoảng trống ngoài ảnh)."""

    zoomChanged = Signal(float)

    def __init__(self, pixmap, parent=None):
        super().__init__(parent)
        self.setFixedSize(FRAME_W, FRAME_H)
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        self._pixmap = pixmap
        self._min_zoom = max(FRAME_W / pixmap.width(), FRAME_H / pixmap.height())
        self._max_zoom = self._min_zoom * 4
        self._zoom = self._min_zoom
        self._ox = (FRAME_W - pixmap.width() * self._zoom) / 2
        self._oy = (FRAME_H - pixmap.height() * self._zoom) / 2
        self._drag_start = None
        self._drag_origin = None

    def reset(self):
        self._zoom = self._min_zoom
        self._ox = (FRAME_W - self._pixmap.width() * self._zoom) / 2
        self._oy = (FRAME_H - self._pixmap.height() * self._zoom) / 2
        self.update()

    def set_zoom(self, zoom):
        zoom = max(self._min_zoom, min(self._max_zoom, zoom))
        # giữ nguyên điểm đang ở giữa khung khi đổi zoom, tránh ảnh "nhảy"
        cx = (FRAME_W / 2 - self._ox) / self._zoom
        cy = (FRAME_H / 2 - self._oy) / self._zoom
        self._zoom = zoom
        self._ox = FRAME_W / 2 - cx * zoom
        self._oy = FRAME_H / 2 - cy * zoom
        self._clamp()
        self.update()

    def zoom_fraction(self):
        if self._max_zoom <= self._min_zoom:
            return 0.0
        return (self._zoom - self._min_zoom) / (self._max_zoom - self._min_zoom)

    def _clamp(self):
        w = self._pixmap.width() * self._zoom
        h = self._pixmap.height() * self._zoom
        self._ox = min(0.0, max(FRAME_W - w, self._ox))
        self._oy = min(0.0, max(FRAME_H - h, self._oy))

    def wheelEvent(self, event):
        step = self._zoom * 0.15 * (1 if event.angleDelta().y() > 0 else -1)
        self.set_zoom(self._zoom + step)
        self.zoomChanged.emit(self.zoom_fraction())

    def mousePressEvent(self, event):
        if event.button() != Qt.MouseButton.LeftButton:
            return
        self._drag_start = event.position()
        self._drag_origin = (self._ox, self._oy)
        self.setCursor(Qt.CursorShape.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        if self._drag_start is None:
            return
        delta = event.position() - self._drag_start
        self._ox = self._drag_origin[0] + delta.x()
        self._oy = self._drag_origin[1] + delta.y()
        self._clamp()
        self.update()

    def mouseReleaseEvent(self, event):
        self._drag_start = None
        self.setCursor(Qt.CursorShape.OpenHandCursor)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.fillRect(self.rect(), QColor("#000000"))
        scaled = self._pixmap.scaled(
            round(self._pixmap.width() * self._zoom), round(self._pixmap.height() * self._zoom),
            Qt.AspectRatioMode.IgnoreAspectRatio, Qt.TransformationMode.SmoothTransformation,
        )
        painter.drawPixmap(round(self._ox), round(self._oy), scaled)

        # lưới 1/3 hỗ trợ canh bố cục (vd mắt/vai)
        pen = QPen(QColor(255, 255, 255, 110))
        pen.setWidth(1)
        painter.setPen(pen)
        for i in (1, 2):
            x = FRAME_W * i / 3
            painter.drawLine(round(x), 0, round(x), FRAME_H)
            y = FRAME_H * i / 3
            painter.drawLine(0, round(y), FRAME_W, round(y))

        pen = QPen(QColor("#2f6fed"))
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawRect(1, 1, FRAME_W - 2, FRAME_H - 2)

    def crop_box_on_source(self):
        """Toạ độ (x0,y0,x1,y1) của khung crop trên ảnh gốc (chưa scale)."""
        x0 = -self._ox / self._zoom
        y0 = -self._oy / self._zoom
        x1 = (FRAME_W - self._ox) / self._zoom
        y1 = (FRAME_H - self._oy) / self._zoom
        return x0, y0, x1, y1


class PhotoCropDialog(QDialog):
    """Dialog chỉnh sửa ảnh thẻ - kéo/phóng để chọn đúng vùng ảnh (tỉ lệ
    3x4) trước khi lưu làm ảnh đại diện giáo viên. Dùng result_image() sau
    khi exec() == Accepted để lấy ảnh PIL đã crop."""

    def __init__(self, photo_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chỉnh sửa ảnh thẻ (tỉ lệ 3x4)")
        self.setModal(True)

        source = Image.open(photo_path).convert("RGB")
        source.thumbnail((_WORKING_MAX_DIM, _WORKING_MAX_DIM), Image.LANCZOS)
        self._source = source
        self._result = None

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(
            "Kéo ảnh để di chuyển, cuộn chuột (hoặc kéo thanh trượt bên dưới) để\n"
            "phóng to/thu nhỏ. Phần nằm trong khung xanh là đúng phần ảnh sẽ in lên thẻ."
        ))

        canvas_row = QHBoxLayout()
        canvas_row.addStretch(1)
        self._canvas = _CropCanvas(_pil_to_qpixmap(self._source))
        self._canvas.zoomChanged.connect(self._on_canvas_zoom_changed)
        canvas_row.addWidget(self._canvas)
        canvas_row.addStretch(1)
        layout.addLayout(canvas_row)

        zoom_row = QHBoxLayout()
        zoom_row.addWidget(QLabel("Thu nhỏ"))
        self._slider = QSlider(Qt.Orientation.Horizontal)
        self._slider.setRange(0, 1000)
        self._slider.setValue(0)
        self._slider.valueChanged.connect(self._on_slider_changed)
        zoom_row.addWidget(self._slider, 1)
        zoom_row.addWidget(QLabel("Phóng to"))
        layout.addLayout(zoom_row)

        btn_row = QHBoxLayout()
        btn_reset = QPushButton("Đặt lại")
        btn_reset.clicked.connect(self._on_reset)
        btn_row.addWidget(btn_reset)
        btn_row.addStretch(1)
        btn_cancel = QPushButton("Hủy")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        btn_save = QPushButton("Lưu ảnh thẻ")
        btn_save.setObjectName("primary")
        btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(btn_save)
        layout.addLayout(btn_row)

    def _on_slider_changed(self, value):
        zoom = self._canvas._min_zoom + (value / 1000) * (
            self._canvas._max_zoom - self._canvas._min_zoom
        )
        self._canvas.set_zoom(zoom)

    def _on_canvas_zoom_changed(self, fraction):
        self._slider.blockSignals(True)
        self._slider.setValue(round(fraction * 1000))
        self._slider.blockSignals(False)

    def _on_reset(self):
        self._canvas.reset()
        self._on_canvas_zoom_changed(self._canvas.zoom_fraction())

    def _on_save(self):
        x0, y0, x1, y1 = self._canvas.crop_box_on_source()
        cropped = self._source.crop((round(x0), round(y0), round(x1), round(y1)))
        self._result = cropped.resize((OUTPUT_W, OUTPUT_H), Image.LANCZOS)
        self.accept()

    def result_image(self):
        """Ảnh PIL (RGB, OUTPUT_W x OUTPUT_H, đúng tỉ lệ 3x4) sau khi bấm Lưu
        - None nếu dialog bị hủy."""
        return self._result

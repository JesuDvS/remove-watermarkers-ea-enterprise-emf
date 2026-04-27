import sys
import os
import shutil
import base64
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                             QVBoxLayout, QListWidget, QLabel, QPushButton,
                             QFileDialog, QSplitter, QMessageBox)
from PyQt6.QtCore import (Qt, QPoint, QRect, QRectF, QMimeData, 
                          QByteArray, QBuffer, QIODevice)
from PyQt6.QtGui import QPixmap, QPainter, QPen, QColor, QGuiApplication, QImage

try:
    import win32com.client
    from remove_watermark import remove_watermark_from_emf, recrop_emf
except ImportError:
    pass


class ClickableLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.start_point    = None
        self.end_point      = None
        self.selection_rect = None
        self.pixmap_rect    = QRectF()
        self.scale_factor   = 1.0
        self.offset         = QPoint(0, 0)
        self.setMouseTracking(True)
        self.image_loaded   = False
        self.dragging       = False

    def set_image(self, pixmap):
        self.start_point    = None
        self.end_point      = None
        self.selection_rect = None
        self.dragging       = False
        self.image_loaded   = True

        label_size  = self.size()
        pixmap_size = pixmap.size()
        sw = label_size.width()  / pixmap_size.width()
        sh = label_size.height() / pixmap_size.height()
        self.scale_factor = min(sw, sh)

        new_w = pixmap_size.width()  * self.scale_factor
        new_h = pixmap_size.height() * self.scale_factor
        self.offset = QPoint(
            int((label_size.width()  - new_w) / 2),
            int((label_size.height() - new_h) / 2)
        )
        self.pixmap_rect = QRectF(
            float(self.offset.x()), float(self.offset.y()),
            float(new_w), float(new_h)
        )
        self.setPixmap(pixmap.scaled(
            label_size,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        ))
        self.window().set_copy_button_enabled(False)

    # ── Drag-to-select ─────────────────────────────────────────────────

    def mousePressEvent(self, event):
        if not self.image_loaded:
            return
        if self.pixmap_rect.contains(event.position()):
            self.start_point    = event.position().toPoint()
            self.end_point      = self.start_point
            self.selection_rect = None
            self.dragging       = True
            self.window().set_copy_button_enabled(False)
            self.update()

    def mouseMoveEvent(self, event):
        if self.dragging and self.start_point:
            self.end_point = event.position().toPoint()
            self.update()

    def mouseReleaseEvent(self, event):
        if not self.dragging:
            return
        self.dragging = False
        if self.start_point and self.end_point:
            rect = QRect(self.start_point, self.end_point).normalized()
            if rect.width() > 5 and rect.height() > 5:
                self.selection_rect = rect
                self.window().set_copy_button_enabled(True)
            else:
                self.start_point    = None
                self.end_point      = None
                self.selection_rect = None
                self.window().set_copy_button_enabled(False)
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self.dragging and self.start_point and self.end_point:
            self._draw_selection(QRect(self.start_point, self.end_point).normalized())
        elif self.selection_rect:
            self._draw_selection(self.selection_rect)

    def _draw_selection(self, rect):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(rect, QColor(255, 50, 50, 60))
        painter.setPen(QPen(QColor(255, 50, 50), 2, Qt.PenStyle.SolidLine))
        painter.drawRect(rect)
        painter.setBrush(QColor(255, 50, 50))
        for corner in [rect.topLeft(), rect.topRight(),
                        rect.bottomLeft(), rect.bottomRight()]:
            painter.drawEllipse(corner, 5, 5)

    def get_proportional_rect(self):
        """Fracción [0,1] de la imagen real."""
        if not self.pixmap() or not self.selection_rect:
            return 0, 0, 0, 0

        def clamp(v): return max(0.0, min(1.0, v))

        disp_w = self.pixmap().width()
        disp_h = self.pixmap().height()

        r       = self.selection_rect
        corners = [r.topLeft(), r.topRight(), r.bottomLeft(), r.bottomRight()]
        xs, ys  = [], []
        for p in corners:
            xs.append(clamp((p.x() - self.offset.x()) / disp_w))
            ys.append(clamp((p.y() - self.offset.y()) / disp_h))
        return min(xs), min(ys), max(xs), max(ys)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EMF Processor & Extractor")
        self.resize(1100, 700)
        self.setAcceptDrops(True)
        self.repo_base    = "repositorio"
        os.makedirs(self.repo_base, exist_ok=True)
        self.current_emf  = None
        self.current_png  = None
        self.setup_ui()
        self.load_processed_files()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        splitter    = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self.display_selected_image)
        splitter.addWidget(self.list_widget)

        right_panel  = QWidget()
        right_layout = QVBoxLayout(right_panel)

        self.btn_load = QPushButton("📂 Cargar archivo EMF desde el explorador")
        self.btn_load.setFixedHeight(40)
        self.btn_load.clicked.connect(self.load_from_explorer)
        right_layout.addWidget(self.btn_load)

        self.lbl_viewer = ClickableLabel(self)
        self.lbl_viewer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_viewer.setStyleSheet(
            "background-color: #2b2b2b; color: #aaa; "
            "border: 2px dashed #555; font-size: 16px;"
        )
        self.lbl_viewer.setText(
            "Arrastra un archivo .emf aquí\no usa el botón superior.\n\n"
            "Haz clic y arrastra para seleccionar el área."
        )
        right_layout.addWidget(self.lbl_viewer)

        self.btn_copy = QPushButton("✂️ Copiar área seleccionada como PNG")
        self.btn_copy.setFixedHeight(50)
        self.btn_copy.setEnabled(False)
        self._style_copy_btn(enabled=False)
        self.btn_copy.clicked.connect(self.copy_selection_as_png)
        right_layout.addWidget(self.btn_copy)

        splitter.addWidget(right_panel)
        splitter.setSizes([250, 850])

    def _style_copy_btn(self, enabled: bool):
        if enabled:
            self.btn_copy.setStyleSheet(
                "background-color: #2ecc71; color: white; "
                "font-weight: bold; font-size: 14px;"
            )
        else:
            self.btn_copy.setStyleSheet(
                "background-color: #555; color: #999; "
                "font-weight: bold; font-size: 14px;"
            )

    def set_copy_button_enabled(self, enabled: bool):
        self.btn_copy.setEnabled(enabled)
        self._style_copy_btn(enabled)

    # ── Lista ──────────────────────────────────────────────────────────

    def load_processed_files(self):
        self.list_widget.clear()
        for root, dirs, files in os.walk(self.repo_base):
            for file in files:
                if file.endswith("_limpio_recortado.emf"):
                    self.list_widget.addItem(os.path.join(root, file))

    # ── Drag & drop / carga ────────────────────────────────────────────

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith('.emf'):
                self.process_file(path)

    def load_from_explorer(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar EMF", "", "EMF (*.emf)")
        for f in files:
            self.process_file(f)

    def process_file(self, file_path):
        name   = os.path.splitext(os.path.basename(file_path))[0]
        folder = os.path.join(self.repo_base, name)
        os.makedirs(folder, exist_ok=True)
        target = os.path.join(folder, os.path.basename(file_path))
        if not os.path.exists(target):
            try:
                shutil.copy(file_path, target)
            except shutil.SameFileError:
                pass
        try:
            self.lbl_viewer.setText(f"⏳ Procesando {name}...")
            QApplication.processEvents()
            clean = remove_watermark_from_emf(target)
            final = recrop_emf(clean)
            if final:
                self.load_processed_files()
                items = self.list_widget.findItems(final, Qt.MatchFlag.MatchExactly)
                if items:
                    self.list_widget.setCurrentItem(items[0])
                    self.display_selected_image(items[0])
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            self.lbl_viewer.setText("Arrastra un archivo .emf aquí")

    def display_selected_image(self, item):
        self.current_emf = item.text()
        png_path = os.path.splitext(self.current_emf)[0] + ".png"
        if os.path.exists(png_path):
            self.current_png = png_path
            self.lbl_viewer.setStyleSheet(
                "background-color: #2b2b2b; border: 1px solid #555;"
            )
            self.lbl_viewer.set_image(QPixmap(png_path))
        else:
            self.current_png = None
            self.lbl_viewer.setText(
                "No se encontró la vista previa PNG.\n"
                "Asegúrate de que el script original exporta el PNG."
            )

    # ── Copiar recorte PNG al portapapeles ─────────────────────────────

    def copy_selection_as_png(self):
        if not self.current_png or not os.path.exists(self.current_png):
            QMessageBox.warning(self, "Aviso", "No hay PNG cargado.")
            return

        left_pct, top_pct, right_pct, bottom_pct = self.lbl_viewer.get_proportional_rect()

        full_pixmap = QPixmap(self.current_png)
        if full_pixmap.isNull():
            QMessageBox.critical(self, "Error", "No se pudo leer el PNG.")
            return

        img_w = full_pixmap.width()
        img_h = full_pixmap.height()

        crop_x = int(left_pct  * img_w)
        crop_y = int(top_pct   * img_h)
        crop_w = int((right_pct  - left_pct) * img_w)
        crop_h = int((bottom_pct - top_pct)  * img_h)

        crop_w = max(1, crop_w)
        crop_h = max(1, crop_h)

        cropped = full_pixmap.copy(QRect(crop_x, crop_y, crop_w, crop_h))

        # --- SOLUCIÓN AVANZADA PARA LA TRANSPARENCIA DEL PORTAPAPELES ---
        # Aseguramos el formato con canal Alpha (ARGB32)
        cropped_image = cropped.toImage().convertToFormat(QImage.Format.Format_ARGB32)
        
        mime_data = QMimeData()
        
        # 1. Dato estándar de imagen
        mime_data.setImageData(cropped_image)
        
        # 2. Generar el binario en formato PNG puro
        byte_array = QByteArray()
        buffer = QBuffer(byte_array)
        buffer.open(QIODevice.OpenModeFlag.WriteOnly)
        cropped_image.save(buffer, "PNG")
        
        # 3. Forzar el MIME de PNG
        mime_data.setData("image/png", byte_array)
        
        # 4. EL TRUCO: Etiqueta HTML con base64 para evitar el fondo negro en Word/Excel/Navegadores
        b64_data = base64.b64encode(byte_array.data()).decode('utf-8')
        html_tag = f'<img src="data:image/png;base64,{b64_data}">'
        mime_data.setHtml(html_tag)
        
        QGuiApplication.clipboard().setMimeData(mime_data)
        # ----------------------------------------------------------------

        QMessageBox.information(
            self, "Éxito",
            f"✅ Área copiada al portapapeles ({crop_w}×{crop_h} px).\n"
            "Pega con Ctrl+V en cualquier aplicación."
        )

        self.lbl_viewer.selection_rect = None
        self.lbl_viewer.start_point    = None
        self.lbl_viewer.end_point      = None
        self.lbl_viewer.update()
        self.set_copy_button_enabled(False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    palette = app.palette()
    palette.setColor(palette.ColorRole.Window,          QColor(53, 53, 53))
    palette.setColor(palette.ColorRole.WindowText,      Qt.GlobalColor.white)
    palette.setColor(palette.ColorRole.Base,            QColor(25, 25, 25))
    palette.setColor(palette.ColorRole.AlternateBase,   QColor(53, 53, 53))
    palette.setColor(palette.ColorRole.ToolTipBase,     Qt.GlobalColor.white)
    palette.setColor(palette.ColorRole.ToolTipText,     Qt.GlobalColor.white)
    palette.setColor(palette.ColorRole.Text,            Qt.GlobalColor.white)
    palette.setColor(palette.ColorRole.Button,          QColor(53, 53, 53))
    palette.setColor(palette.ColorRole.ButtonText,      Qt.GlobalColor.white)
    palette.setColor(palette.ColorRole.BrightText,      Qt.GlobalColor.red)
    palette.setColor(palette.ColorRole.Link,            QColor(42, 130, 218))
    palette.setColor(palette.ColorRole.Highlight,       QColor(42, 130, 218))
    palette.setColor(palette.ColorRole.HighlightedText, Qt.GlobalColor.black)
    app.setPalette(palette)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
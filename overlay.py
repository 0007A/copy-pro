import sys
import tempfile
import os
from PyQt6.QtWidgets import QWidget, QApplication
from PyQt6.QtCore import Qt, QRect, pyqtSignal, QPoint
from PyQt6.QtGui import QPainter, QPen, QColor, QScreen
from PIL import ImageGrab

class Overlay(QWidget):
    # Signal emitted when an image is captured. Sends the temp image path.
    image_captured = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setCursor(Qt.CursorShape.CrossCursor)
        
        # Get all screens geometry to cover multiple monitors
        rect = QRect()
        for screen in QApplication.screens():
            rect = rect.united(screen.geometry())
            
        self.setGeometry(rect)
        
        self.start_point = QPoint()
        self.end_point = QPoint()
        self.is_drawing = False

    def paintEvent(self, event):
        painter = QPainter(self)
        # Semi-transparent dark background for the rest of the screen
        painter.fillRect(self.rect(), QColor(0, 0, 0, 100)) 
        
        if self.is_drawing:
            rect = QRect(self.start_point, self.end_point).normalized()
            
            # Draw the highlighter effect (Semi-transparent Yellow)
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
            
            # Highlighter color: Vibrant Yellow with 40/255 transparency (much lighter)
            highlighter_color = QColor(255, 255, 0, 40)
            painter.fillRect(rect, highlighter_color)
            
            # Draw a subtle bright yellow border
            pen = QPen(QColor(255, 255, 0, 200), 2)
            painter.setPen(pen)
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.start_point = event.pos()
            self.end_point = self.start_point
            self.is_drawing = True
            self.update()

    def mouseMoveEvent(self, event):
        if self.is_drawing:
            self.end_point = event.pos()
            self.update()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        super().keyPressEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.is_drawing = False
            self.end_point = event.pos()
            rect = QRect(self.start_point, self.end_point).normalized()
            
            # Hide overlay window before grabbing clean screen
            self.hide()
            QApplication.processEvents()
            
            # Capture the screen area
            if rect.width() > 5 and rect.height() > 5:
                try:
                    screen = QApplication.primaryScreen()
                    if screen:
                        pixmap = screen.grabWindow(0, rect.x(), rect.y(), rect.width(), rect.height())
                        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                        temp_path = temp_file.name
                        temp_file.close()
                        pixmap.save(temp_path, "PNG")
                        print(f"Captured image saved to: {temp_path}")
                        self.image_captured.emit(temp_path)
                except Exception as e:
                    print(f"Error capturing image: {e}")
            
            self.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    overlay = Overlay()
    overlay.show()
    sys.exit(app.exec())

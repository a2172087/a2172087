from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog,
    QInputDialog, QHBoxLayout, QSizePolicy, QGridLayout, QMessageBox,
    QGraphicsScene, QGraphicsView, QGraphicsItem,
    QDesktopWidget, QSlider, QDialog, QLineEdit, QGraphicsItem, QGraphicsItemGroup, QFrame
)
from PyQt5.QtGui import QPixmap, QFont, QImage, QPainter, QColor, QPen, QIcon, QPainterPath, QBrush
from PyQt5.QtCore import Qt, QPoint, QTimer, QSize, QRectF, QPointF, QLineF, QRect
from pathlib import Path
from PyQt5 import QtGui
import os
import re
import shutil
import subprocess
import sys
import socket
import datetime
import math
import logging
import time

STYLE_SHEET = """
QWidget {
    background-color: #2D2D2D;
    color: #E0E0E0;
    font-family: "微軟正黑體";
    font-size: 9pt;
}

QSpinBox {
    background-color: #333333;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 8px 12px;
    color: #E0E0E0;
    min-height: 24px;
}

QSpinBox:focus {
    border: 1px solid #505050;
    background-color: #383838;
}

QSpinBox:hover {
    background-color: #383838;
    border: 1px solid #505050;
}

QSpinBox::up-button, QSpinBox::down-button {
    background: none;
    border: none;
    width: 20px;
}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background-color: #454545;
}

QSpinBox::up-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-bottom: 5px solid #E0E0E0;
}

QSpinBox::down-arrow {
    image: none;
    border-left: 5px solid transparent;
    border-right: 5px solid transparent;
    border-top: 5px solid #E0E0E0;
}

QLineEdit {
    background-color: #333333;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 8px 12px;
    color: #E0E0E0;
    min-height: 24px;
    selection-background-color: #454545;
}

QLineEdit:focus {
    border: 1px solid #505050;
    background-color: #383838;
}

QLineEdit::placeholder {
    color: #808080;
}

QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #454545, stop:1 #383838);
    border: 1px solid #454545;
    border-radius: 6px;
    padding: 8px 16px;
    color: #E0E0E0;
    min-width: 120px;
    min-height: 20px;
    margin: 2px;
}

QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #505050, stop:1 #404040);
    border: 1px solid #505050;
}

QPushButton:pressed {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                              stop:0 #383838, stop:1 #404040);
    padding-top: 9px;
}

QPushButton:disabled {
    background: #333333;
    border: 1px solid #383838;
    color: #666666;
}

QListWidget {
    background-color: #333333;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 5px;
}

QListWidget::item {
    padding: 8px;
    margin: 2px;
    border-radius: 4px;
}

QListWidget::item:selected {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #505050, stop:1 #404040);
    color: #FFFFFF;
}

QListWidget::item:hover {
    background-color: #383838;
}

QTextEdit {
    background-color: #333333;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 8px;
    selection-background-color: #454545;
    color: #E0E0E0;
}

QTextEdit:focus {
    border: 1px solid #505050;
    background-color: #383838;
}

QProgressBar {
    border: none;
    border-radius: 4px;
    text-align: center;
    padding: 1px;
    background-color: #333333;
    height: 12px;
}

QProgressBar::chunk {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                              stop:0 #505050, stop:1 #404040);
    border-radius: 3px;
}

QLabel {
    color: #E0E0E0;
    padding: 2px;
}

QMessageBox {
    background-color: #2D2D2D;
}

QMessageBox QPushButton {
    min-width: 80px;
    min-height: 25px;
}

QScrollBar:vertical {
    border: none;
    background: #333333;
    width: 8px;
    margin: 0px;
    border-radius: 4px;
}

QScrollBar::handle:vertical {
    background: #454545;
    min-height: 20px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #505050;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: none;
}

QMenu {
    background-color: #2D2D2D;
    border: 1px solid #404040;
    border-radius: 6px;
    padding: 5px;
}

QMenu::item {
    background-color: transparent;
    padding: 8px 32px;
    border-radius: 4px;
    margin: 2px 5px;
}

QMenu::item:selected {
    background-color: #454545;
}

QMenu::item:hover {
    background-color: #383838;
}

QMenu::separator {
    height: 1px;
    background: #404040;
    margin: 5px 0px;
}

PhotomoveApp {
    background-color: #2D2D2D;
    border: 1px solid #404040;
    border-radius: 8px;
}
"""

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(application_path, 'format.ico')

class ResizableRotatableItem(QGraphicsItemGroup):
    def __init__(self, x, y, width, height):
        super().__init__()
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemSendsGeometryChanges)
        self._brush = QBrush(Qt.NoBrush)  
        self._pen = QPen(Qt.red, 2)
        self.handle_size = 5
        self.handles = []
        self.rotate_handle = None
        self.rotate_handle_distance = 20
        self.rect = QRectF(0, 0, width, height)
        self.rotation_angle = 0
        self.active_handle = None
        self.start_pos = None
        self.original_rect = None
        self.original_angle = 0
        self.update_handles()
        self.setFiltersChildEvents(True)

    def boundingRect(self):
        return self.rect.adjusted(-self.handle_size, -self.handle_size - self.rotate_handle_distance,
                                  self.handle_size, self.handle_size)

    def shape(self):
        path = QPainterPath()
        path.addRect(self.rect)
        for handle in self.handles + [self.rotate_handle]:
            path.addEllipse(self.handle_rect(handle))
        return path

    def paint(self, painter, option, widget):
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setBrush(self._brush)
        painter.setPen(self._pen)
        
        painter.save()
        painter.translate(self.rect.center())
        painter.rotate(self.rotation_angle)
        painter.translate(-self.rect.center())
        
        self.draw_shape(painter)
        
        if self.isSelected():
            painter.setBrush(Qt.green)
            for handle in self.handles:
                painter.drawEllipse(self.handle_rect(handle))
            
            painter.setBrush(Qt.blue)
            painter.drawEllipse(self.handle_rect(self.rotate_handle))
            painter.drawLine(self.rect.center(), self.rotate_handle)
        
        painter.restore()

    def draw_shape(self, painter):
        painter.drawRect(self.rect)

    def update_handles(self):
        self.handles = [
            self.rect.topLeft(), self.rect.topRight(),
            self.rect.bottomLeft(), self.rect.bottomRight(),
            QPointF(self.rect.left(), self.rect.center().y()),
            QPointF(self.rect.right(), self.rect.center().y()),
            QPointF(self.rect.center().x(), self.rect.top()),
            QPointF(self.rect.center().x(), self.rect.bottom())
        ]
        self.rotate_handle = QPointF(self.rect.center().x(), self.rect.top() - self.rotate_handle_distance)

    def handle_rect(self, pt):
        return QRectF(pt.x() - self.handle_size / 2, pt.y() - self.handle_size / 2,
                      self.handle_size, self.handle_size)

    def handle_contains(self, handle, pos):
        handle_rect = self.handle_rect(handle)
        return handle_rect.contains(pos) or QLineF(handle_rect.center(), pos).length() < self.handle_size / 2

    def mousePressEvent(self, event):
        self.start_pos = event.pos()
        self.original_rect = self.rect
        self.original_angle = self.rotation_angle
        
        if self.handle_contains(self.rotate_handle, event.pos()):
            self.active_handle = 'rotate'
            event.accept()
            return
        
        for i, handle in enumerate(self.handles):
            if self.handle_contains(handle, event.pos()):
                self.active_handle = i
                event.accept()
                return
        self.active_handle = None
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.active_handle == 'rotate':
            self.interactive_rotate(event.pos())
        elif self.active_handle is not None:
            self.interactive_resize(event.pos())
        else:
            super().mouseMoveEvent(event)
        self.prepareGeometryChange()
        self.scene().update()
        event.accept()

    def mouseReleaseEvent(self, event):
        self.active_handle = None
        super().mouseReleaseEvent(event)
        self.prepareGeometryChange()
        self.scene().update()
        event.accept()

    def interactive_resize(self, mouse_pos):
        diff = mouse_pos - self.start_pos
        new_rect = QRectF(self.original_rect)

        if self.active_handle in [0, 1, 2, 3]:
            if self.active_handle in [0, 2]:
                new_rect.setLeft(new_rect.left() + diff.x())
            else:
                new_rect.setRight(new_rect.right() + diff.x())
            
            if self.active_handle in [0, 1]:
                new_rect.setTop(new_rect.top() + diff.y())
            else:
                new_rect.setBottom(new_rect.bottom() + diff.y())

        elif self.active_handle == 4:
            new_rect.setLeft(new_rect.left() + diff.x())
        elif self.active_handle == 5:
            new_rect.setRight(new_rect.right() + diff.x())
        elif self.active_handle == 6:
            new_rect.setTop(new_rect.top() + diff.y())
        elif self.active_handle == 7:
            new_rect.setBottom(new_rect.bottom() + diff.y())

        if new_rect.width() >= 10 and new_rect.height() >= 10:
            self.rect = new_rect
            self.update_handles()

    def interactive_rotate(self, mouse_pos):
        center = self.rect.center()
        start_vector = self.start_pos - center
        current_vector = mouse_pos - center
        angle = math.degrees(math.atan2(current_vector.y(), current_vector.x()) - 
                             math.atan2(start_vector.y(), start_vector.x()))
        self.rotation_angle = (self.original_angle + angle) % 360

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemSelectedHasChanged:
            self.setZValue(10 if self.isSelected() else 0)
        return super().itemChange(change, value)

    def get_actual_shape(self):
        path = QPainterPath()
        path.addRect(self.rect)
        return path

    def get_actual_area(self):
        return self.rect.width() * self.rect.height()

class ResizableRotatableRectItem(ResizableRotatableItem):
    def draw_shape(self, painter):
        painter.drawRect(self.rect)

class ResizableRotatableEllipseItem(ResizableRotatableItem):
    def __init__(self, x, y, width, height):
        super().__init__(x, y, width, height)
        self._brush = QBrush(Qt.NoBrush)  
        self._pen = QPen(Qt.blue, 2)

    def draw_shape(self, painter):
        painter.drawEllipse(self.rect)

    def get_actual_shape(self):
        path = QPainterPath()
        path.addEllipse(self.rect)
        return path

    def get_actual_area(self):
        return math.pi * (self.rect.width() / 2) * (self.rect.height() / 2)

class MeasurementWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.shapes = []
        self.initUI()
        self.setWindowTitle('量測Probe mark Area(%)')
        self.resize(int(main_window.width() * 0.6), int(main_window.height() * 1.1))
        self.scale_factor = 1.0
        self.current_shape = None
        self.resize_handle = None
        self.rotate_handle = None

    def initUI(self):
        self.layout = QVBoxLayout(self)
        
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        
        self.button_layout = QHBoxLayout()
        
        self.draw_rect_button = QPushButton('繪製矩形', self)
        self.draw_rect_button.clicked.connect(self.draw_rectangle)
        self.button_layout.addWidget(self.draw_rect_button)
        
        self.draw_circle_button = QPushButton('繪製圓形', self)
        self.draw_circle_button.clicked.connect(self.draw_circle)
        self.button_layout.addWidget(self.draw_circle_button)
        
        self.calculate_area_button = QPushButton('計算面積佔比', self)
        self.calculate_area_button.clicked.connect(self.calculate_area_ratio)
        self.button_layout.addWidget(self.calculate_area_button)
        
        self.close_button = QPushButton('關閉', self)
        self.close_button.clicked.connect(self.close)
        self.button_layout.addWidget(self.close_button)
        
        self.layout.addLayout(self.button_layout)

        self.zoom_slider = QSlider(Qt.Horizontal)
        self.zoom_slider.setMinimum(10)
        self.zoom_slider.setMaximum(200)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setTickPosition(QSlider.TicksBelow)
        self.zoom_slider.setTickInterval(10)
        self.zoom_slider.valueChanged.connect(self.zoom_image)
        self.layout.addWidget(self.zoom_slider)

        self.setup_scene()

    def setup_scene(self):
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene, self)
        self.view.setRenderHint(QPainter.Antialiasing)
        self.view.setRenderHint(QPainter.SmoothPixmapTransform)
        self.view.setViewportUpdateMode(QGraphicsView.FullViewportUpdate)
        
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)
        
        self.layout.addWidget(self.view)

    def update_image(self, pixmap):
        self.original_pixmap = pixmap
        self.scene.clear()
        self.pixmap_item = self.scene.addPixmap(pixmap)
        self.view.setSceneRect(QRectF(pixmap.rect()))
        self.view.fitInView(self.scene.sceneRect(), Qt.KeepAspectRatio)
        self.scale_factor = 1.0
        self.zoom_slider.setValue(100)

    def zoom_image(self, value):
        try:
            self.scale_factor = value / 50.0
            self.view.resetTransform()
            self.view.scale(self.scale_factor, self.scale_factor)
            
            for shape in self.shapes:
                self.view.ensureVisible(shape)
        except Exception as e:
            print(f"縮放圖片時發生錯誤: {e}")

    def draw_rectangle(self):
        rect_item = ResizableRotatableRectItem(0, 0, 100, 100)
        self.add_shape(rect_item)

    def draw_circle(self):
        ellipse_item = ResizableRotatableEllipseItem(0, 0, 100, 100)
        self.add_shape(ellipse_item)

    def add_shape(self, shape):
        shape.setPos(self.view.mapToScene(self.view.viewport().rect().center()) - shape.boundingRect().center())
        self.scene.addItem(shape)
        self.shapes.append(shape)
        self.view.ensureVisible(shape)

    def calculate_area_ratio(self):
        if len(self.shapes) < 2:
            QMessageBox.information(self, '提示', '請至少繪製兩個圖形以進行面積比較')
            return
        areas = [shape.get_actual_area() for shape in self.shapes]
        min_area = min(areas)
        max_area = max(areas)
        ratio = min_area / max_area * 100
        QMessageBox.information(self, '面積佔比', f'最小圖形面積與最大圖形面積的比值: {ratio:.2f}%')

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            for item in self.scene.selectedItems():
                if item in self.shapes:
                    self.shapes.remove(item)
                    self.scene.removeItem(item)

    def reset_shapes_transform(self):
        for shape in self.shapes:
            shape.setTransform(QtGui.QTransform())

    def clear_shapes(self):
        for shape in self.shapes:
            self.scene.removeItem(shape)
        self.shapes.clear()

class UmSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.initUI()
        self.setFixedSize(350, 500)  
        
    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(15)  
        layout.setContentsMargins(20, 20, 20, 20) 
        
        title_container = QFrame()
        title_container.setStyleSheet("""
            QFrame {
                background-color: #1E1E1E;
                border-radius: 8px;
                margin-bottom: 10px;
            }
        """)
        title_layout = QVBoxLayout(title_container)
        title_layout.setContentsMargins(0, 15, 0, 15)  
        
        title = QLabel("設置測量參數")
        title.setStyleSheet("""
            QLabel {
                color: #FFFFFF;
                font-size: 20px;
                font-weight: bold;
            }
        """)
        title.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(title)
        layout.addWidget(title_container)
        
        pix_spec_container = QFrame()
        pix_spec_container.setStyleSheet("""
            QFrame {
                background-color: #2A2A2A;
                border-radius: 8px;
                padding: 5px;
            }
        """)
        pix_spec_layout = QVBoxLayout(pix_spec_container)
        pix_spec_layout.setSpacing(8) 
        pix_spec_layout.setContentsMargins(15, 15, 15, 15) 
        
        pix_spec_title = QLabel("Pix Spec 設置")
        pix_spec_title.setStyleSheet("color: #FFFFFF; font-size: 16px; font-weight: bold;")
        pix_spec_layout.addWidget(pix_spec_title)
        
        self.pix_spec_input = QLineEdit()
        self.pix_spec_input.setPlaceholderText("請輸入 pix spec 值")
        self.pix_spec_input.setStyleSheet("""
            QLineEdit {
                background-color: #363636;
                border: 1px solid #404040;
                border-radius: 6px;
                padding: 8px;
                color: #FFFFFF;
                min-height: 20px;
            }
        """)
        pix_spec_layout.addWidget(self.pix_spec_input)
        layout.addWidget(pix_spec_container)

        um_container = QFrame()
        um_container.setStyleSheet("""
            QFrame {
                background-color: #2A2A2A;
                border-radius: 8px;
                padding: 5px;
            }
        """)
        um_layout = QVBoxLayout(um_container)
        um_layout.setSpacing(8)
        um_layout.setContentsMargins(15, 15, 15, 15)
        
        um_title = QLabel("μm 設置")
        um_title.setStyleSheet("color: #FFFFFF; font-size: 16px; font-weight: bold;")
        um_layout.addWidget(um_title)
        
        self.um_input = QLineEdit()
        self.um_input.setPlaceholderText("請輸入 μm 值")
        self.um_input.setStyleSheet("""
            QLineEdit {
                background-color: #363636;
                border: 1px solid #404040;
                border-radius: 6px;
                padding: 8px;
                color: #FFFFFF;
                min-height: 20px;
            }
        """)
        um_layout.addWidget(self.um_input)
        layout.addWidget(um_container)

        link_container = QFrame()
        link_container.setStyleSheet("""
            QFrame {
                background-color: #2A2A2A;
                border-radius: 8px;
            }
        """)
        link_layout = QVBoxLayout(link_container)
        link_layout.setContentsMargins(15, 10, 15, 10)
        
        self.param_link = QLabel()
        self.param_link.setAlignment(Qt.AlignCenter)
        self.param_link.setText('<a href="file:///C:/BITools/Classify/AVI Resolution 對照表.xlsx" style="color: #0078D4; text-decoration: none;">參數對照表 →</a>')
        self.param_link.setOpenExternalLinks(True)
        link_layout.addWidget(self.param_link)
        layout.addWidget(link_container)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        cancel_button = QPushButton("取消")
        cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #404040;
                border: none;
                border-radius: 6px;
                color: #FFFFFF;
                padding: 8px 20px;
                min-width: 80px;
            }
        """)
        cancel_button.clicked.connect(self.reject)
        
        confirm_button = QPushButton("確認")
        confirm_button.setStyleSheet("""
            QPushButton {
                background-color: #0078D4;
                border: none;
                border-radius: 6px;
                color: #FFFFFF;
                padding: 8px 20px;
                min-width: 80px;
            }
        """)
        confirm_button.clicked.connect(self.accept)
        
        button_layout.addStretch()
        button_layout.addWidget(cancel_button)
        button_layout.addWidget(confirm_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)
        self.setWindowTitle("測量參數設置")

    def get_settings(self):
        return self.pix_spec_input.text(), self.um_input.text()

class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.Window | Qt.WindowType.WindowMinimizeButtonHint | Qt.WindowType.WindowCloseButtonHint)
        self.initUI()
        self.setWindowTitle('Classify')
        app.setWindowIcon(QIcon(icon_path))
        self.log_file = None
        self.user_id = None
        self.classified_photos = []
        self.show_red_circle = False 
        self.save_log()
        self.check_version()
        self.is_resizing = False
        self.start_pos = None
        self.start_radius = None
        self.is_moving_circle = False
        self.is_resizing_circle = False
        self.last_classification_time = 0
        self.classification_cooldown = 0.1  

    def initUI(self):
        main_layout = QHBoxLayout(self)

        left_panel = QVBoxLayout()
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        left_panel.addWidget(self.image_label)

        right_panel = QVBoxLayout()

        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        if screen_width >= 3840 and screen_height >= 2160:
            button_width_factor = 2
            button_height_factor = 0.2
        else:
            button_width_factor = 1.2
            button_height_factor = 0.11

        self.save_path_button = QPushButton('另存儲存路徑', self)
        self.save_path_button.clicked.connect(self.select_save_path)
        self.save_path_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.save_path_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.save_path_button.hide()
        right_panel.addWidget(self.save_path_button)

        self.open_original_button = QPushButton('開啟圖片原始檔', self)
        self.open_original_button.clicked.connect(self.open_original_image)
        self.open_original_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.open_original_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.open_original_button.hide()
        right_panel.addWidget(self.open_original_button)

        self.select_folder_button = QPushButton('請選擇要Review的資料夾路徑', self)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.select_folder_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.select_folder_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        right_panel.addWidget(self.select_folder_button)

        self.um_input_button = QPushButton('量測Defect size (μm) -- 限定500x500尺寸使用' , self)
        self.um_input_button.clicked.connect(self.change_um_size)
        self.um_input_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.um_input_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.um_input_button.setEnabled(False)
        self.um_input_button.hide()
        right_panel.addWidget(self.um_input_button)

        self.measure_area_button = QPushButton('量測Probe mark Area (%)', self)
        self.measure_area_button.clicked.connect(self.open_measurement_ui)
        self.measure_area_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.measure_area_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.measure_area_button.setEnabled(False)
        self.measure_area_button.hide()
        right_panel.addWidget(self.measure_area_button)
        
        self.go_back_button = QPushButton('返回上一張 (快捷鍵 : 空白鍵)', self)
        self.go_back_button.clicked.connect(self.go_back_photo)
        self.go_back_button.setShortcut(' ')
        self.go_back_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.go_back_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.go_back_button.hide()
        right_panel.addWidget(self.go_back_button)

        self.remaining_photos_label = QLabel('剩餘照片數量: 0', self)
        self.remaining_photos_label.setAlignment(Qt.AlignCenter)
        self.remaining_photos_label.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.remaining_photos_label.hide()
        right_panel.addWidget(self.remaining_photos_label)

        self.button_grid_layout = QGridLayout()
        right_panel.addLayout(self.button_grid_layout)

        self.source_folder = None
        self.target_folders = {}
        self.save_folder = None

        self.circle_center = QPoint(250, 250)
        self.circle_radius = 50
        self.circle_color = QColor(255, 0, 0, 150)

        screen_resolution = app.desktop().screenGeometry()
        if screen_resolution.width() == 3840 and screen_resolution.height() == 2160:
            self.setGeometry(350, 600, 400, 250)
        else:
            self.setGeometry(100, 200, 350, 150)

        self.resize_timer = QTimer(self)

        main_layout.addLayout(left_panel)
        main_layout.addLayout(right_panel)

        self.original_um_value = 0 

    def setFocusOnMainWindow(self):
        self.setFocus()
        self.activateWindow()

    def keyPressEvent(self, event):
        shortcuts = {
            Qt.Key_I: 'ugly die',
            Qt.Key_E: 'foreign material',
            Qt.Key_U: 'Particle',
            Qt.Key_R: 'Probe mark shift',
            Qt.Key_Y: 'Over kill',
            Qt.Key_Q: 'Process defect',
            Qt.Key_A: 'Other',
            Qt.Key_P: 'Al particle out of pad',
            Qt.Key_J: 'Al particle',
            Qt.Key_W: 'Wafer Scratch',
            Qt.Key_F: 'PM area out spec',
            Qt.Key_L: 'Bump PM shift',
            Qt.Key_Z: 'Bump scratch',
            Qt.Key_X: 'Bump foreign Material',
            Qt.Key_D: 'Bump PM diameter out of spec',
            Qt.Key_T: 'PM No. Out Spec',
            Qt.Key_O: 'Pad discoloration',
            Qt.Key_S: 'Irregular bump',
            Qt.Key_G: 'Probing Void',
            Qt.Key_H: 'Missing Probe Mark',
            Qt.Key_K: 'Surface(Incoming defect)',
            Qt.Key_C: 'Missing bump',
            Qt.Key_V: 'Bump residue',
            Qt.Key_B: 'Bump house defect',
            Qt.Key_N: 'Large defect',
            Qt.Key_1: 'Large bump',
            Qt.Key_2: 'small bump',
            Qt.Key_3: '380 special PM shift',
        }

        key = event.key()
        if key in shortcuts:
            self.classify_photo_with_cooldown(shortcuts[key])
        else:
            super().keyPressEvent(event)

    def select_save_path(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec_():
            save_path = dialog.selectedFiles()[0]
            self.save_folder = Path(save_path)

    def open_original_image(self):
        if self.photo_files:
            photo_path = self.photo_files[0]
            if os.path.exists(photo_path):
                os.system(f'start "" "{photo_path}"')

    def select_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)  
        dialog.setOption(QFileDialog.ShowDirsOnly, False) 
        dialog.setViewMode(QFileDialog.Detail) 
        dialog.resize(1000, 600)
        dialog.setStyleSheet("""
            QFileDialog, QLabel, QPushButton, QComboBox, QLineEdit, QListView, QTreeView {
                background-color: white;
                color: black;
            }
            QFileDialog QListView::item:selected, QFileDialog QTreeView::item:selected {
                background-color: #0078D7;
                color: white;
            }
        """)

        if dialog.exec_():
            folder_path = dialog.selectedFiles()[0]
            self.source_folder = Path(folder_path)
            self.original_source_folder = self.source_folder

            folder_names = {
                'ugly die': '151_Ugly_Die(2D)',
                'foreign material': '102_Foreign_material(09)',
                'Particle': '000_Particle(16)',
                'Probe mark shift': '200_Probe_Mark_Shift(10)',
                'Bump foreign Material': '502_Bump foreign Material(25)',
                'Over kill': '000_Over_kill(15)',
                'Process defect': '100_Process_Defect(07)',
                'Al particle out of pad': '205_Al_particle_out_of_pad(0F)',
                'Al particle': '999_Al_particle(18)',
                'Wafer Scratch': '101_Wafer Scratch(08)',
                'PM area out spec': '202_PM area out spec.(12)',
                'Bump PM shift': '500_Bump PM shift(23)',
                'Bump scratch': '501_Bump scratch(24)',
                'Bump PM diameter out of spec': '507_Bump PM diameter out of spec(0E)',
                'Other': '186_Other(BA)',
                'PM No. Out Spec': '201_PM No. Out Spec(11)',
                'Pad discoloration': '117_Pad discoloration(1B)',
                'Irregular bump': '505_Irregular bump(0A)',
                'Probing Void': '203_Probing Void(13)',
                'Missing Probe Mark': '204_Missing Probe Mark(14)',
                'Surface(Incoming defect)': '100_Surface(Incoming defect)(1C)',
                'Missing bump': '503_Missing bump(26)',
                'Bump residue': '504_Bump residue(27)',
                'Bump house defect': '506_Bump house defect(28)',
                'Large defect': '115_Large defect(31)',
                'Large bump': '510_Large bump(32)',
                'small bump': '521_small bump(3D)',
                '380 special PM shift': '522_380 special PM shift(3E)',
            }

            self.target_folders = {name: self.source_folder / folder for name, folder in folder_names.items()}
            self.photo_files = self.load_photos_in_folder(self.source_folder)

            if self.photo_files:
                self.um_input_button.setEnabled(True)
                self.um_input_button.show()
                self.open_original_button.show()
                self.remaining_photos_label.show()
                self.save_path_button.show()
                self.go_back_button.show()
                self.measure_area_button.setEnabled(True)
                self.measure_area_button.show()
                self.reset_button_grid_layout()
                
                if hasattr(self, 'measurement_window'):
                    self.measurement_window.reset_shapes_transform()
                    self.measurement_window.clear_shapes()
                
                self.show_red_circle = False  
                self.show_photo(self.photo_files[0])
                self.select_folder_button.hide()
                self.update_remaining_photos_label()

            parent_folder = self.source_folder.parent
            subfolders = [f for f in parent_folder.iterdir() if f.is_dir()]
            valid_subfolders = []
            for subfolder in subfolders:
                if any(str(i).zfill(2) in subfolder.name for i in range(1, 26)):
                    photo_files = self.load_photos_in_folder(subfolder)
                    if photo_files and subfolder != self.source_folder:
                        valid_subfolders.append(subfolder)

            if valid_subfolders:
                self.next_folder_list = valid_subfolders
            else:
                self.next_folder_list = []

            self.selected_folders = []

    def reset_button_grid_layout(self):
        for i in reversed(range(self.button_grid_layout.count())):
            widget_to_remove = self.button_grid_layout.itemAt(i).widget()
            self.button_grid_layout.removeWidget(widget_to_remove)
            widget_to_remove.setParent(None)

        row, col = 0, 0
        for class_name in self.target_folders.keys():
            button = QPushButton(class_name, self)
            self.set_button_shortcut(button, class_name)
            self.button_grid_layout.addWidget(button, row, col)

            col += 1
            if col > 3:
                col = 0
                row += 1

    def set_button_shortcut(self, button, class_name):
        shortcuts = {
            'ugly die': ('I', 'Ugly die (I)'),
            'foreign material': ('E', 'Foreign material (E)'),
            'Particle': ('U', 'Particle (U)'),
            'Probe mark shift': ('R', 'Probe mark shift (R)'),
            'Over kill': ('Y', 'Over kill (Y)'),
            'Process defect': ('Q', 'Process defect (Q)'),
            'Other': ('A', 'Other (A)'),
            'Al particle out of pad': ('P', 'Al particle out of pad (P)'),
            'Al particle': ('J', 'Al particle (J)'),
            'Wafer Scratch': ('W', 'Wafer Scratch (W)'),
            'PM area out spec': ('F', 'PM area out spec (F)'),
            'Bump PM shift': ('L', 'Bump PM shift (L)'),
            'Bump scratch': ('Z', 'Bump scratch (Z)'),
            'Bump foreign Material': ('X', 'Bump foreign Material (X)'),
            'Bump PM diameter out of spec': ('D', 'Bump PM diameter out spec (D)'),
            'PM No. Out Spec': ('T', 'PM No. Out Spec (T)'),
            'Pad discoloration': ('O', 'Pad discoloration (O)'),
            'Irregular bump': ('S', 'Irregular bump (S)'),
            'Probing Void': ('G', 'Probing Void (G)'),
            'Missing Probe Mark': ('H', 'Missing Probe Mark (H)'),
            'Surface(Incoming defect)': ('K', 'Surface(In... defect) (K)'),
            'Missing bump': ('C', 'Missing bump (C)'),
            'Bump residue': ('V', 'Bump residue (V)'),
            'Bump house defect': ('B', 'Bump house defect (B)'),
            'Large defect': ('N', 'Large defect (N)'),
            'Large bump': ('1', 'Large bump (1)'),
            'small bump': ('2', 'small bump (2)'),
            '380 special PM shift': ('3', '380 special PMS (3)'),
        }
        if class_name in shortcuts:
            shortcut, text = shortcuts[class_name]
            button.setShortcut(shortcut)
            button.setText(text)
            button.clicked.connect(lambda: self.classify_photo_with_cooldown(class_name))

    def classify_photo_with_cooldown(self, class_name):
        current_time = time.time()
        if current_time - self.last_classification_time >= self.classification_cooldown:
            self.classify_photo(class_name)
            self.last_classification_time = current_time

    def draw_scale_bar(self, painter, pixmap_width, pixmap_height, pix_spec, um_size):
        logging.debug(f"開始繪製比例尺: pixmap_width={pixmap_width}, pixmap_height={pixmap_height}, pix_spec={pix_spec}, um_size={um_size}")
        
        try:
            scale_values = [0, 25, 50, 75, 100, 125, 150, 175, 200]  
            scale_bar_height = 40 
            margin = 30  
            vertical_position = pixmap_height * 0.95  

            screen_resolution = QDesktopWidget().screenGeometry()
            screen_width = screen_resolution.width()
            screen_height = screen_resolution.height()
            
            if screen_width == 3840 and screen_height == 2160:
                scale_factor = 2.1
            else:
                scale_factor = 1.3

            scale_bar_width_um = 200  
            scale_bar_width = (scale_bar_width_um / float(pix_spec)) * scale_factor
            horizontal_position = (pixmap_width - scale_bar_width) / 2  

            logging.debug(f"設置比例尺參數: scale_bar_width={scale_bar_width}, scale_bar_height={scale_bar_height}, margin={margin}, vertical_position={vertical_position}, horizontal_position={horizontal_position}")
            
            font = painter.font()
            font.setPointSize(7)
            font.setBold(True)
            painter.setFont(font)
            
            painter.setBrush(QColor(0, 0, 0, 100))  
            painter.setPen(Qt.NoPen)
            background_rect = QRectF(horizontal_position - margin, 
                                    vertical_position - scale_bar_height, 
                                    scale_bar_width + margin * 2 + 30, scale_bar_height)  
            painter.drawRect(background_rect)
            logging.debug(f"繪製背景: rect={background_rect}")
            
            painter.setPen(QPen(Qt.white, 2))
            logging.debug("開始繪製刻度和標籤")
            for value in scale_values:
                x = int(horizontal_position + (value / 200) * scale_bar_width)
                y = int(vertical_position - 5)
                painter.drawLine(x, y, x, y - 10)  
                painter.drawText(x - 10, y + 25, str(value))  
                logging.debug(f"繪製刻度: x={x}, y={y}, label={value}")
            
            start_x = int(horizontal_position)
            end_x = int(horizontal_position + scale_bar_width)
            y = int(vertical_position - 5)
            painter.drawLine(start_x, y, end_x, y)
            logging.debug(f"繪製橫線: start_x={start_x}, end_x={end_x}, y={y}")
            
            unit_x = int(end_x + 20)  
            unit_y = int(y + 25)
            painter.drawText(unit_x, unit_y, 'μm')
            logging.debug(f"繪製單位: x={unit_x}, y={unit_y}")
            
            logging.debug("比例尺繪製完成")
        except Exception as e:
            logging.error(f"繪製比例尺時發生錯誤: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())

    def show_photo(self, photo_path):
        try:
            logging.debug(f"開始顯示圖片: {photo_path}")
            screen_resolution = QDesktopWidget().screenGeometry()
            screen_width = screen_resolution.width()
            screen_height = screen_resolution.height()
            logging.debug(f"螢幕解析度: {screen_width}x{screen_height}")

            image = QImage(photo_path)
            if image.isNull():
                logging.error(f"無法載入圖片: {photo_path}")
                return
            logging.debug(f"圖片成功載入: {photo_path}")

            is_large_image = image.width() >= 4090 and image.height() >= 2168
            logging.debug(f"圖片尺寸: {image.width()}x{image.height()}, is_large_image: {is_large_image}")
            
            if screen_width >= 3840 and screen_height >= 2160:
                display_size = QSize(1024, 768)
                scaling_factor = 2.0 if is_large_image else 1.4
            else:
                display_size = QSize(800, 600)
                scaling_factor = 1.3 if is_large_image else 1.1

            logging.debug(f"顯示大小: {display_size}, 縮放因子: {scaling_factor}")

            image = QImage(photo_path)
            if image.isNull():
                logging.error(f"無法載入圖片: {photo_path}")
                return
            logging.debug(f"圖片成功載入: {photo_path}")

            is_500x500 = image.width() == 500 and image.height() == 500
            self.um_input_button.setEnabled(is_500x500)
            if is_500x500:
                logging.debug("圖片尺寸為 500x500，啟用 '量測Defect size (μm)' 按鈕")
            else:
                logging.debug(f"圖片尺寸不是 500x500 ({image.width()}x{image.height()})，禁用 '量測Defect size (μm)' 按鈕")

            self.um_input_button.show()

            pixmap = QPixmap.fromImage(image)
            pixmap = pixmap.scaled(display_size * scaling_factor, Qt.KeepAspectRatio)
            logging.debug(f"圖片已縮放至: {pixmap.width()}x{pixmap.height()}")
            
            final_pixmap = pixmap.copy()
            painter = QPainter(final_pixmap)
            painter.setRenderHint(QPainter.Antialiasing)
            
            if is_500x500 and (self.show_red_circle or hasattr(self, 'pix_spec')):
                if self.show_red_circle:
                    logging.debug(f"繪製紅圈: 中心={self.circle_center}, 半徑={self.circle_radius}")
                    pen = QPen(self.circle_color)
                    pen.setWidth(2)
                    painter.setPen(pen)
                    
                    center = QPoint(int(self.circle_center.x()), int(self.circle_center.y()))
                    radius = int(self.circle_radius)
                    painter.drawEllipse(center, radius, radius)
                    
                    if hasattr(self, 'um_size'):
                        logging.debug(f"開始繪製 um 值: {self.um_size}")
                        font = painter.font()
                        font.setPointSize(10)
                        painter.setFont(font)
                        
                        painter.setBrush(QColor(0, 0, 0, 0))
                        painter.setPen(Qt.NoPen)
                        
                        text = f"{self.um_size:.2f} μm"
                        text_rect = painter.boundingRect(QRect(), Qt.AlignLeft, text)
                        
                        text_x = int(center.x() + radius + 5)
                        text_y = int(center.y() - radius - 5 - text_rect.height())
                        
                        background_rect = text_rect.adjusted(-5, -5, 5, 5)
                        background_rect.moveTopLeft(QPoint(text_x, text_y))
                        painter.drawRect(background_rect)
                        
                        painter.setPen(Qt.white)
                        painter.drawText(text_x, text_y + text_rect.height(), text)
                        logging.debug(f"um 值已繪製: {text} 在位置 ({text_x}, {text_y})")
                
                if hasattr(self, 'pix_spec') and hasattr(self, 'um_size'):
                    logging.debug(f"繪製比例尺: pix_spec={self.pix_spec}, um_size={self.um_size}")
                    try:
                        self.draw_scale_bar(painter, pixmap.width(), pixmap.height(), self.pix_spec, self.um_size)
                    except Exception as e:
                        logging.error(f"繪製比例尺時發生錯誤: {str(e)}", exc_info=True)

            file_name = os.path.basename(photo_path)
            font = painter.font()
            font.setPointSize(10)
            painter.setFont(font)
            painter.setPen(Qt.white)
            text_rect = painter.boundingRect(QRect(), Qt.AlignLeft, file_name)
            text_x = int((pixmap.width() - text_rect.width()) / 2)
            text_y = int(pixmap.height() * 0.05) 

            bg_rect = text_rect.adjusted(-5, -5, 5, 5)
            bg_rect.moveCenter(QPoint(pixmap.width() // 2, text_y))
            painter.fillRect(bg_rect, QColor(0, 0, 0, 100))
            
            painter.drawText(text_x, text_y, file_name)
            logging.debug(f"文件名已繪製: {file_name} 在位置 ({text_x}, {text_y})")
            
            painter.end()
            self.image_label.setPixmap(final_pixmap)
            logging.debug("最終圖片已設置到 image_label")

            self.update_measurement_ui_image()
            logging.debug("測量 UI 圖像已更新")
            logging.debug("圖片顯示完成")
        except Exception as e:
            logging.error(f"在 show_photo 方法中發生錯誤: {str(e)}", exc_info=True)
        finally:
            if 'painter' in locals():
                painter.end()
            if 'final_pixmap' in locals():
                del final_pixmap
            if 'pixmap' in locals():
                del pixmap
            if 'image' in locals():
                del image

    def change_um_size(self):
        logging.debug("開始更改 um 尺寸")
        dialog = UmSettingsDialog(self)
        
        if hasattr(self, 'pix_spec'):
            dialog.pix_spec_input.setText(str(self.pix_spec))
        
        if hasattr(self, 'um_size'):
            dialog.um_input.setText(str(self.um_size))
        
        if dialog.exec_():
            self.pix_spec, um_size = dialog.get_settings()
            self.um_size = float(um_size) if um_size else 0

            screen_resolution = QDesktopWidget().screenGeometry()
            screen_width = screen_resolution.width()
            screen_height = screen_resolution.height()
            
            if self.pix_spec and self.um_size:
                pix_spec = float(self.pix_spec)
                
                if screen_width == 3840 and screen_height == 2160:
                    scale_factor = 2.1
                else:
                    scale_factor = 1.3
                
                self.circle_radius = (self.um_size / pix_spec) * scale_factor / 2
                self.show_red_circle = True
                logging.debug(f"更新設置: pix_spec={self.pix_spec}, um_size={self.um_size}, circle_radius={self.circle_radius}")
            else:
                self.show_red_circle = False 
                logging.debug("未設置 pix_spec 或 um_size，紅圈將不顯示")

            if self.photo_files:
                self.show_photo(self.photo_files[0])
        logging.debug("um 尺寸更改完成")

    def update_um_size(self):
        if hasattr(self, 'pix_spec'):
            pix_spec = float(self.pix_spec)
            screen_resolution = QDesktopWidget().screenGeometry()
            screen_width = screen_resolution.width()
            screen_height = screen_resolution.height()
            
            if screen_width == 3840 and screen_height == 2160:
                scale_factor = 2.1
            else:
                scale_factor = 1.3
            
            self.um_size = (self.circle_radius * 2 * pix_spec) / scale_factor
            self.um_size = round(self.um_size, 2) 
        logging.debug(f"更新 um 尺寸: {self.um_size}")

    def show_photo_no_circle(self, photo_path):
        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        if screen_width >= 3840 and screen_height >= 2160:
            display_size = QSize(1024, 768)
            scaling_factor = 1
        else:
            display_size = QSize(800, 600)
            scaling_factor = 0.8

        image = QImage(photo_path)
        pixmap = QPixmap.fromImage(image)
        pixmap = pixmap.scaled(display_size * scaling_factor, Qt.KeepAspectRatio)
        return pixmap

    def classify_photo(self, class_name):
        if not self.photo_files:
            return

        current_photo = self.photo_files.pop(0)

        if self.save_folder and self.source_folder == self.original_source_folder:
            folder_name = self.target_folders[class_name].name
            target_folder = self.save_folder / folder_name
        else:
            target_folder = self.target_folders[class_name]

        if not target_folder.exists():
            target_folder.mkdir(parents=True, exist_ok=True)

        target_photo_path = target_folder / Path(current_photo).name

        if target_photo_path.exists():
            print(f"已跳過: {current_photo} (目標資料夾中已存在)")
        else:
            try:
                shutil.move(current_photo, target_folder)
            except Exception as e:
                print(f"移動檔案時發生錯誤: {e}")
                self.photo_files.insert(0, current_photo)
                return

        self.classified_photos.append((current_photo, target_photo_path))
        if len(self.classified_photos) > 20:
            self.classified_photos.pop(0)

        if not self.photo_files:
            for row in range(self.button_grid_layout.rowCount()):
                for col in range(self.button_grid_layout.columnCount()):
                    item = self.button_grid_layout.itemAtPosition(row, col)
                    if item and item.widget():
                        item.widget().hide()
            
            self.image_label.hide()
            self.um_input_button.hide()
            self.open_original_button.hide()
            self.remaining_photos_label.hide()
            self.save_path_button.hide()
            self.measure_area_button.hide()

            if self.next_folder_list:
                self.selected_folders.append(self.source_folder)
                valid_folders = [folder for folder in self.next_folder_list if folder not in self.selected_folders]

                if valid_folders:
                    folder_names = [folder.name for folder in valid_folders]
                    selected_folder, ok = QInputDialog.getItem(self, "選擇資料夾", "選擇下一個要分類的資料夾:", folder_names, 0, False)
                    if ok:
                        selected_folder_path = valid_folders[folder_names.index(selected_folder)]
                        self.source_folder = selected_folder_path
                        self.photo_files = self.load_photos_in_folder(self.source_folder)
                        
                        if hasattr(self, 'measurement_window'):
                            self.measurement_window.reset_shapes_transform()
                            self.measurement_window.clear_shapes()
                        
                        self.show_photo(self.photo_files[0])
                        self.update_remaining_photos_label()
                        self.update_measurement_ui_image()
                        self.setFocusOnMainWindow()

                        for name, folder in self.target_folders.items():
                            self.target_folders[name] = self.source_folder / folder.name

                        self.um_input_button.setEnabled(True)
                        self.um_input_button.show()
                        self.open_original_button.show()
                        self.remaining_photos_label.show()
                        self.save_path_button.show()
                        self.go_back_button.hide()
                        self.image_label.show()
                        self.measure_area_button.show()
                        self.reset_button_grid_layout()
                        self.classified_photos.clear()

                        return
            
            self.select_folder_button.setText("完成")
            self.button_grid_layout.setEnabled(False)
            self.select_folder_button.setEnabled(False)
            self.select_folder_button.show()
            self.go_back_button.hide()
            self.update_remaining_photos_label()
            return

        if hasattr(self, 'measurement_window'):
            self.measurement_window.reset_shapes_transform()
            self.measurement_window.clear_shapes()

        self.show_photo(self.photo_files[0])
        self.update_remaining_photos_label()
        self.go_back_button.show()

        self.update_measurement_ui_image()

        if hasattr(self, 'measurement_window'):
            self.measurement_window.zoom_slider.setValue(100)

    def go_back_photo(self):
        if self.classified_photos:
            last_photo, target_photo_path = self.classified_photos.pop()
            shutil.move(target_photo_path, last_photo)
            self.photo_files.insert(0, last_photo)
            self.show_photo(last_photo)
            self.update_remaining_photos_label()
            self.update_measurement_ui_image()

    def update_remaining_photos_label(self):
        remaining_photos = len(self.photo_files)
        self.remaining_photos_label.setText(f"剩餘照片數量: {remaining_photos}")

    def mousePressEvent(self, event):
        if self.image_label.geometry().contains(event.pos()):
            image_position = self.image_label.mapFromGlobal(event.globalPos())
            image_position_scaled = self.mapImageToScaled(image_position)
            
            if event.button() == Qt.LeftButton:
                self.circle_center = image_position_scaled
                self.show_photo(self.photo_files[0])
            elif event.button() == Qt.RightButton:
                self.is_resizing_circle = True
                self.resize_start_pos = event.pos()
                self.start_radius = self.circle_radius
            event.accept()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.image_label.geometry().contains(event.pos()):
            image_position = self.image_label.mapFromGlobal(event.globalPos())
            image_position_scaled = self.mapImageToScaled(image_position)
            
            if event.buttons() == Qt.LeftButton:
                self.circle_center = image_position_scaled
                self.show_photo(self.photo_files[0])
            elif event.buttons() == Qt.RightButton and self.is_resizing_circle:
                diff = event.pos() - self.resize_start_pos
                new_radius = self.start_radius + diff.x() / 2
                if new_radius > 0:
                    self.circle_radius = new_radius
                    self.update_um_size()
                self.show_photo(self.photo_files[0])
            event.accept()
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            self.is_resizing_circle = False
            self.resize_start_pos = None
        super().mouseReleaseEvent(event)

    def mapImageToScaled(self, image_position):
        image = QImage(self.photo_files[0])
        scaled_position = QPoint(image_position.x(), image_position.y())
        return scaled_position
    
    def is_point_in_circle(self, point):
        return (point.x() - self.circle_center.x()) ** 2 + (point.y() - self.circle_center.y()) ** 2 <= self.circle_radius ** 2

    def load_photos_in_folder(self, folder_path):
        photo_files = []
        for file in os.listdir(folder_path):
            if file.endswith(('.jpeg', '.jpg', '.png')):
                photo_files.append(os.path.join(folder_path, file))
        return photo_files

    def open_measurement_ui(self):
        self.measurement_window = MeasurementWindow(self)
        self.measurement_window.show()
        self.update_measurement_ui_image()

    def update_measurement_ui_image(self):
        if hasattr(self, 'measurement_window') and self.photo_files:
            self.measurement_window.clear_shapes()
            self.measurement_window.reset_shapes_transform()
            pixmap = self.show_photo_no_circle(self.photo_files[0])
            self.measurement_window.update_image(pixmap)

    def show_photo_no_circle(self, photo_path):
        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        if screen_width >= 3840 and screen_height >= 2160:
            display_size = QSize(1024, 768)
            scaling_factor = 1
        else:
            display_size = QSize(800, 600)
            scaling_factor = 0.8

        image = QImage(photo_path)
        pixmap = QPixmap.fromImage(image)
        pixmap = pixmap.scaled(display_size * scaling_factor, Qt.KeepAspectRatio)
        return pixmap

    def closeEvent(self, event):
        if hasattr(self, 'measurement_window'):
            self.measurement_window.close()
        event.accept()

    def save_log(self):
        try:
            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            username = match.group(1) if match else 'Unknown'
            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            log_folder = r'M:\BI_Database\Apps\Database\Apps_Log_History'
            log_file = os.path.join(log_folder, 'Log result.txt')
            
            log_message = f"{current_datetime},{username},Open,Classify\n"
            
            os.makedirs(log_folder, exist_ok=True)
            
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(log_message)
                
        except Exception as e:
            print(f"寫入log時發生錯誤: {e}")

    def check_version(self):
        try:
            app_folder = os.path.normpath(os.path.join("M:", "BI_Database", "Apps", "Database", "Apps_Installation_package", "Q_All"))
            exe_files = [os.path.join(app_folder, f) for f in os.listdir(app_folder) 
                        if f.startswith("Classify_V") and f.endswith(".exe")]

            if not exe_files:
                msg = QMessageBox()
                msg.setWindowTitle('未獲取啟動權限')
                msg.setText('未獲取啟動權限, 請申請M:\BI_Program權限, 並聯絡#1082 Racky')
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: #2A2A2A;
                        color: white;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QLabel {
                        color: white;
                        font-size: 9pt;
                        padding: 10px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton {
                        background-color: #424242;
                        color: white;
                        border: none;
                        border-radius: 6px;
                        min-width: 80px;
                        min-height: 25px;
                        padding: 5px 15px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton:hover {
                        background-color: #525252;
                    }
                    QMessageBox QPushButton:pressed {
                        background-color: #383838;
                    }
                """)
                msg.exec()
                sys.exit(1)

            def parse_version(version_str):
                match = re.search(r'_V(\d+)\.(\d+)', version_str)
                if match:
                    return tuple(map(int, match.groups()))
                return (0, 0)

            latest_version = max(parse_version(os.path.basename(f)) for f in exe_files)
            latest_exe = max((f for f in exe_files 
                            if parse_version(os.path.basename(f)) == latest_version), 
                        key=os.path.getmtime)

            current_version_match = re.search(r'_V(\d+)\.(\d+)', os.path.basename(sys.executable))
            if current_version_match:
                current_version = tuple(map(int, current_version_match.groups()))
            else:
                current_version = (43, 0)

            if current_version[0] != latest_version[0]:
                msg = QMessageBox()
                msg.setWindowTitle('進版通知')
                msg.setText(f'將更新至新版本 V{latest_version[0]}.{latest_version[1]}')
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.setIcon(QMessageBox.Icon.Information)
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: #2A2A2A;
                        color: white;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QLabel {
                        color: white;
                        font-size: 9pt;
                        padding: 10px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton {
                        background-color: #424242;
                        color: white;
                        border: none;
                        border-radius: 6px;
                        min-width: 80px;
                        min-height: 25px;
                        padding: 5px 15px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton:hover {
                        background-color: #525252;
                    }
                    QMessageBox QPushButton:pressed {
                        background-color: #383838;
                    }
                """)
                msg.exec()
                
                bat_content = '''@echo off
    timeout /t 2 /nobreak
    rmdir /s /q "C:\\BITools\\Classify"
    del "%~f0"
    '''
                bat_path = os.path.join(os.environ['TEMP'], 'delete_outside_fab.bat')
                with open(bat_path, 'w') as f:
                    f.write(bat_content)
                
                subprocess.Popen(['cmd', '/c', bat_path], shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                
                os.startfile(latest_exe)
                sys.exit(1)

            if current_version[0] != 43:
                msg = QMessageBox()
                msg.setWindowTitle('版本錯誤')
                msg.setText('軟件版本錯誤')
                msg.setStandardButtons(QMessageBox.StandardButton.Ok)
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setStyleSheet("""
                    QMessageBox {
                        background-color: #2A2A2A;
                        color: white;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QLabel {
                        color: white;
                        font-size: 9pt;
                        padding: 10px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton {
                        background-color: #424242;
                        color: white;
                        border: none;
                        border-radius: 6px;
                        min-width: 80px;
                        min-height: 25px;
                        padding: 5px 15px;
                        font-family: "微軟正黑體";
                    }
                    QMessageBox QPushButton:hover {
                        background-color: #525252;
                    }
                    QMessageBox QPushButton:pressed {
                        background-color: #383838;
                    }
                """)
                msg.exec()
                sys.exit(1)

        except FileNotFoundError:
            msg = QMessageBox()
            msg.setWindowTitle('未獲取啟動權限')
            msg.setText('未獲取啟動權限, 請申請M:\BI_Program權限, 並聯絡#1082 Racky')
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setStyleSheet("""
                QMessageBox {
                    background-color: #2A2A2A;
                    color: white;
                    font-family: "微軟正黑體";
                }
                QMessageBox QLabel {
                    color: white;
                    font-size: 9pt;
                    padding: 10px;
                    font-family: "微軟正黑體";
                }
                QMessageBox QPushButton {
                    background-color: #424242;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    min-width: 80px;
                    min-height: 25px;
                    padding: 5px 15px;
                    font-family: "微軟正黑體";
                }
                QMessageBox QPushButton:hover {
                    background-color: #525252;
                }
                QMessageBox QPushButton:pressed {
                    background-color: #383838;
                }
            """)
            msg.exec()
            sys.exit(1)

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 9)
    font.setBold(True)  # 設置為粗體
    app.setFont(font)
    app.setStyleSheet(STYLE_SHEET)
    window = ImageClassifier()
    window.show()
    sys.exit(app.exec_())
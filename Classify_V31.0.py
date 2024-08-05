from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog,
    QInputDialog, QHBoxLayout, QSizePolicy, QGridLayout, QMessageBox,
    QGraphicsScene, QGraphicsView, QGraphicsItem, QGraphicsRectItem, QGraphicsEllipseItem,
    QDesktopWidget, QSlider, QDialog, QGroupBox, QLineEdit
)
from PyQt5.QtGui import QPixmap, QFont, QImage, QPainter, QColor, QPen, QIcon, QIntValidator, QDoubleValidator, QPainterPath, QBrush
from PyQt5.QtCore import Qt, QPoint, QTimer, QSize, QRectF, QPointF
from pathlib import Path
from PyQt5 import QtGui
import os
import re
import shutil
import sys
import socket
import qtmodern.styles
import qtmodern.windows
import datetime
import py7zr
import tempfile
import math

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(application_path, 'format.ico')

class ResizableRotatableItem(QGraphicsItem):
    def __init__(self, x, y, width, height):
        super().__init__()
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemSendsGeometryChanges)
        self._brush = QBrush(QColor(255, 0, 0, 105))
        self._pen = QPen(Qt.red, 2)
        self.handle_size = 8
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

    def boundingRect(self):
        return self.rect.adjusted(-self.handle_size, -self.handle_size - self.rotate_handle_distance,
                                  self.handle_size, self.handle_size)

    def shape(self):
        path = QPainterPath()
        path.addRect(self.rect)
        for handle in self.handles + [self.rotate_handle]:
            path.addRect(self.handle_rect(handle))
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
                painter.drawRect(self.handle_rect(handle))
            
            painter.setBrush(Qt.blue)
            painter.drawEllipse(self.handle_rect(self.rotate_handle))
            painter.drawLine(self.rect.center(), self.rotate_handle)
        
        painter.restore()

    def draw_shape(self, painter):
        pass

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

    def mousePressEvent(self, event):
        self.start_pos = event.pos()
        self.original_rect = self.rect
        self.original_angle = self.rotation_angle
        
        if self.handle_rect(self.rotate_handle).contains(event.pos()):
            self.active_handle = 'rotate'
            event.accept()
            return
        
        for i, handle in enumerate(self.handles):
            if self.handle_rect(handle).contains(event.pos()):
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
        self.update()
        event.accept()

    def mouseReleaseEvent(self, event):
        self.active_handle = None
        super().mouseReleaseEvent(event)
        self.prepareGeometryChange()
        self.update()
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
        if change == QGraphicsItem.ItemPositionChange and self.scene():
            return value
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
        self._brush = QBrush(QColor(0, 0, 255, 105))
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
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
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
            self.scale_factor = value / 100.0
            self.view.resetTransform()
            self.view.scale(self.scale_factor, self.scale_factor)
            self.view.centerOn(self.pixmap_item)
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
        
    def initUI(self):
        layout = QVBoxLayout()
        layout.setSpacing(10) 
        
        group_style = """
        QGroupBox {
            background-color: #2D2D2D;
            border: 1px solid #555555;
            border-radius: 5px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top center;
            padding: 0 5px;
            color: white;
            font-weight: bold;
        }
        """

        pix_spec_group = QGroupBox("")
        pix_spec_group.setStyleSheet(group_style)
        pix_spec_layout = QHBoxLayout()
        self.pix_spec_input = QLineEdit()
        pix_spec_layout.addWidget(QLabel("請輸入pix spec值:"))
        pix_spec_layout.addWidget(self.pix_spec_input)
        pix_spec_group.setLayout(pix_spec_layout)

        um_group = QGroupBox("")
        um_group.setStyleSheet(group_style)
        um_layout = QHBoxLayout()
        self.um_input = QLineEdit()
        um_layout.addWidget(QLabel("請輸入µm值:"))
        um_layout.addWidget(self.um_input)
        um_group.setLayout(um_layout)
        
        self.param_link = QLabel()
        self.param_link.setAlignment(Qt.AlignCenter) 
        self.param_link.setStyleSheet("""
            QLabel {
                qproperty-alignment: AlignCenter;
                padding: 5px;
            }
        """)
        self.param_link.setText('<a href="file:///D:/本地應用程式/Classify/AVI Resolution 對照表.xlsx" style="color: #DFD7D5;">參數對照表,請點選連結</a>')
        self.param_link.setOpenExternalLinks(True)
        
        layout.addWidget(pix_spec_group)
        layout.addWidget(um_group)
        layout.addWidget(self.param_link)
        
        confirm_button = QPushButton("確認")
        confirm_button.clicked.connect(self.accept)
        layout.addWidget(confirm_button)
        
        self.setLayout(layout)
        self.setWindowTitle("設置參數")

        screen = QDesktopWidget().screenNumber(self)
        screen_resolution = QDesktopWidget().screenGeometry(screen)
        width, height = screen_resolution.width(), screen_resolution.height()
        
        if width == 3840 and height == 2160:
            self.setFixedSize(500, 300)
        else:
            self.setFixedSize(400, 230)

    def get_settings(self):
        pix_spec = self.pix_spec_input.text()
        um_size = self.um_input.text()
        
        return pix_spec, um_size

class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setWindowTitle('Classify')
        app.setWindowIcon(QIcon(icon_path))
        self.log_file = None
        self.user_id = None
        self.classified_photos = []

        # 呼叫 check_version 方法
        self.check_version()
        self.save_log()

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
            button_width_factor = 1.1
            button_height_factor = 0.14
        else:
            button_width_factor = 0.84
            button_height_factor = 0.07

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

        self.um_input_button = QPushButton('量測Defect size (μm)' , self)
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
            self.setGeometry(350, 600, 250, 300)
        else:
            self.setGeometry(100, 200, 200, 150)

        self.resize_timer = QTimer(self)
        self.resize_timer.timeout.connect(self.update_circle_radius)

        main_layout.addLayout(left_panel)
        main_layout.addLayout(right_panel)

        self.original_um_value = 0  # 初始化原始um值

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
            class_name = shortcuts[key]
            self.classify_photo(class_name)
        else:
            super().keyPressEvent(event)

    def select_save_path(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec_():
            save_path = dialog.selectedFiles()[0]
            self.save_folder = Path(save_path)

    def write_close_log(self):
        try:
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.log_file, 'a') as file:
                file.write(f"{current_time} {self.user_id} Close\n")
        except:
            pass

    def change_um_size(self):
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
            
            # 計算 self.circle_radius
            if self.pix_spec and self.um_size:
                
                if screen_width == 3840 and screen_height == 2160:
                    scale_factor = 1.8
                else:
                    scale_factor = 1.1

                self.circle_radius = (float(self.pix_spec) * self.um_size * scale_factor) / 2 

            self.show_photo(self.photo_files[0])

    def set_circle_radius(self, radius):
        self.circle_radius = radius
        self.show_photo(self.photo_files[0])

    def open_original_image(self):
        if self.photo_files:
            photo_path = self.photo_files[0]
            if os.path.exists(photo_path):
                os.system(f'start "" "{photo_path}"')

    def select_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)  
        dialog.setOption(QFileDialog.DontUseNativeDialog, True)
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
            button.clicked.connect(lambda checked, name=class_name: self.classify_photo(name))
            self.button_grid_layout.addWidget(button, row, col)

            col += 1
            if col > 1:
                col = 0
                row += 1

    def set_button_shortcut(self, button, class_name):
        shortcuts = {
            'ugly die': ('I', 'ugly die ( I )'),
            'foreign material': ('E', 'foreign material (  E )'),
            'Particle': ('U', 'Particle ( U )'),
            'Probe mark shift': ('R', 'Probe mark shift ( R )'),
            'Over kill': ('Y', 'Over kill ( Y )'),
            'Process defect': ('Q', 'Process defect ( Q )'),
            'Other': ('A', 'Other ( A )'),
            'Al particle out of pad': ('P', 'Al particle out of pad ( P )'),
            'Al particle': ('J', 'Al particle ( J )'),
            'Wafer Scratch': ('W', 'Wafer Scratch ( W )'),
            'PM area out spec': ('F', 'PM area out spec ( F )'),
            'Bump PM shift': ('L', 'Bump PM shift ( L )'),
            'Bump scratch': ('Z', 'Bump scratch ( Z )'),
            'Bump foreign Material': ('X', 'Bump foreign Material ( X )'),
            'Bump PM diameter out of spec': ('D', 'Bump PM diameter out of spec ( D )'),
            'PM No. Out Spec': ('T', 'PM No. Out Spec ( T )'),
            'Pad discoloration': ('O', 'Pad discoloration ( O )'),
            'Irregular bump': ('S', 'Irregular bump ( S )'),
            'Probing Void': ('G', 'Probing Void ( G )'),
            'Missing Probe Mark': ('H', 'Missing Probe Mark ( H )'),
            'Surface(Incoming defect)': ('K', 'Surface(Incoming defect) ( K )'),
            'Missing bump': ('C', 'Missing bump ( C )'),
            'Bump residue': ('V', 'Bump residue ( V )'),
            'Bump house defect': ('B', 'Bump house defect ( B )'),
            'Large defect': ('N', 'Large defect ( N )'),
            'Large bump': ('1', 'Large bump ( 1 )'),
            'small bump': ('2', 'small bump ( 2 )'),
            '380 special PM shift': ('3', '380 special PM shift ( 3 )'),
        }
        if class_name in shortcuts:
            shortcut, text = shortcuts[class_name]
            button.setShortcut(shortcut)
            button.setText(text)

    def show_photo(self, photo_path):
        # 獲取螢幕解析度
        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        # 根據螢幕解析度調整 display_size
        if screen_width >= 3840 and screen_height >= 2160:
            display_size = QSize(1024, 768)
            scaling_factor = 1.4
        else:
            display_size = QSize(800, 600)
            scaling_factor = 1.1

        # 加載和調整圖片
        image = QImage(photo_path)
        pixmap = QPixmap.fromImage(image)
        pixmap = pixmap.scaled(display_size * scaling_factor, Qt.KeepAspectRatio)
        self.image_label.setPixmap(pixmap)
        
        # 繪製圓形
        circle_pixmap = pixmap.copy()
        painter = QPainter(circle_pixmap)
        pen = QPen(self.circle_color)
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawEllipse(self.circle_center, self.circle_radius, self.circle_radius)
        painter.end()

        # 更新圖片
        self.image_label.setPixmap(circle_pixmap)
        self.update_measurement_ui_image()


    def show_photo_no_circle(self, photo_path):
        # 獲取螢幕解析度
        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        # 根據螢幕解析度調整 display_size
        if screen_width >= 3840 and screen_height >= 2160:
            display_size = QSize(1024, 768)
            scaling_factor = 1
        else:
            display_size = QSize(800, 600)
            scaling_factor = 0.8

        # 加載和調整圖片
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
                # 如果移動失敗，將照片放回列表
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
                        
                        # 新添加的代碼
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

        # 確保 Measurement UI 的縮放設置被重置
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

    def update_circle_radius(self):
        mouse_pos = self.image_label.mapFromGlobal(QtGui.QCursor.pos())
        distance = (mouse_pos - self.circle_center).manhattanLength()
        new_radius = distance
        self.set_circle_radius(new_radius)

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resize_timer.start(100)

        if self.image_label.geometry().contains(event.pos()):
            image_position = self.image_label.mapFromGlobal(QtGui.QCursor.pos())
            image_position_scaled = self.mapImageToScaled(image_position)
            image_position_scaled.setX(image_position_scaled.x() - 0)
            self.circle_center = image_position_scaled
            self.show_photo(self.photo_files[0])
            self.update()

    def mouseMoveEvent(self, event):
        if self.resize_timer.isActive():
            image_position = self.image_label.mapFromGlobal(QtGui.QCursor.pos())
            image_position_scaled = self.mapImageToScaled(image_position)
            distance = (image_position_scaled - self.circle_center).manhattanLength()
            new_radius = distance
            self.set_circle_radius(new_radius)

        if self.image_label.geometry().contains(event.pos()):
            if event.buttons() == Qt.LeftButton:
                image_position = self.image_label.mapFromGlobal(QtGui.QCursor.pos())
                image_position_scaled = self.mapImageToScaled(image_position)
                image_position_scaled.setX(image_position_scaled.x() - 0)
                self.circle_center = image_position_scaled
                self.show_photo(self.photo_files[0])
                self.update()

    def mapImageToScaled(self, image_position):
        image = QImage(self.photo_files[0])
        scaled_position = QPoint(image_position.x(), image_position.y())
        return scaled_position

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resize_timer.stop()

    def load_photos_in_folder(self, folder_path):
        photo_files = []
        for file in os.listdir(folder_path):
            if file.endswith(('.jpeg', '.jpg', '.png')):
                photo_files.append(os.path.join(folder_path, file))
        return photo_files

    def save_log(self):
        try:
            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            username = match.group(1) if match else 'Unknown'

            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_folder = r'M:\QA_Program_Raw_Data\Log History'
            archive_path = os.path.join(log_folder, 'Classify.7z')
            log_filename = f'{username}.txt'
            new_log_message = f"{current_datetime} {username} Open\n"
            os.makedirs(log_folder, exist_ok=True)

            if not os.path.exists(archive_path):
                with py7zr.SevenZipFile(archive_path, mode='w', password='@Joe11111111') as archive:
                    archive.writestr(new_log_message, f'Classify/{log_filename}')
            else:
                log_content = ""
                files_to_keep = []

                with py7zr.SevenZipFile(archive_path, mode='r', password='@Joe11111111') as archive:
                    for filename, bio in archive.read().items():
                        if filename == f'Classify/{log_filename}':
                            log_content = bio.read().decode('utf-8')
                        else:
                            files_to_keep.append((filename, bio.read()))

                if new_log_message not in log_content:
                    log_content += new_log_message

                with tempfile.NamedTemporaryFile(delete=False, suffix='.7z') as temp_file:
                    temp_archive_path = temp_file.name

                with py7zr.SevenZipFile(temp_archive_path, mode='w', password='@Joe11111111') as archive:
                    archive.writestr(log_content.encode('utf-8'), f'Classify/{log_filename}')
                    for filename, content in files_to_keep:
                        archive.writestr(content, filename)

                shutil.move(temp_archive_path, archive_path)

        except Exception as e:
            print(f"寫入log時發生錯誤: {e}")

    def check_version(self):
        try:
            app_folder = r"M:\QA_Program_Raw_Data\Apps"
            exe_files = [f for f in os.listdir(app_folder) if f.startswith("Classify_V") and f.endswith(".exe")]

            if not exe_files:
                QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                sys.exit(1)

            # 修改版本號提取邏輯，只取主版本號
            latest_version = max(int(re.search(r'_V(\d+)', f).group(1)) for f in exe_files)

            # 修改當前版本號提取邏輯，只取主版本號
            current_version_match = re.search(r'_V(\d+)', os.path.basename(sys.executable))
            if current_version_match:
                current_version = int(current_version_match.group(1))
            else:
                current_version = 0

            if current_version < latest_version:
                QMessageBox.information(self, '請更新至最新版本', '請更新至最新版本')
                os.startfile(app_folder)  # 開啟指定的資料夾
                sys.exit(0)

            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            if match:
                username = match.group(1)
                if username == "A000000":
                    QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                    sys.exit(1)
            else:
                QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                sys.exit(1)

        except FileNotFoundError:
            QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
            sys.exit(1)

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
        # 獲取螢幕解析度
        screen_resolution = QDesktopWidget().screenGeometry()
        screen_width = screen_resolution.width()
        screen_height = screen_resolution.height()

        # 根據螢幕解析度調整 display_size
        if screen_width >= 3840 and screen_height >= 2160:
            display_size = QSize(1024, 768)
            scaling_factor = 1
        else:
            display_size = QSize(800, 600)
            scaling_factor = 0.8

        # 加載和調整圖片
        image = QImage(photo_path)
        pixmap = QPixmap.fromImage(image)
        pixmap = pixmap.scaled(display_size * scaling_factor, Qt.KeepAspectRatio)
        return pixmap

    def closeEvent(self, event):
        if hasattr(self, 'measurement_window'):
            self.measurement_window.close()
        event.accept()

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 9)
    font.setBold(True)  # 設置為粗體
    app.setFont(font)
    qtmodern.styles.dark(app)
    window = ImageClassifier()  
    win = qtmodern.windows.ModernWindow(window)
    win.show()
    sys.exit(app.exec_())
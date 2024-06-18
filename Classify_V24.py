from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog,
    QInputDialog, QVBoxLayout, QHBoxLayout, QSizePolicy, QGridLayout, QMessageBox,
    QGraphicsScene, QGraphicsView, QGraphicsItem, QGraphicsRectItem, QGraphicsEllipseItem, QDesktopWidget,
    QGraphicsPixmapItem, QDesktopWidget
)
from PyQt5.QtGui import QPixmap, QFont, QImage, QPainter, QColor, QPen, QIcon
from PyQt5.QtCore import Qt, QPoint, QTimer, QSize, QRectF
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

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(application_path, 'format.ico')

class ResizableRectItem(QGraphicsRectItem):
    def __init__(self, x, y, width, height):
        super().__init__(x, y, width, height)
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemSendsGeometryChanges)
        self.setBrush(QColor(255, 0, 0, 150))
        self.resizing = False

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resizing = True
        else:
            self.setSelected(True)  # 點擊圖形時設置為選中狀態
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            self.adjust_size(event)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resizing = False
        super().mouseReleaseEvent(event)

    def adjust_size(self, event):
        new_rect = QRectF(self.rect())
        if event.pos().x() < self.rect().left():
            new_rect.setLeft(event.pos().x())
        else:
            new_rect.setRight(event.pos().x())
        if event.pos().y() < self.rect().top():
            new_rect.setTop(event.pos().y())
        else:
            new_rect.setBottom(event.pos().y())
        self.setRect(new_rect)


class ResizableEllipseItem(QGraphicsEllipseItem):
    def __init__(self, x, y, width, height):
        super().__init__(x, y, width, height)
        self.setFlags(QGraphicsItem.ItemIsSelectable | QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemSendsGeometryChanges)
        self.setBrush(QColor(0, 0, 255, 150))
        self.resizing = False

    def mousePressEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resizing = True
        else:
            self.setSelected(True)  # 點擊圖形時設置為選中狀態
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.resizing:
            self.adjust_size(event)
        else:
            super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.RightButton:
            self.resizing = False
        super().mouseReleaseEvent(event)

    def adjust_size(self, event):
        new_rect = QRectF(self.rect())
        if event.pos().x() < self.rect().left():
            new_rect.setLeft(event.pos().x())
        else:
            new_rect.setRight(event.pos().x())
        if event.pos().y() < self.rect().top():
            new_rect.setTop(event.pos().y())
        else:
            new_rect.setBottom(event.pos().y())
        self.setRect(new_rect)

class MeasurementWindow(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.shapes = []
        self.initUI()
        self.setWindowTitle('Measurement UI')
        self.resize(int(main_window.width() * 0.6), int(main_window.height() * 1.1))
        main_window.geometry().top() + (main_window.height() - self.height()) // 2

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
        
        self.close_button = QPushButton('關閉UI', self)
        self.close_button.clicked.connect(self.close)
        self.button_layout.addWidget(self.close_button)
        
        self.layout.addLayout(self.button_layout)

        self.setup_scene()

    def update_shapes(self):
        for shape in self.shapes:
            shape.setPos(self.image_label.pos())

    def draw_rectangle(self):
        rect_item = ResizableRectItem(50, 50, 200, 150)
        rect_item.setZValue(1)  # 設置矩形的 Z 值為 1
        self.scene.addItem(rect_item)
        self.shapes.append(rect_item)

    def draw_circle(self):
        ellipse_item = ResizableEllipseItem(100, 100, 150, 150)
        ellipse_item.setZValue(1)  # 設置圓形的 Z 值為 1
        self.scene.addItem(ellipse_item)
        self.shapes.append(ellipse_item)

    def calculate_area_ratio(self):
        if len(self.shapes) < 2:
            QMessageBox.information(self, '提示', '請正確圈選出需要量測的probe mark')
            return
        areas = [shape.rect().width() * shape.rect().height() if isinstance(shape, QGraphicsRectItem) else shape.rect().width() * shape.rect().height() * 3.14159 / 4 for shape in self.shapes]
        min_area = min(areas)
        max_area = max(areas)
        ratio = min_area / max_area * 100
        QMessageBox.information(self, '面積佔比', f'最小圖形面積與最大圖形面積的比值: {ratio:.2f}%')

    def update_image(self, pixmap):
        self.image_label.setPixmap(pixmap)
        self.scene.setSceneRect(QRectF(pixmap.rect()))
        
        if not self.shapes:
            self.scene.clear()
            self.scene.addPixmap(pixmap)
        else:
            for item in self.scene.items():
                if isinstance(item, QGraphicsPixmapItem):
                    self.scene.removeItem(item)
                    break
            self.scene.addPixmap(pixmap)
            
            for shape in self.shapes:
                shape.setPos(shape.pos().x(), shape.pos().y())

    def setup_scene(self):
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene, self)
        self.view.setGeometry(self.image_label.geometry())
        self.view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.view.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.layout.addWidget(self.view)
        self.view.show()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete:
            selected_items = self.scene.selectedItems()
            for item in selected_items:
                if item in self.shapes:
                    self.shapes.remove(item)
                    self.scene.removeItem(item)

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
            button_width_factor = 0.7
            button_height_factor = 0.08

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

        self.select_folder_button = QPushButton('分類Defect Image', self)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.select_folder_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.select_folder_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        right_panel.addWidget(self.select_folder_button)

        self.um_input_button = QPushButton('測量尺寸(請以TPAS尺寸作為對照)', self)
        self.um_input_button.clicked.connect(self.change_um_size)
        self.um_input_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.um_input_button.setFixedSize(int(self.width() * button_width_factor), int(self.height() * button_height_factor))
        self.um_input_button.setEnabled(False)
        self.um_input_button.hide()
        right_panel.addWidget(self.um_input_button)

        self.measure_area_button = QPushButton('量測Probe mark area', self)
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
            self.setGeometry(350, 800, 250, 300)
        else:
            self.setGeometry(250, 400, 200, 150)

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
            Qt.Key_O: 'Other',
            Qt.Key_P: 'Al particle out of pad',
            Qt.Key_J: 'Al particle',
            Qt.Key_W: 'Wafer Scratch',
            Qt.Key_F: 'PM area out spec',
            Qt.Key_L: 'Bump PM shift',
            Qt.Key_Z: 'Bump scratch',
            Qt.Key_X: 'Bump foreign Material',
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
        um_value, ok = QInputDialog.getInt(self, "變更尺寸", "輸入變更值1~999:")
        if ok:
            self.original_um_value = um_value  # 更新原始um值

            # 獲取螢幕解析度
            screen_resolution = QDesktopWidget().screenGeometry()
            screen_width = screen_resolution.width()
            screen_height = screen_resolution.height()

            # 根據螢幕解析度調整um大小
            if screen_width >= 3840 and screen_height >= 2160:
                self.set_circle_radius(um_value / 1.9)  # 變更um大小
            else:
                self.set_circle_radius(um_value / 3)  # 變更um大小

            self.update_um_input_button_text()


    def set_circle_radius(self, radius):
        self.circle_radius = radius
        self.update_um_input_button_text()
        self.show_photo(self.photo_files[0])

    def open_original_image(self):
        if self.photo_files:
            photo_path = self.photo_files[0]
            if os.path.exists(photo_path):
                os.system(f'start "" "{photo_path}"')

    def select_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec_():
            folder_path = dialog.selectedFiles()[0]
            self.source_folder = Path(folder_path)
            self.original_source_folder = self.source_folder

            folder_names = {
                'ugly die': '151_Ugly_Die(2D)',
                'foreign material': '102_Foreign_material(09)',
                'Particle': '117_Particle(16)',
                'Probe mark shift': '200_Probe_Mark_Shift(10)',
                'Bump foreign Material': '502_Bump foreign Material(25)',
                'Over kill': '0_Over_kill(15)',
                'Process defect': '100_Process_Defect(07)',
                'Al particle out of pad': '205_Al_particle_out_of_pad(0F)',
                'Al particle': '0_Al_particle(18)',
                'Wafer Scratch': '101_Wafer Scratch(08)',
                'PM area out spec': '202_PM area out spec.(12)',
                'Bump PM shift': '500_Bump PM shift(23)',
                'Bump scratch': '501_Bump scratch(24)',
                'Other': '186_Other(BA)'
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
            'ugly die': ('I', 'ugly die ( 快捷鍵: I )'),
            'foreign material': ('E', 'foreign material ( 快捷鍵: E )'),
            'Particle': ('U', 'Particle ( 快捷鍵: U )'),
            'Probe mark shift': ('R', 'Probe mark shift ( 快捷鍵: R )'),
            'Over kill': ('Y', 'Over kill ( 快捷鍵: Y )'),
            'Process defect': ('Q', 'Process defect ( 快捷鍵: Q )'),
            'Other': ('O', 'Other ( 快捷鍵: O )'),
            'Al particle out of pad': ('P', 'Al particle out of pad ( 快捷鍵: P )'),
            'Al particle': ('J', 'Al particle ( 快捷鍵: J )'),
            'Wafer Scratch': ('W', 'Wafer Scratch ( 快捷鍵: W )'),
            'PM area out spec': ('F', 'PM area out spec ( 快捷鍵: F )'),
            'Bump PM shift': ('L', 'Bump PM shift ( 快捷鍵: L )'),
            'Bump scratch': ('Z', 'Bump scratch ( 快捷鍵: Z )'),
            'Bump foreign Material': ('X', 'Bump foreign Material ( 快捷鍵: X )'),
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
            scaling_factor = 1
        else:
            display_size = QSize(800, 600)
            scaling_factor = 0.8

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
            print(f"Skipped: {current_photo} (already exists in target folder)")
        else:
            shutil.move(current_photo, target_folder)

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
                    selected_folder, ok = QInputDialog.getItem(self, "Select Folder", "Choose the next folder to classify:", folder_names, 0, False)
                    if ok:
                        selected_folder_path = valid_folders[folder_names.index(selected_folder)]
                        self.source_folder = selected_folder_path
                        self.photo_files = self.load_photos_in_folder(self.source_folder)
                        self.show_photo(self.photo_files[0])
                        self.update_remaining_photos_label()
                        self.update_measurement_ui_image()
                        self.setFocusOnMainWindow()
                        self.update_measurement_ui_image()

                        for name, folder in self.target_folders.items():
                            self.target_folders[name] = self.source_folder / folder.name

                        self.um_input_button.setEnabled(True)
                        self.um_input_button.show()
                        self.open_original_button.show()
                        self.remaining_photos_label.show()
                        self.save_path_button.show()
                        self.go_back_button.hide()  # 禁用go_back_button
                        self.image_label.show()
                        self.measure_area_button.show()
                        self.reset_button_grid_layout()
                        self.classified_photos.clear()  # 清空分類過的照片列表

                        return
            
            self.select_folder_button.setText("Complete")
            self.button_grid_layout.setEnabled(False)
            self.select_folder_button.setEnabled(False)
            self.select_folder_button.show()
            self.go_back_button.hide()
            self.update_remaining_photos_label()
            return

        self.show_photo(self.photo_files[0])
        self.update_remaining_photos_label()
        self.go_back_button.show()  # 啟用go_back_button

        um_size = self.um_input_button.text()
        try:
            self.circle_radius = int(um_size) / 2
            self.show_photo(self.photo_files[0])
        except ValueError:
            pass

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
        self.update_um_input_button_text()

    def update_um_input_button_text(self):
        um_value = self.original_um_value  # 使用原始um值來顯示
        self.um_input_button.setText(f'變更尺寸 (目前值: {um_value:.2f} )')

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
            self.update_um_input_button_text()

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
            self.update_um_input_button_text()

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
            if match:
                username = match.group(1)
            else:
                username = 'Unknown'

            current_datetime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_folder =r'M:\QA_Program_Raw_Data\Log History\Classify'
            os.makedirs(log_folder, exist_ok=True)
            log_file = os.path.join(log_folder, f"{username}.txt")

            log_message = f"{current_datetime} {username} Open\n"
            with open(log_file, 'a') as file:
                file.write(log_message)

        except Exception as e:
            print(f"寫入log時發生錯誤: {e}")
            pass

    def check_version(self):
        try:
            app_folder = r"M:\QA_Program_Raw_Data\Apps"
            exe_files = [f for f in os.listdir(app_folder) if f.startswith("Classify_V") and f.endswith(".exe")]

            if not exe_files:
                QMessageBox.warning(self, '未獲取啟動權限', '未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky')
                sys.exit(1)

            latest_version = max(int(re.search(r'_V(\d+)\.exe', f).group(1)) for f in exe_files)

            current_version_match = re.search(r'_V(\d+)\.exe', os.path.basename(sys.executable))
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
            pixmap = self.show_photo_no_circle(self.photo_files[0])
            self.measurement_window.update_image(pixmap)


if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 10)
    app.setFont(font)
    qtmodern.styles.dark(app)
    window = ImageClassifier()  
    win = qtmodern.windows.ModernWindow(window)
    win.show()
    sys.exit(app.exec_())
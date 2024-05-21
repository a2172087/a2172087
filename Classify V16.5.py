from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QInputDialog, QVBoxLayout, QHBoxLayout, QSizePolicy, QGridLayout, QMessageBox
from PyQt5.QtGui import QPixmap, QFont, QImage, QPainter, QColor, QPen, QIcon
from PyQt5.QtCore import Qt, QPoint, QTimer, QSize
from pathlib import Path
from PyQt5 import QtGui
import os
import shutil
import sys
import qtmodern.styles
import qtmodern.windows

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(application_path, 'format.ico')

class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.setWindowTitle('Classify')
        app.setWindowIcon(QIcon(icon_path))

    def initUI(self):
        main_layout = QHBoxLayout(self)

        # 左側面板
        left_panel = QVBoxLayout()
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        left_panel.addWidget(self.image_label)

        # 右側面板
        right_panel = QVBoxLayout()

        self.open_original_button = QPushButton('開啟圖片原始檔\n(快捷鍵 : 空白鍵)', self)
        self.open_original_button.clicked.connect(self.open_original_image)
        self.open_original_button.setShortcut(' ')
        self.open_original_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.open_original_button.setFixedSize(600, 120)  # 調整寬度和高度的值以滿足您的需求
        self.open_original_button.hide()
        right_panel.addWidget(self.open_original_button)

        self.select_folder_button = QPushButton('選擇資料夾', self)
        self.select_folder_button.clicked.connect(self.select_folder)
        self.select_folder_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.select_folder_button.setFixedSize(500, 50)  # 調整寬度和高度的值以滿足您的需求
        right_panel.addWidget(self.select_folder_button)

        self.um_input_button = QPushButton('測量尺寸(請以TPAS尺寸作為對照)', self)
        self.um_input_button.clicked.connect(self.change_um_size)
        self.um_input_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.um_input_button.setFixedSize(600, 120)  # 調整寬度和高度的值以滿足您的需求
        self.um_input_button.setEnabled(False)  # 初始禁用
        self.um_input_button.hide()
        right_panel.addWidget(self.um_input_button)

        # 新增：按鈕的GridLayout
        self.button_grid_layout = QGridLayout()
        right_panel.addLayout(self.button_grid_layout)

        self.source_folder = None
        self.target_folders = {}

        self.circle_center = QPoint(250, 250)
        self.circle_radius = 50
        self.circle_color = QColor(255, 0, 0, 150)

        # 設定視窗位置
        screen_resolution = app.desktop().screenGeometry()
        if screen_resolution.width() == 3840 and screen_resolution.height() == 2160:
            self.setGeometry(350, 800, 250, 300)
        else:
            self.setGeometry(250, 200, 500, 300)

        self.resize_timer = QTimer(self)
        self.resize_timer.timeout.connect(self.update_circle_radius)

        main_layout.addLayout(left_panel)  # 左側面板添加到主佈局
        main_layout.addLayout(right_panel)  # 右側面板添加到主佈局

        # 使用QVBoxLayout來排列按鈕
        button_layout = QVBoxLayout()
        self.classification_buttons = {}
        for class_name in self.target_folders.keys():
            button = QPushButton('', self)
            self.set_button_shortcut(button, class_name)
            button.clicked.connect(lambda checked, name=class_name: self.classify_photo(name))
            button_layout.addWidget(button)
            self.classification_buttons[class_name] = button
        right_panel.addLayout(button_layout)  # 將button_layout添加到right_panel

        self.original_um_value = 0  # 初始化原始um值

    def change_um_size(self):
        um_value, ok = QInputDialog.getInt(self, "變更尺寸", "輸入變更值1~999:")
        if ok:
            self.original_um_value = um_value  # 更新原始um值
            self.set_circle_radius(um_value / 2.2) #變更um大小
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

            folder_names = {
                'ugly die': '2D_Ugly_Die',
                'foreign material': '09_Foreign_material',
                'Particle': '16_Particle',
                'Probe mark shift': '10_Probe_Mark_Shift',
                'Over kill': '15_Over_kill',
                'Process defect': '07_Process_Defect',
                'Al particle out of pad': '0F_Al_particle_out_of_pad',
                'Al particle': '18_Al_particle',
                'Wafer Scratch': '00_Wafer Scratch',
                'PM area out spec': '00_PM area out spec',
                'Bump PM shift': '00_Bump PM shift',
                'Bump scratch': '00_Bump scratch',
                'Other': '00_Other'
            }

            # 這裡只是儲存了資料夾的目標路徑，而不是實際創建資料夾
            self.target_folders = {name: self.source_folder / folder for name, folder in folder_names.items()}

            # 載入資料夾中的所有照片檔案，這一步不涉及創建資料夾
            self.photo_files = self.load_photos_in_folder(self.source_folder)

            if self.photo_files:
                self.um_input_button.setEnabled(True)
                self.um_input_button.show()
                self.open_original_button.show()

                # 重設GridLayout按鈕，這裡假設您已經有了相應的邏輯
                self.reset_button_grid_layout()

                # 顯示第一張照片
                self.show_photo(self.photo_files[0])
                
                # 隱藏選擇資料夾按鈕，因為此時用戶已選擇了資料夾
                self.select_folder_button.hide()


    def reset_button_grid_layout(self):
        # 清理GridLayout中的所有按鈕
        for i in reversed(range(self.button_grid_layout.count())): 
            widget_to_remove = self.button_grid_layout.itemAt(i).widget()
            self.button_grid_layout.removeWidget(widget_to_remove)
            widget_to_remove.setParent(None)

        row, col = 0, 0
        for class_name in self.target_folders.keys():
            button = QPushButton(class_name, self)
            self.set_button_shortcut(button, class_name)
            button.clicked.connect(lambda checked, name=class_name: self.classify_photo(name))
            
            # 將按鈕添加到GridLayout中
            self.button_grid_layout.addWidget(button, row, col)

            col += 1
            if col > 1:  # 根據需要調整列數
                col = 0
                row += 1


    def set_button_shortcut(self, button, class_name):
        if class_name == 'ugly die':
            button.setText('ugly die ( 快捷鍵: I )')
            button.setShortcut('I')
        elif class_name == 'foreign material':
            button.setText('foreign material ( 快捷鍵: E )')
            button.setShortcut('E')
        elif class_name == 'Particle':
            button.setText('Particle ( 快捷鍵: U )')
            button.setShortcut('U')
        elif class_name == 'Probe mark shift':
            button.setText('Probe mark shift ( 快捷鍵: R )')
            button.setShortcut('R')
        elif class_name == 'Over kill':
            button.setText('Over kill ( 快捷鍵: Y )')
            button.setShortcut('Y')
        elif class_name == 'Process defect':
            button.setText('Process defect ( 快捷鍵: Q )')
            button.setShortcut('Q')
        elif class_name == 'Other':
            button.setText('Other ( 快捷鍵: O )')
            button.setShortcut('O')
        elif class_name == 'Al particle out of pad':
            button.setText('Al particle out of pad ( 快捷鍵: P )')
            button.setShortcut('P')
        elif class_name == 'Al particle':
            button.setText('Al particle ( 快捷鍵: J )')
            button.setShortcut('J')
        elif class_name == 'Wafer Scratch':
            button.setText('Wafer Scratch ( 快捷鍵: W )')
            button.setShortcut('W')       
        elif class_name == 'PM area out spec':
            button.setText('PM area out spec ( 快捷鍵: F )')
            button.setShortcut('F')   
        elif class_name == 'Bump PM shift':
            button.setText('Bump PM shift ( 快捷鍵: L )')
            button.setShortcut('L')   
        elif class_name == 'Bump scratch':
            button.setText('Bump scratch ( 快捷鍵: Z )')
            button.setShortcut('Z')   

    def show_photo(self, photo_path):
        image = QImage(photo_path)
        pixmap = QPixmap.fromImage(image)
        #將尺寸固定為1024x768
        display_size = QSize(1024, 768)
        #將照片大小放大n倍(可調整0~3)
        pixmap = pixmap.scaled(display_size * 0.8, Qt.KeepAspectRatio)

        circle_pixmap = pixmap.copy()
        painter = QPainter(circle_pixmap)
        pen = QPen(self.circle_color)
        #紅圈厚度
        pen.setWidth(2)
        painter.setPen(pen)
        painter.drawEllipse(self.circle_center, self.circle_radius, self.circle_radius)
        painter.end()

        self.image_label.setPixmap(circle_pixmap)

    def classify_photo(self, class_name):
        current_photo = self.photo_files.pop(0)
        target_folder = self.target_folders[class_name]

        # 檢查目標資料夾是否存在，如果不存在，則建立它
        # This check now happens right before moving a photo, ensuring the folder is created only if needed.
        if not target_folder.exists():
            target_folder.mkdir(parents=True, exist_ok=True)

        target_photo_path = target_folder / Path(current_photo).name

        if target_photo_path.exists():
            print(f"Skipped: {current_photo} (already exists in target folder)")
        else:
            shutil.move(current_photo, target_folder)

        if not self.photo_files:
            for row in range(self.button_grid_layout.rowCount()):
                for col in range(self.button_grid_layout.columnCount()):
                    item = self.button_grid_layout.itemAtPosition(row, col)
                    if item and item.widget():
                        item.widget().hide()
            
            self.image_label.hide()
            self.um_input_button.hide()
            self.open_original_button.hide()
            self.select_folder_button.setText("照片均已分類完成")
            self.button_grid_layout.setEnabled(False)
            self.select_folder_button.setEnabled(False)
            self.select_folder_button.show()
            return

        self.show_photo(self.photo_files[0])

        um_size = self.um_input_button.text()
        try:
            self.circle_radius = int(um_size) / 2
            self.show_photo(self.photo_files[0])
        except ValueError:
            pass

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

#    #新增遞迴方法，遍歷資料夾及其子資料夾並收集照片檔案(有疑慮)
#    def load_photos_in_folder(self, folder_path):
#        photo_files = []
#        for root, dirs, files in os.walk(folder_path):
#            for file in files:
#                if file.endswith(('.jpeg', '.jpg', '.png')):
#                    photo_files.append(os.path.join(root, file))
#        return photo_files
            
    #僅讀取指定資料夾內的照片，而不是遞迴遍歷子資料夾
    def load_photos_in_folder(self, folder_path):
        photo_files = []
        # 直接列出指定資料夾下的所有檔案，不遍歷子資料夾
        for file in os.listdir(folder_path):
            if file.endswith(('.jpeg', '.jpg', '.png')):
                photo_files.append(os.path.join(folder_path, file))
        return photo_files


if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 10)
    app.setFont(font)
    qtmodern.styles.dark(app)
    window = ImageClassifier()
    win = qtmodern.windows.ModernWindow(window)
    win.show()
    sys.exit(app.exec_())

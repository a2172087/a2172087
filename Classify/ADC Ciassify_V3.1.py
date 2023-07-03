import os
import shutil
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog
from PyQt5.QtGui import QPixmap, QFont, QIcon, QImage, QPainter
from PyQt5.QtCore import Qt, QRect
import qtmodern.styles
import qtmodern.windows
from pathlib import Path
from PyQt5 import QtGui


class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

        self.setWindowTitle('ADC Classify')
        app.setWindowIcon(QtGui.QIcon('abhl9.ico'))

    def initUI(self):
        # Build the main layout
        main_layout = QVBoxLayout(self)
        main_layout.setAlignment(Qt.AlignCenter)

        # Label for displaying the photo
        self.image_label = QLabel(self)
        self.image_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.image_label)

        # Create layout for folder selection button
        button_layout = QVBoxLayout()
        button_layout.addStretch()

        # Create folder selection button
        self.select_folder_button = QPushButton('選擇資料夾', self)
        self.select_folder_button.clicked.connect(self.select_folder)
        button_layout.addWidget(self.select_folder_button)

        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        # Folder paths
        self.source_folder = None
        self.target_folders = {}

        self.setWindowTitle("圖片分類程式")
        self.setGeometry(350, 100, 810, 610)

    def select_folder(self):
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        if dialog.exec_():
            folder_path = dialog.selectedFiles()[0]
            self.source_folder = folder_path

            # Get desktop path
            desktop_path = Path.home() / 'Desktop'

            # Folder names
            folder_names = {
                'ugly die': 'ugly die',
                'foreign material': 'foreign material',
                'Particle': 'Particle',
                'Probe mark shift': 'Probe mark shift',
                'Over kill': 'Over kill'
            }

            # Folder paths
            self.target_folders = {name: desktop_path / folder for name, folder in folder_names.items()}

            # Create target folders (if they don't exist)
            for folder in self.target_folders.values():
                folder.mkdir(parents=True, exist_ok=True)

            # Read the photos in the folder
            self.photo_files = []
            for file in os.listdir(self.source_folder):
                if file.endswith(('.jpeg', '.jpg', '.png')):
                    self.photo_files.append(os.path.join(self.source_folder, file))

            if self.photo_files:
                # Clear old buttons
                while self.layout().count() > 2:
                    item = self.layout().itemAt(2)
                    widget = item.widget()
                    self.layout().removeWidget(widget)
                    widget.setParent(None)

                # Create new buttons
                self.classification_buttons = []
                for class_name in self.target_folders.keys():
                    button = QPushButton('', self)
                    self.set_button_shortcut(button, class_name)  # Set shortcut
                    button.clicked.connect(lambda checked, name=class_name: self.classify_photo(name))
                    self.classification_buttons.append(button)
                    self.layout().addWidget(button)

                # Display the first photo
                self.show_photo(self.photo_files[0])

            # Hide the folder selection button
            self.select_folder_button.hide()

    def set_button_shortcut(self, button, class_name):
        if class_name == 'ugly die':
            button.setText('ugly die (快捷鍵: Q)')
            button.setShortcut('Q')
        elif class_name == 'foreign material':
            button.setText('foreign material (快捷鍵: W)')
            button.setShortcut('W')
        elif class_name == 'Particle':
            button.setText('Particle (快捷鍵: E)')
            button.setShortcut('E')
        elif class_name == 'Probe mark shift':
            button.setText('Probe mark shift (快捷鍵: R)')
            button.setShortcut('R')
        elif class_name == 'Over kill':
            button.setText('Over kill (快捷鍵: T)')
            button.setShortcut('T')

    def show_photo(self, photo_path):
        image = QImage(photo_path)
        image = image.scaled(500, 500, Qt.AspectRatioMode.KeepAspectRatio)

        # Crop the image to 500x500 from the center
        cropped_image = QImage(500, 500, QImage.Format_RGB888)
        painter = QPainter(cropped_image)
        target_rect = QRect(0, 0, 500, 500)
        source_rect = image.rect()
        source_rect.moveCenter(target_rect.center())
        painter.drawImage(target_rect, image, source_rect)
        painter.end()

        pixmap = QPixmap.fromImage(cropped_image)
        self.image_label.setPixmap(pixmap)

    def classify_photo(self, class_name):
        current_photo = self.photo_files.pop(0)
        target_folder = self.target_folders[class_name]
        shutil.move(current_photo, target_folder)

        if not self.photo_files:
            # All photos classified
            for button in self.classification_buttons:
                button.hide()
            self.image_label.hide()
            self.select_folder_button.setText("照片均已分類完成")
            self.select_folder_button.setEnabled(False)
            self.select_folder_button.show()
            return

        # Display the next photo
        self.show_photo(self.photo_files[0])


if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 10)
    app.setFont(font)
    qtmodern.styles.dark(app)

    window = ImageClassifier()
    win = qtmodern.windows.ModernWindow(window)
    win.show()

    sys.exit(app.exec_())
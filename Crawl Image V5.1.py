import sys
import os
import threading
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QFileDialog, QCheckBox, QMessageBox, QProgressBar
from PyQt5.QtGui import QFont, QIcon
import qtmodern.styles
import qtmodern.windows
import shutil
from PyQt5 import QtGui

class CrawlImageApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Crawl Image')
        app.setWindowIcon(QtGui.QIcon('icon_result.ico'))

        main_widget = QWidget()
        layout = QVBoxLayout()

        self.folder_label = QLabel('選擇抓取照片的資料夾:')
        self.folder_path = QLineEdit()
        self.folder_btn = QPushButton('選取資料夾')
        self.folder_btn.clicked.connect(self.openFolderDialog)

        self.auto_filter_checkbox = QCheckBox('自動篩選照片')
        self.auto_filter_checkbox.stateChanged.connect(self.autoFilterStateChanged)

        self.filter_folder_label = QLabel('篩選資料夾:')
        self.filter_folder_path = QLineEdit()
        self.filter_folder_btn = QPushButton('選取資料夾')
        self.filter_folder_btn.setEnabled(False)
        self.filter_folder_btn.clicked.connect(self.openFilterFolderDialog)

        self.max_images_label = QLabel('請輸入要抓取照片的最大數量:')
        self.max_images_input = QLineEdit()

        self.output_folder_label = QLabel('存放路徑:')
        self.output_folder_path = QLineEdit()
        self.output_folder_btn = QPushButton('選取資料夾')
        self.output_folder_btn.clicked.connect(self.openOutputFolderDialog)

        self.start_btn = QPushButton('啟動程式')
        self.start_btn.clicked.connect(self.startCrawling)

        self.progress_bar = QProgressBar()

        layout.addWidget(self.folder_label)
        layout.addWidget(self.folder_path)
        layout.addWidget(self.folder_btn)
        layout.addWidget(self.auto_filter_checkbox)
        layout.addWidget(self.filter_folder_label)
        layout.addWidget(self.filter_folder_path)
        layout.addWidget(self.filter_folder_btn)
        layout.addWidget(self.max_images_label)
        layout.addWidget(self.max_images_input)
        layout.addWidget(self.output_folder_label)
        layout.addWidget(self.output_folder_path)
        layout.addWidget(self.output_folder_btn)
        layout.addStretch(1)
        layout.addWidget(self.start_btn)
        layout.addWidget(self.progress_bar)
        layout.addStretch(1)

        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

        self.setGeometry(500, 300, 600, 470)

    def openFolderDialog(self):
        folder = QFileDialog.getExistingDirectory(self, '選取資料夾')
        self.folder_path.setText(folder)

    def autoFilterStateChanged(self, state):
        self.filter_folder_btn.setEnabled(state)

    def openFilterFolderDialog(self):
        folder = QFileDialog.getExistingDirectory(self, '選取篩選資料夾')
        self.filter_folder_path.setText(folder)

    def openOutputFolderDialog(self):
        folder = QFileDialog.getExistingDirectory(self, '選取存放路徑')
        self.output_folder_path.setText(folder)

    def startCrawling(self):
        folder_path = self.folder_path.text()
        filter_enabled = self.auto_filter_checkbox.isChecked()
        filter_folder_path = self.filter_folder_path.text()

        max_images_input = self.max_images_input.text()
        if not max_images_input.isdigit():
            QMessageBox.critical(self, '錯誤', '請正確輸入要抓取的照片最大數量')
            return

        max_images = int(max_images_input)
        output_folder = self.output_folder_path.text()

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        images_crawled = 0
        images_copied = 0

        def crawl_images():
            nonlocal images_copied, images_crawled

            total_files = sum(len(files) for _, _, files in os.walk(folder_path))
            current_progress = 0

            for root, _, files in os.walk(folder_path):
                for file in files:
                    if images_crawled >= max_images:
                        break

                    source_path = os.path.join(root, file)
                    dest_path = os.path.join(output_folder, file)

                    if filter_enabled and os.path.exists(os.path.join(filter_folder_path, file)):
                        continue

                    if filter_enabled:
                        found_in_filter = False
                        for filter_root, _, filter_files in os.walk(filter_folder_path):
                            if file in filter_files:
                                found_in_filter = True
                                break

                        if found_in_filter:
                            continue

                    shutil.copy(source_path, dest_path)
                    images_copied += 1
                    images_crawled += 1

                    current_progress = int(images_copied / max_images * 100)
                    self.progress_bar.setValue(current_progress)

            if images_copied < max_images:
                msg_box = QMessageBox()
                msg_box.setText(f'程式已執行完畢，共複製 {images_copied} 張照片，樣本數不足，請提高樣本數量!!!!')
                msg_box.exec_()
            else:
                msg_box = QMessageBox()
                msg_box.setText(f'程式已執行完畢，共複製 {images_copied} 張照片。')
                msg_box.exec_()

            self.start_btn.setEnabled(True)

        self.start_btn.setEnabled(False)

        worker_thread = threading.Thread(target=crawl_images)
        worker_thread.start()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 10)
    app.setFont(font)
    qtmodern.styles.dark(app)

    modern_window = qtmodern.windows.ModernWindow(CrawlImageApp())
    modern_window.show()

    sys.exit(app.exec_())

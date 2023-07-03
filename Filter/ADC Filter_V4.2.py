import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtGui import QFont, QIcon, QColor
from PyQt5.QtCore import Qt
import os
import shutil
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtGui


class FileMover(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(350, 250, 400, 400)
        self.setWindowTitle('Filter')
        app.setWindowIcon(QtGui.QIcon('aak51.ico'))

        # 創建選擇檔案按鈕
        self.select_button = QPushButton('選擇記事本檔案', self)
        self.select_button.setGeometry(100, 150, 200, 50)
        self.select_button.setFont(QFont("微軟正黑體", 8))
        self.select_button.clicked.connect(self.selectFile)

    def selectFile(self):
        file_dialog = QFileDialog(self)
        file_dialog.setFont(QFont("微軟正黑體", 8))  # 設置對話框字體
        file_dialog.setNameFilter('Text Files (*.txt)')  # 只顯示文字檔案
        file_dialog.setViewMode(QFileDialog.List)  # 以列表形式顯示檔案
        file_dialog.setFileMode(QFileDialog.ExistingFile)  # 只能選擇現有的檔案
        file_dialog.setOptions(QFileDialog.ReadOnly)  # 設定為唯讀模式
        file_dialog.setWindowTitle('選擇記事本檔案')

        if file_dialog.exec_():
            file_paths = file_dialog.selectedFiles()
            if len(file_paths) > 0:
                file_path = file_paths[0]
                self.processFile(file_path)

    def processFile(self, file_path):
        # 建立目標資料夾 'Need to check' 於桌面
        desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
        destination_folder = os.path.join(desktop_path, 'Need to check')
        if not os.path.exists(destination_folder):
            os.mkdir(destination_folder)

        try:
            # 讀取內容，指定編碼為 utf-8
            with open(file_path, 'r', encoding='utf-8') as file:
                paths = file.readlines()

            # 移除每行末尾的換行符號
            paths = [path.strip() for path in paths]

            # 處理每個路徑
            for path in paths:
                # 擷取檔案名稱
                file_name = os.path.basename(path)

                # 檢查檔案是否存在
                if os.path.exists(path):
                    # 移動檔案至目標資料夾
                    destination_path = os.path.join(destination_folder, file_name)
                    shutil.move(path, destination_path)
                    print(f"已移動檔案: {file_name}")
                else:
                    print(f"找不到檔案: {file_name}")

            # 刪除原始的 TXT 檔案
            os.remove(file_path)

            # 顯示成功訊息視窗
            QMessageBox.information(self, '成功', '程式已成功執行完成', QMessageBox.Ok, QMessageBox.Ok)
        except Exception as e:
            # 顯示錯誤訊息視窗
            QMessageBox.critical(self, '錯誤', f'程式執行時發生錯誤：\n{str(e)}', QMessageBox.Ok, QMessageBox.Ok)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # 套用 qtmodern 的 dark 樣式
    font = QFont("微軟正黑體", 8)
    app.setFont(font)
    qtmodern.styles.dark(app)
    window = FileMover()
    window = qtmodern.windows.ModernWindow(window)
    window.show()
    sys.exit(app.exec_())



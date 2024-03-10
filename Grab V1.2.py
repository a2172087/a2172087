import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QFileDialog, QLabel, QProgressBar, QMessageBox
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import QThread, pyqtSignal
import qtmodern.styles
import qtmodern.windows
import os
import shutil

class Worker(QThread):
    updateProgress = pyqtSignal(int)
    finished = pyqtSignal()
    errorOccurred = pyqtSignal(str)  # New signal for error reporting

    def __init__(self, chosenPath, folderNames, destination):
        super().__init__()
        self.chosenPath = chosenPath
        self.folderNames = folderNames
        self.destination = destination

    def run(self):
        try:
            totalFolders = len(self.folderNames)
            count = 0

            for root, dirs, files in os.walk(self.chosenPath):
                for dirName in dirs:
                    if dirName in self.folderNames:
                        srcPath = os.path.join(root, dirName)
                        dstPath = os.path.join(self.destination, dirName)
                        shutil.copytree(srcPath, dstPath, dirs_exist_ok=True)
                        print(f"已複製: {dirName}")
                        count += 1
                        self.updateProgress.emit(int((count / totalFolders) * 100))

            self.finished.emit()
        except Exception as e:
            self.errorOccurred.emit(str(e))

class ImageClassifier(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Grab')
        self.setWindowIcon(QIcon('icon_resultt.ico'))
        layout = QVBoxLayout()

        # 將視窗移動到特定位置
        self.move(600, 800)

        self.pathLabel = QLabel("未選擇路徑")
        layout.addWidget(self.pathLabel)

        self.btnChoosePath = QPushButton('選擇搜尋路徑')
        self.btnChoosePath.clicked.connect(self.choosePath)
        layout.addWidget(self.btnChoosePath)

        self.txtFolderNames = QTextEdit()
        layout.addWidget(self.txtFolderNames)

        self.btnExecute = QPushButton('執行')
        self.btnExecute.clicked.connect(self.executeSearch)
        layout.addWidget(self.btnExecute)

        self.progressBar = QProgressBar(self)
        layout.addWidget(self.progressBar)

        self.setLayout(layout)
        self.chosenPath = ''

    def choosePath(self):
        self.chosenPath = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        self.pathLabel.setText(f"選擇的路徑: {self.chosenPath}")

    def executeSearch(self):
        if not self.chosenPath:
            QMessageBox.warning(self, "Error", "請先選擇一個搜尋路徑")
            return

        folderNames = self.txtFolderNames.toPlainText().split('\n')
        destination = os.path.join(os.path.expanduser('~'), 'Desktop', '符合搜尋要求的資料夾')

        if not os.path.exists(destination):
            os.makedirs(destination)

        self.worker = Worker(self.chosenPath, folderNames, destination)
        self.worker.updateProgress.connect(self.progressBar.setValue)
        self.worker.finished.connect(self.onFinished)
        self.worker.errorOccurred.connect(self.onError)  # Connect the error signal
        self.worker.start()

    def onFinished(self):
        self.progressBar.setValue(100)
        QMessageBox.information(self, "Operation Complete", "搜尋和複製操作完成！")
        print("搜尋和複製操作完成！")

    def onError(self, message):
        QMessageBox.warning(self, "Operation Failed", f"搜尋和複製操作失敗，請再試一次。\nError: {message}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 9)
    app.setFont(font)
    qtmodern.styles.dark(app)

    window = ImageClassifier()
    win = qtmodern.windows.ModernWindow(window)
    win.show()
    sys.exit(app.exec_())

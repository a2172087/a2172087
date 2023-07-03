import sys
import traceback
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QMessageBox, QFileDialog
from PyQt5.QtGui import QFont, QIcon, QColor
import qtmodern.styles
from qtmodern.windows import ModernWindow
import pandas as pd
from openpyxl import load_workbook
from PyQt5 import QtCore
from PyQt5 import QtGui

class Application(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Age Hold lot status')
        app.setWindowIcon(QtGui.QIcon('apw13.ico'))

        layout = QVBoxLayout()
        start_button = QPushButton('開始執行')
        start_button.setStyleSheet('font-size: 13px;')
        start_button.clicked.connect(self.execute_program)
        #start_button.setIcon(QIcon('ayh9m.ico'))

        layout.addWidget(start_button)
        self.setLayout(layout)

    def execute_program(self):
        file_dialog = QFileDialog()
        file_dialog.setDirectory("M:\QAReport\zz-Q000\Age Hold lot status")
        file_path, _ = file_dialog.getOpenFileName(self, '選擇文件', '', 'Excel Files (*.xlsx)')

        if file_path:
            try:
                xl = pd.ExcelFile(file_path)
                sheet_names = xl.sheet_names
                d_l_dict = {}

                for sheet_name in sheet_names:
                    df = xl.parse(sheet_name, header=0)
                    print(df.columns)
                    
                    if df.columns[3] != 'ADT Lot ID':
                        continue

                    if 'QE/PE Remark On yesterday' in df.columns:
                        l_column = 'QE/PE Remark On yesterday'
                    elif 'QE/PE Remark' in df.columns:
                        l_column = 'QE/PE Remark'
                    else:
                        continue

                    d_values = df['ADT Lot ID'].tolist()
                    l_values = df[l_column].tolist()

                    for d, l in zip(d_values, l_values):
                        if d != 'ADT Lot ID' and pd.notnull(l):
                            d_l_dict[d] = l

                book = load_workbook(file_path)

                for sheet_name in sheet_names:
                    sheet = book[sheet_name]

                    if sheet['D1'].value != 'ADT Lot ID':
                        continue

                    if sheet['L1'].value != 'QE/PE Remark On yesterday' and sheet['L1'].value != 'QE/PE Remark':
                        continue

                    for d_cell, l_cell in zip(sheet['D'], sheet['L']):
                        d = d_cell.value
                        l = l_cell.value
                        if d in d_l_dict and (l is None or l == ''):
                            l_cell.value = d_l_dict[d]

                book.save(file_path)
                book.close()  # 關閉Excel文件
                QMessageBox.information(self, '', '程序執行完成！')
            except Exception as e:
                error_message = traceback.format_exc()
                QMessageBox.critical(self, '錯誤', f'程序發生異常(常見狀況例如:有人正在使用工作表、檔案目前是開啟狀態)：\n\n{error_message}')
                sys.exit(1)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 8)
    app.setFont(font)
    qtmodern.styles.dark(app)
    window = ModernWindow(Application())
    window.setWindowOpacity(1)  # 更改透明度0.1~1
    window.setAttribute(QtCore.Qt.WA_TranslucentBackground)  # Enable translucent background
    window.setGeometry(350, 150, 400, 400)  # Set the geometry of the window
    window.show()
    sys.exit(app.exec_())

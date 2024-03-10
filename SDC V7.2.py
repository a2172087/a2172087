import sys
import os
from os.path import expanduser
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QTextEdit, QListWidget, QLineEdit, QComboBox, QMessageBox  # 加入 QMessageBox
from PyQt5.QtGui import QFont
import qtmodern.styles
import qtmodern.windows
from PyQt5 import QtGui
import subprocess
from PyQt5.QtCore import QThread, pyqtSignal

app = QApplication(sys.argv)

# 在顶部添加新的全局变量 proc_option
proc_option = 'proc.ntc'  # 默认值为 'proc.ntc'

class Worker(QThread):
    finished = pyqtSignal(bool)

    def run(self):
        try:
            # 在这里添加您要执行的代码
            execution_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Execution.py")
            if os.path.exists(execution_script_path):
                subprocess.run(["python", execution_script_path], check=True)
            self.finished.emit(True)
        except subprocess.CalledProcessError as e:
            print("執行 Execution.py 時出錯:", e)
            self.finished.emit(False)

class DataProcessingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.setWindowTitle('SDC')
        app.setWindowIcon(QtGui.QIcon('pie-chartt.ico'))

        # 初始化按鈕狀態
        self.input_config_button.setEnabled(False)
        self.run_button.setEnabled(False)

    def init_ui(self):
        self.setWindowTitle("SDC")
        self.setGeometry(100, 100, 500, 50)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.label = QLabel("資料夾路徑:")
        self.layout.addWidget(self.label)

        self.folder_list_widget = QListWidget()
        self.folder_list_widget.setFixedHeight(400)
        self.layout.addWidget(self.folder_list_widget)

        self.browse_button = QPushButton("請選擇資料夾路徑")
        self.browse_button.clicked.connect(self.browse_folder)
        self.layout.addWidget(self.browse_button)

        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        self.layout.addWidget(self.file_path_edit)
        self.file_path_edit.hide() #先隱藏

        # 添加下拉式選單供用户選擇 proc_option 的值
        self.proc_option_label = QLabel("選擇 proc 選項:")
        self.layout.addWidget(self.proc_option_label)

        self.proc_option_combobox = QComboBox()
        self.proc_option_combobox.addItem("請選擇 ntc or pass")
        self.proc_option_combobox.addItem("proc.ntc")
        self.proc_option_combobox.addItem("proc.pass")
        self.proc_option_combobox.currentIndexChanged.connect(self.update_proc_option)
        self.layout.addWidget(self.proc_option_combobox)

        self.central_widget.setLayout(self.layout)

        self.execute_button = QPushButton("請選擇訓練時的config.json模型檔案")
        self.execute_button.clicked.connect(self.execute_processing)
        self.layout.addWidget(self.execute_button)
        self.input_config_button = self.execute_button  # 將按鈕另存為 input_config_button

        self.run_button = QPushButton("執行")
        self.run_button.clicked.connect(self.run_code)
        self.layout.addWidget(self.run_button)

        self.log_label = QLabel("執行日誌:")
        self.log_label.hide()
        self.layout.addWidget(self.log_label)

        self.log_text_edit = QTextEdit()
        self.log_text_edit.setFixedHeight(70)
        self.log_text_edit.hide() #先隱藏

        # 新增欄位以顯示描述
        self.defect_code_description_list = QListWidget()
        self.defect_code_description_list.setFixedHeight(20)  # 這裡設定您想要的高度
        self.layout.addWidget(self.defect_code_description_list)
        self.defect_code_description_list.hide() #先隱藏

        # 初始化排序後的鍵列表和Defect code values
        self.sorted_defect_codes = []
        self.defect_code_values = {}

        self.defect_code_selection_list = QListWidget()
        self.defect_code_selection_list.itemSelectionChanged.connect(self.update_selected_defect_codes)
        self.defect_code_selection_list.setVisible(True)  # 隱藏欄位
        self.layout.addWidget(self.defect_code_selection_list)
        self.defect_code_selection_list.hide() #先隱藏

    def browse_folder(self):
        folder_dialog = QFileDialog.getExistingDirectory(self, "選擇資料夾")
        if folder_dialog:
            self.folder_list_widget.addItem(folder_dialog)

    #我從"Input Defect code list"的路徑中讀取檔案內容
    def input_defect_codes(self):
        file_dialog = QFileDialog.getOpenFileName(self, "選擇 Defect code 檔案")
        if file_dialog[0]:
            file_path = file_dialog[0]
            self.file_path_edit.setText(file_path)
            defect_code_descriptions = self.read_defect_code_descriptions(file_path)

            if defect_code_descriptions:
                self.sorted_defect_codes = sorted(defect_code_descriptions.keys())
                log_message = "Defect code 描述：\n"
                for defect_code in self.sorted_defect_codes:
                    description = defect_code_descriptions[defect_code]
                    log_message += f"{defect_code}: {description}\n"

    def process_subfolders(self, root_folder):
        selected_lines = []

        for root, _, files in os.walk(root_folder):
            for file_name in files:
                if file_name.endswith('.txt') and proc_option in file_name:  # 使用 proc_option 变量
                    input_path = os.path.join(root, file_name)
                    with open(input_path, 'r') as input_file:
                        lines = input_file.readlines()

                    for line in lines:
                        data = line.strip().split(', ')
                        ntd_check_code = data[2]
                        normal_dont_check_code = data[1][-2:]  # 取normal_dont_check_code的後兩位字元

                        # 檢查ntd_check_code和normal_dont_check_code的後兩位字元是否相同，如果相同則跳過該行數據
                        #if ntd_check_code[-2:] == normal_dont_check_code[-2:]:
                            #continue

                        # 新增條件：如果ntd_check_code只有一個字元，則跳過該行
                        if len(ntd_check_code) == 1:
                            continue

                        selected_lines.append(line)

        return selected_lines

    def read_data(self, file_path, defect_code_descriptions):
        defect_code_values = {}

        try:
            with open(file_path, 'r') as file:
                lines = file.readlines()

            for line in lines:
                data = line.strip().split(', ')
                confidence_data = data[4].split(':')[-1].split()

                for i, value in enumerate(confidence_data):
                    defect_code = self.sorted_defect_codes[i]
                    if defect_code not in defect_code_values:
                        defect_code_values[defect_code] = []
                    defect_code_values[defect_code].append(float(value))

            return defect_code_values

        except Exception as e:
            print("發生錯誤:", e)
            return None

    def display_defect_code_values(self, defect_code_values):
            if defect_code_values:
                log_message = "Defect code 數值:\n"
                for defect_code in self.sorted_defect_codes: 
                    description = self.defect_code_descriptions.get(defect_code, defect_code)
                    values = defect_code_values.get(defect_code, [])
                    log_message += f"{description}: {values}\n"
            else:
                log_message = "無法讀取 Defect code 數值。\n"
            self.log_text_edit.setPlainText(log_message)

    def update_selected_defect_codes(self):
        selected_items = self.defect_code_selection_list.selectedItems()

    def update_proc_option(self, index):
        # 当用户选择不同的 proc_option 时更新全局变量
        selected_option = self.proc_option_combobox.currentText()
        global proc_option
        proc_option = selected_option

        # 在這裡呼叫 save_proc_option_to_file 函數，將選擇的 proc_option 儲存到桌面的檔案
        self.save_proc_option_to_files()

        # 更新按钮状态
        self.input_config_button.setEnabled(True)

    def save_proc_option_to_files(self):
        # 取得使用者的桌面路徑
        desktop_path = os.path.expanduser("~/Desktop")

        # 組合檔案的完整路徑
        file_path = os.path.join(desktop_path, "proc_option.txt")

        try:
            # 打開檔案以寫入模式，如果不存在則自動創建
            with open(file_path, 'w+') as file:
                # 將選擇的 proc_option 寫入檔案
                file.write(proc_option)
                
        except Exception as e:
            # 處理檔案寫入時的錯誤
            print(f"儲存檔案時發生錯誤：{str(e)}")

    def execute_processing(self):
        selected_folders = [self.folder_list_widget.item(i).text() for i in range(self.folder_list_widget.count())]

        log_message = ""
        selected_lines_list = []
        for folder_path in selected_folders:
            selected_lines = self.process_subfolders(folder_path)
            if selected_lines:
                selected_lines_list.extend(selected_lines)

        if selected_lines_list:
            output_file_name = 'standard_deviation_for_confidence.txt'
            desktop_path = os.path.join(expanduser("~"), "Desktop")
            output_path = os.path.join(desktop_path, output_file_name)

            with open(output_path, 'w') as output_file:
                output_file.writelines(selected_lines_list)
            log_message = f"已將執行日誌保存到 {output_path} 中。\n"
        else:
            log_message = "所有資料夾處理完成，未找到符合條件的數據。\n"
        self.log_text_edit.setPlainText(log_message)

        file_path = os.path.join(desktop_path, output_file_name)
        self.defect_code_descriptions = self.read_defect_code_descriptions(self.file_path_edit.text())
        self.defect_code_values = self.read_data(file_path, self.defect_code_descriptions)
        self.display_defect_code_values(self.defect_code_values)

        self.defect_code_selection_list.clear()
        self.defect_code_description_list.clear()
        self.defect_code_selection_list.addItems(self.sorted_defect_codes)
        
        #執行 Input config.exe
        executable_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Input config.exe")
        if os.path.exists(executable_path):
            try:
                subprocess.run([executable_path], check=True, cwd=os.path.dirname(os.path.abspath(__file__)))
            except subprocess.CalledProcessError as e:
                print("執行 Input config.exe 時出錯:", e)

#        # 執行 Input config.py
#        disassemble_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Input config.py")
#        if os.path.exists(disassemble_script_path):
#            try:
#                subprocess.run(["python", disassemble_script_path], check=True)
#            except subprocess.CalledProcessError as e:
#                print("執行 Input config.py 時出錯:", e)

        # 啟用 "執行" 按鈕
        self.run_button.setEnabled(True)

    def run_code(self):
        #執行與.py相同資料夾內的Execution.exe
        executable_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Execution.exe")
        if os.path.exists(executable_path):
            try:
                subprocess.run([executable_path], check=True)
            except subprocess.CalledProcessError as e:
                print("執行 Execution.exe 時出錯:", e)

#    def run_code(self):
        # 在这里添加要执行的代码
        # 執行與.py相同資料夾內的Execution.py
#        execution_script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Execution.py")
#        if os.path.exists(execution_script_path):
#            try:
#                subprocess.run(["python", execution_script_path], check=True)
#            except subprocess.CalledProcessError as e:
#                print("執行 Execution.py 時出錯:", e)

    def read_defect_code_descriptions(self, file_path):
        defect_code_descriptions = {}

        try:
            with open(file_path, 'r') as file:
                data = file.read()
                import re
                descriptions = re.findall(r'"description":\s*"([^"]+)"', data)
                for i, description in enumerate(descriptions, start=1):
                    defect_code = f'Defect code {i:02}'
                    defect_code_descriptions[defect_code] = description

            return defect_code_descriptions

        except Exception as e:
            print("發生錯誤:", e)
            return None
        
    def run_code(self):
        self.worker = Worker()
        self.worker.finished.connect(self.on_execution_finished)
        self.worker.start()

    def on_execution_finished(self, success):
        if success:
            QMessageBox.information(self, "執行結果", "程式已執行完畢")
        else:
            QMessageBox.critical(self, "執行錯誤", "程式執行失敗，請重新再試一次")

def main():
    app = QApplication(sys.argv)
    font = QFont("微軟正黑體", 9)
    app.setFont(font)
    qtmodern.styles.dark(app)

    window = DataProcessingApp()

    # 設定視窗顯示的 X 和 Y 座標位置
    x_coordinate = 250  # 設定 X 座標的數值
    y_coordinate = 200  # 設定 Y 座標的數值

    # 調整視窗的位置
    window.move(x_coordinate, y_coordinate)

    window = qtmodern.windows.ModernWindow(window)
    window.show()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
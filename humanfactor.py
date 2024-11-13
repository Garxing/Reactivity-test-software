import sys
import random
import time
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QTextEdit, QGridLayout,QMessageBox
)
from PyQt5.QtGui import QIcon
from datetime import datetime
import pandas as pd
from PyQt5.QtWidgets import QInputDialog

class GameWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()
        self.start_time = None
        self.correct_answers = 0
        self.total_answers = 0
        self.questions_answered = 0  # 新增属性，用于追踪已回答的问题数量
        self.xlsx_file = 'results.xlsx'  # xlsx 文件名
        self.calculate = 15 # 设置统一答题数量
        self.create_xlsx_file()  # 创建 xlsx 文件和表头
        self.load_results()  # 启动时加载最新结果

    def init_ui(self):
        self.setWindowTitle("汉字选择游戏")
        self.setGeometry(100, 100, 600, 600)
        self.setWindowIcon(QIcon('image.png'))

        layout = QVBoxLayout()
        self.setLayout(layout)

        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(["12", "18", "24", "30"])

        self.font_style_combo = QComboBox()
        self.font_style_combo.addItems(["宋体", "楷体", "黑体", "隶书"])

        self.color_scheme_combo = QComboBox()
        self.color_scheme_combo.addItems(["蓝字白色背景","红字黑色背景", "黄字红色背景"])

        self.option_size_combo = QComboBox()
        self.option_size_combo.addItems(["4*4", "6*6", "8*8"])

        self.font_size_combo.currentIndexChanged.connect(self.update_buttons)
        self.font_style_combo.currentIndexChanged.connect(self.update_buttons)
        self.color_scheme_combo.currentIndexChanged.connect(self.update_buttons)
        self.option_size_combo.currentIndexChanged.connect(self.update_buttons)

        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("字体大小:"))
        control_layout.addWidget(self.font_size_combo)
        control_layout.addWidget(QLabel("字体样式:"))
        control_layout.addWidget(self.font_style_combo)
        control_layout.addWidget(QLabel("颜色方案:"))
        control_layout.addWidget(self.color_scheme_combo)
        control_layout.addWidget(QLabel("选项数量:"))
        control_layout.addWidget(self.option_size_combo)

        layout.addLayout(control_layout)

        self.prompt_label = QLabel("请选择出汉字矩阵中唯一不同的汉字")
        layout.addWidget(self.prompt_label)

        self.grid_layout = QGridLayout()
        self.grid_layout.setSpacing(5)  # 设置网格布局的间距
        self.buttons = []
        self.create_game_buttons(4)
        layout.addLayout(self.grid_layout)

        # 操作按钮
        self.start_button = QPushButton("Start")
        self.restart_button = QPushButton("Restart")
        self.close_button = QPushButton("Close")

        button_style = (
            "QPushButton {"
            "background-color: rgb(86, 78, 215); "
            "color: white; "
            "border-style: solid; "
            "border-width: 1px; "
            "border-radius: 10px; "
            "padding: 10px; "
            "margin: 10px; "
            "min-height: 15px;"
            "}"
        )
        
        self.start_button.setStyleSheet(button_style)
        self.restart_button.setStyleSheet(button_style)
        self.close_button.setStyleSheet(button_style)

        self.start_button.clicked.connect(self.start_game)
        self.restart_button.clicked.connect(self.restart_game)
        self.close_button.clicked.connect(self.close)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.restart_button)
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        layout.addWidget(self.result_text)

    def create_game_buttons(self, size):
        for button in self.buttons:
            self.grid_layout.removeWidget(button)
            button.deleteLater()
        self.buttons.clear()

        font_size = int(self.font_size_combo.currentText())
        font_style = self.font_style_combo.currentText()
        color_scheme = self.color_scheme_combo.currentText()

        for i in range(size):
            for j in range(size):
                button = QPushButton("")
                button.setFixedSize(50, 50)
                button.clicked.connect(self.button_clicked)

                button.setStyleSheet(self.get_button_style(font_size, font_style, color_scheme))
                
                self.grid_layout.addWidget(button, i, j)
                self.buttons.append(button)

    def get_button_style(self, font_size, font_style, color_scheme):
        style_map = {
            "红字黑色背景": f"font-size: {font_size}px; font-family: {font_style}; color: red; background-color: black;",
            "黄字红色背景": f"font-size: {font_size}px; font-family: {font_style}; color: yellow; background-color: red;",
            "蓝字白色背景": f"font-size: {font_size}px; font-family: {font_style}; color: blue; background-color: white;"
        }
        return style_map.get(color_scheme, f"font-size: {font_size}px; font-family: {font_style};")

    def update_buttons(self):
        selected_size = self.option_size_combo.currentText().split('*')
        size = int(selected_size[0])
        self.create_game_buttons(size)

    # def start_game(self):
    #     self.start_time = time.time()
    #     self.correct_answers = 0
    #     self.total_answers = 0
    #     self.result_text.clear()
    #     self.show_buttons()
    # def start_game(self):
    #         self.start_time = time.time()
    #         self.correct_answers = 0
    #         self.total_answers = 0
    #         self.questions_answered = 0  # 重置已回答的问题计数
    #         self.result_text.clear()
    #         self.show_buttons()
    def start_game(self):
        # 弹出对话框获取被调查者的姓名
        text, ok = QInputDialog.getText(self, '输入姓名', '请输入您的姓名:')
        if ok and text:
            self.investigator_name = text  # 存储被调查者的姓名
        else:
            self.investigator_name = "未知"  # 如果用户取消或未输入姓名，默认为“未知”

        self.start_time = time.time()
        self.correct_answers = 0
        self.total_answers = 0
        self.questions_answered = 0  # 重置已回答的问题计数
        self.result_text.clear()
        self.show_buttons()  # 只有在点击 Start 按钮后才显示按钮上的汉字

    def restart_game(self):
        self.correct_answers = 0
        self.total_answers = 0
        self.result_text.clear()
        self.prompt_label.setText("请选择出，仅出现过一次的汉字")
        self.show_buttons()

    def show_buttons(self):
        for button in self.buttons:
            button.setText("")  # 先清空所有按钮的文本
        num_buttons = len(self.buttons)
        options = ["上"] * (num_buttons - 1) + ["下"]
        random.shuffle(options)

        for button, text in zip(self.buttons, options):
            button.setText(text)  # 再填充新的汉字
    # def button_clicked(self):
    #     self.total_answers += 1
    #     self.questions_answered += 1  # 每次点击按钮时增加已回答的问题计数
    #     if self.sender().text() == "下":
    #         self.correct_answers += 1
    #     self.show_buttons()

    #     if self.total_answers >= 5:
    #         elapsed_time = time.time() - self.start_time
    #         accuracy = (self.correct_answers / self.total_answers) * 100
    #         self.result_text.setText(f"花费时间: {elapsed_time:.2f} 秒\n准确度: {accuracy:.2f}%")
    #         self.save_result(elapsed_time, accuracy)  # 保存结果
    def button_clicked(self):
        self.total_answers += 1
        self.questions_answered += 1  # 每次点击按钮时增加已回答的问题计数

        if self.sender().text() == "下":
            self.correct_answers += 1
        
        if self.questions_answered >= self.calculate:
            # 当回答了5个问题后，隐藏按钮文本
            for button in self.buttons:
                button.setText("")
            elapsed_time = time.time() - self.start_time
            accuracy = (self.correct_answers / self.total_answers) * 100
            self.result_text.setText(f"花费时间: {elapsed_time:.2f} 秒\n准确度: {accuracy:.2f}%")
            self.save_result(elapsed_time, accuracy)
        else:
            # 如果还没有达到5个问题，则继续显示汉字
            self.show_buttons()

    # def create_xlsx_file(self):
    #     # 检查并创建 Excel 文件，写入表头
    #     if not os.path.isfile(self.xlsx_file):
    #         df = pd.DataFrame(columns=["日期", "答题时间", "准确率", "字体大小", "字体样式", "字体颜色", "选项数量"])
    #         df.to_excel(self.xlsx_file, index=False)  # 创建一个新的 Excel 文件
    def create_xlsx_file(self):
        # 检查并创建 Excel 文件，写入表头
        if not os.path.isfile(self.xlsx_file):
            df = pd.DataFrame(columns=["姓名", "日期", "答题时间", "准确率", "字体大小", "字体样式", "字体颜色", "选项数量"])
            df.to_excel(self.xlsx_file, index=False)  # 创建一个新的 Excel 文件

    def save_result(self, elapsed_time, accuracy):
        # 获取当前日期和时间
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # 获取下拉菜单的当前值
        font_size = self.font_size_combo.currentText()
        font_style = self.font_style_combo.currentText()
        font_color = self.color_scheme_combo.currentText()
        option_size = self.option_size_combo.currentText()

        # 将结果添加到 DataFrame
        data = {
            "姓名": [self.investigator_name],
            "日期": [current_time],
            "答题时间": [f"{elapsed_time:.2f} 秒"],
            "准确率": [f"{accuracy:.2f}%"],
            "字体大小": [font_size],
            "字体样式": [font_style],
            "字体颜色": [font_color],
            "选项数量": [option_size],
        }

        new_row = pd.DataFrame(data)

        try:
            # 尝试读取原来的 Excel 文件，如果存在
            existing_df = pd.read_excel(self.xlsx_file)  # 读取现有 Excel 文件
            df = pd.concat([existing_df, new_row], ignore_index=True)
        except FileNotFoundError:
            df = new_row  # 如果文件不存在，则使用新行创建 DataFrame
        except PermissionError:
            # 如果文件被占用，显示错误消息
            QMessageBox.critical(self, "文件访问错误", "文件 results.xlsx 正在被其他程序使用，请关闭文件后再试。")
            return

        # 写入 Excel 文件，格式化列宽
        try:
            with pd.ExcelWriter(self.xlsx_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                worksheet = writer.book.active  # 获取当前活动的工作表
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    adjusted_width = (max_length + 5)  # 增加宽度设置为 max_length + 5
                    worksheet.column_dimensions[column_letter].width = adjusted_width
        except PermissionError:
            # 如果文件被占用，显示错误消息
            QMessageBox.critical(self, "文件访问错误", "文件 results.xlsx 正在被其他程序使用，请关闭文件后再试。")

    def load_results(self):
        # 从 xlsx 文件中加载最新结果
        try:
            existing_df = pd.read_excel(self.xlsx_file)  # 使用 pandas 读取 Excel 文件
            if not existing_df.empty:  # 检查是否有数据
                latest_result = existing_df.iloc[-1]  # 获取最新的一行数据
                
                # 转换所有元素为字符串并格式化输出
                latest_result_str = list(map(str, latest_result.values))  # 转换为字符串
                
                self.result_text.setPlainText(', '.join(latest_result_str[:3]) + '，' + 
                                                '，'.join(latest_result_str[3:]))  # 格式化输出
            else:
                self.result_text.setPlainText("没有记录结果。\n")
        except FileNotFoundError:
            self.result_text.setPlainText("没有记录结果。\n")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    game_window = GameWindow()
    game_window.show()
    sys.exit(app.exec_())

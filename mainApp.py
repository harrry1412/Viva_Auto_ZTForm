import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QComboBox, QRadioButton, QDateEdit, QMessageBox, QButtonGroup
)
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QFont, QIcon
from openpyxl import Workbook
from dataProcessor import DataProcessor
import requests
import os
import json
import re

# 全局常量
CONFIG_FILENAME = "config.json"
ICON_FILENAME = "app_icon.png"
APP_TITLE = "VIVA自提单自动生成工具 V1.0.0"


class DataExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()  # 加载配置文件
        self.processor = DataProcessor()  # 实例化数据处理类
        self.session = None  # 全局 requests.Session 对象，用于复用 cookie
        self.init_ui()

    def load_config(self):
        """加载配置文件"""
        base_path = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(base_path, CONFIG_FILENAME)
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        else:
            raise FileNotFoundError(f"配置文件未找到: {config_path}")

    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QIcon(self.get_icon_path()))

        screen_size = QApplication.primaryScreen().size()
        self.resize(screen_size.width() // 3, screen_size.height() * 3 // 4)

        layout = QVBoxLayout()

        font = QFont("Arial", 14)
        self.setFont(font)

        # 登录按钮
        self.login_button = QPushButton("登录")
        self.login_button.clicked.connect(self.on_login_click)
        layout.addWidget(self.login_button)

        # 登录状态文本（初始隐藏）
        self.login_status_label = QLabel("登录成功！")
        self.login_status_label.setVisible(False)
        layout.addWidget(self.login_status_label)

        # 模式选择
        self.mode_group = QButtonGroup(self)
        self.date_mode_button = QRadioButton("按日期生成")
        self.number_mode_button = QRadioButton("按单号生成")
        self.mode_group.addButton(self.date_mode_button)
        self.mode_group.addButton(self.number_mode_button)
        self.date_mode_button.setChecked(True)

        layout.addWidget(QLabel("选择生成模式:"))
        layout.addWidget(self.date_mode_button)
        layout.addWidget(self.number_mode_button)

        self.date_mode_button.toggled.connect(self.update_input_fields)

        # 日期输入
        self.target_date_input = QDateEdit()
        self.target_date_input.setCalendarPopup(True)
        self.target_date_input.setDate(QDate.currentDate())

        # 单号输入
        self.target_number_input = QLineEdit(self.fetch_default_order_number())
        self.target_number_input.setVisible(False)

        layout.addWidget(QLabel("目标日期或单号:"))
        layout.addWidget(self.target_date_input)
        layout.addWidget(self.target_number_input)

        # 其他输入框
        self.output_filename_input = QLineEdit("Viva自提单生成")
        layout.addWidget(QLabel("输出文件名:"))
        layout.addWidget(self.output_filename_input)

        self.login_url_input = QLineEdit(self.config.get("login_url", ""))
        layout.addWidget(QLabel("登录页面 URL:"))
        layout.addWidget(self.login_url_input)

        self.url1_input = QLineEdit(self.config.get("url1", ""))
        layout.addWidget(QLabel("数据 URL:"))
        layout.addWidget(self.url1_input)

        # 下拉框选项
        self.include_stock_status_input = QComboBox()
        self.include_stock_status_input.addItems(["否", "是"])
        layout.addWidget(QLabel("是否生成订货列:"))
        layout.addWidget(self.include_stock_status_input)

        self.finished_filter_input = QComboBox()
        self.finished_filter_input.addItems(["全部", "仅已标记完结的订单", "仅未标记完结的订单"])
        layout.addWidget(QLabel("生成已标记完结还是未完结的订单:"))
        layout.addWidget(self.finished_filter_input)

        self.skip_negative_qty_input = QComboBox()
        self.skip_negative_qty_input.addItems(["是", "否"])
        layout.addWidget(QLabel("跳过负库存记录:"))
        layout.addWidget(self.skip_negative_qty_input)

        # 生成按钮
        self.generate_button = QPushButton("生成")
        self.generate_button.clicked.connect(self.on_generate_click)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)

    def update_input_fields(self):
        """根据选择的生成模式切换输入框"""
        if self.date_mode_button.isChecked():
            self.target_date_input.setVisible(True)
            self.target_number_input.setVisible(False)
        else:
            self.target_date_input.setVisible(False)
            self.target_number_input.setVisible(True)

    def on_login_click(self):
        """点击登录按钮的处理逻辑"""
        try:
            login_url = self.login_url_input.text()
            if not login_url:
                QMessageBox.warning(self, "警告", "登录页面 URL 不能为空！")
                return

            # 设置登录按钮状态为“登录中，请稍后”
            self.login_button.setText("登录中，请稍后")
            self.login_button.setEnabled(False)  # 禁用按钮以防重复点击
            QApplication.processEvents()  # 刷新界面

            # 执行登录操作
            self.session = self.processor.get_authenticated_session(login_url)

            # 登录成功后尝试加载默认单号
            default_order_number = self.fetch_default_order_number()
            if default_order_number == "解析错误" or not default_order_number:
                raise ValueError("默认单号解析错误，登录失败。")

            # 默认单号解析成功
            self.target_number_input.setText(default_order_number)
            self.login_button.setText("登录成功")
            self.login_button.setEnabled(False)  # 禁用按钮，防止再次点击
            self.login_status_label.setVisible(True)
            QMessageBox.information(self, "提示", f"登录成功！默认单号: {default_order_number}")

        except Exception as e:
            # 登录失败时恢复按钮状态
            self.login_button.setText("登录")
            self.login_button.setEnabled(True)
            QMessageBox.critical(self, "错误", f"登录失败: {str(e)}")

    def fetch_default_order_number(self):
        """从 URL1 的 datalist 提取第一个字典的 Number 值"""
        url1 = self.config.get("url1", "")
        if not url1:
            return "URL错误"

        try:
            response = self.session.get(url1)  # 使用已登录的 session
            response.raise_for_status()

            # 尝试解析 datalist 数据
            match = re.search(r"var\s+datalist\s*=\s*(\[.*?\]);", response.text, re.DOTALL)
            if match:
                datalist = json.loads(match.group(1))
                if datalist and isinstance(datalist, list):
                    first_item = datalist[0]
                    if "Number" in first_item:
                        return first_item["Number"]
        except Exception as e:
            print(f"无法获取默认单号: {e}")

        # 解析失败，打印 HTML 源代码的第 100-150 行
        try:
            lines = response.text.splitlines()  # 分割 HTML 为按行的列表
            snippet = "\n".join(lines[99:150])  # 提取第 100-150 行
            print("网页源代码 (第 100-150 行):")
            print(snippet)
        except Exception as inner_e:
            print(f"无法打印网页源代码: {inner_e}")

        return "解析错误"


    def on_generate_click(self):
        """点击生成按钮的处理逻辑"""
        try:
            # 禁用生成按钮并修改按钮文本
            self.generate_button.setText("正在生成，请稍后")
            self.generate_button.setEnabled(False)  # 禁用按钮
            QApplication.processEvents()  # 刷新界面

            if not self.session:
                QMessageBox.warning(self, "警告", "请先登录！")
                self.generate_button.setText("生成")
                self.generate_button.setEnabled(True)  # 恢复按钮
                return

            login_url = self.login_url_input.text()
            url1 = self.url1_input.text()
            base_url = self.config.get("base_url", "")
            include_stock_status = self.include_stock_status_input.currentText() == "是"
            finished_filter = self.finished_filter_input.currentIndex() - 1
            skip_negative_qty = self.skip_negative_qty_input.currentText() == "是"
            output_filename = self.output_filename_input.text().strip()

            if not output_filename:
                QMessageBox.warning(self, "警告", "输出文件名不能为空！")
                self.generate_button.setText("生成")
                self.generate_button.setEnabled(True)  # 恢复按钮
                return

            if self.date_mode_button.isChecked():
                target = self.target_date_input.date().toPyDate()
                mode = "date"
            else:
                target = self.target_number_input.text()
                mode = "orderNumber"

            # 使用 DataProcessor 处理数据
            response = self.session.get(url1)
            datalist = self.processor.extract_datalist(response.text)
            filtered_data = self.processor.filter_data(datalist, target, mode, finished_filter)
            data_rows = self.processor.fetch_and_format_data(filtered_data, self.session, base_url, include_stock_status, skip_negative_qty)

            # 检查是否有内容可写入 Excel
            if not data_rows:
                QMessageBox.warning(self, "提示", "解析到的内容为空，未生成文件。")
                self.generate_button.setText("生成")
                self.generate_button.setEnabled(True)  # 恢复按钮
                return

            # 保存到 Excel
            output_filepath = f"//VIVA303-WORK/Viva店面共享/{output_filename}.xlsx"
            self.write_to_excel(data_rows, output_filepath)
            QMessageBox.information(self, "完成", f"数据处理完成，文件已保存为：{output_filepath}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生错误: {str(e)}")

        finally:
            # 恢复生成按钮状态
            self.generate_button.setText("生成")
            self.generate_button.setEnabled(True)



    def write_to_excel(self, data_rows, filename):
        """保存数据到 Excel 文件"""
        headers = ["空A", "销售", "单号", "空D", "产品型号", "供货商", "数量", "顾客姓名", "电话", "家具自提", "留言", "货期", "订货"]
        wb = Workbook()
        ws = wb.active
        ws.title = "数据提取"

        ws.append(headers)
        for row in data_rows:
            ws.append(row)
        wb.save(filename)

    def get_icon_path(self):
        """获取图标路径"""
        base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, ICON_FILENAME)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    extractor_app = DataExtractorApp()
    extractor_app.show()
    sys.exit(app.exec_())

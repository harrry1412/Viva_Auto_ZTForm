import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import json
from datetime import datetime
from openpyxl import Workbook
from PyQt5.QtWidgets import (
    QApplication, QVBoxLayout, QLineEdit, QLabel, QPushButton, QComboBox, QWidget, QMessageBox, QDateEdit, QFileDialog
)
from PyQt5.QtWidgets import QToolButton, QHBoxLayout
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QFont, QIcon
import sys
import os

# 从配置文件加载配置
CONFIG_FILENAME = "config.json"
ICON_FILENAME = "app_icon.png"
APP_VERSION = 'V1.0.0'
APP_TITLE = f'VIVA自提单自动生成工具 {APP_VERSION} - Designed by Harry'


def load_config():
    """加载配置文件"""
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(base_path, CONFIG_FILENAME)
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"配置文件未找到: {config_path}")

def get_icon_path():
    """获取图标路径，兼容直接运行和打包后的路径"""
    if getattr(sys, 'frozen', False):
        # 如果是 PyInstaller 打包后的路径
        base_path = sys._MEIPASS
    else:
        # 如果是直接运行
        base_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_path, ICON_FILENAME)
    print(f"图标文件路径: {icon_path}, 存在: {os.path.exists(icon_path)}")
    return icon_path

def write_to_excel(data_rows, filename):
    headers = ["空A", "销售", "单号", "空D", "产品型号", "供货商", "数量", "顾客姓名", "电话", "家具自提", "留言", "货期", "订货"]
    wb = Workbook()
    ws = wb.active
    ws.title = "数据提取"

    # 写入表头
    ws.append(headers)

    # 写入数据行
    for row in data_rows:
        formatted_row = []
        for i, value in enumerate(row):
            if headers[i] == "数量":  # 如果是数量列，保持为数字
                try:
                    formatted_row.append(float(value) if '.' in str(value) else int(value))
                except ValueError:
                    formatted_row.append('')  # 如果转换失败，填入默认值0
            else:
                formatted_row.append(str(value))  # 其他列转为字符串
        ws.append(formatted_row)

    # 设置列宽（根据需要调整宽度）
    column_widths = [10, 15, 20, 10, 30, 20, 10, 20, 20, 15, 15, 15, 15]
    for i, width in enumerate(column_widths, 1):
        ws.column_dimensions[chr(64 + i)].width = width

    # 保存文件
    wb.save(filename)




def process_data(session, login_url, url1, base_url, target_date, include_stock_status, finished_filter, skip_negative_qty):
    driver = webdriver.Chrome()
    driver.get(login_url)
    QMessageBox.information(None, "提示", "请在浏览器中完成登录后点击确定继续。")

    cookies = driver.get_cookies()
    driver.quit()

    session = requests.Session()
    for cookie in cookies:
        session.cookies.set(cookie['name'], cookie['value'])

    response1 = session.get(url1)
    response1.raise_for_status()
    html_content1 = response1.text

    match_datalist = re.search(r"var\s+datalist\s*=\s*(\[.*?\]);", html_content1, re.DOTALL)
    if not match_datalist:
        QMessageBox.critical(None, "错误", "未找到 datalist 数据。")
        return []

    datalist_content = json.loads(match_datalist.group(1))
    filtered_data = [
        {
            "OriginalID": item.get("OriginalID"),
            "UserName": item.get("UserName", "无此字段"),
            "FirstName": item.get("FirstName", "无此字段"),
            "LastName": item.get("LastName", "无此字段"),
            "Number": item.get("Number", "无此字段")
        }
        for item in datalist_content
        if (finished_filter not in [0, 1] or item.get("finished") == finished_filter)
        and "Created" in item
        and datetime.strptime(item["Created"], "%Y-%m-%d %H:%M:%S").date() == target_date
    ]

    data_rows = []

    for data in filtered_data:
        original_id = data["OriginalID"]
        url2 = f"{base_url}{original_id}"
        response2 = session.get(url2)
        response2.raise_for_status()
        html_content2 = response2.text

        match_data = re.search(r"var\s+data\s*=\s*(\{.*?\});", html_content2, re.DOTALL)
        if match_data:
            data_content = json.loads(match_data.group(1))
            items = data_content.get("items", [])

            if skip_negative_qty:
                items = [item for item in items if float(item.get("Qty", 0)) >= 0]
            if skip_negative_qty and not items:
                continue

            phone_numbers = [
                data_content.get("PhoneCell", ""),
                data_content.get("PhoneHome", ""),
                data_content.get("PhoneOffice", "")
            ]
            phone_numbers = list(filter(None, phone_numbers))
            phone_combined = "/".join(phone_numbers) if phone_numbers else ""

            data_row = [
                "", data["UserName"], data["Number"], "", "", "", "",
                f"{data['FirstName']} {data['LastName']}", phone_combined, "", "", "", ""
            ]
            data_rows.append(data_row)

            for item in items:
                qty = float(item.get("Qty", 0))
                qty_oh = float(item.get("Qty_OH", 0))
                stock_status = ""
                if include_stock_status:
                    stock_status = "现货" if qty_oh - qty >= 1 else "需要订货"

                item_row = [
                    "", "", "", "", item.get("VendorPLU", ""), item.get("VendorName", ""), item.get("Qty", ""),
                    "", "", "", "", "", stock_status
                ]
                data_rows.append(item_row)

            data_rows.append(["" for _ in range(13)])

    return data_rows

class DataExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QIcon(get_icon_path()))

        screen_size = QApplication.primaryScreen().size()
        self.resize(screen_size.width() // 3, screen_size.height() * 3 // 4)

        layout = QVBoxLayout()

        help_button = QToolButton(self)
        help_button.setText("?")
        help_button.setToolTip("关于")
        help_button.clicked.connect(self.show_about_dialog)

        layout = QVBoxLayout(self)
        hbox = QHBoxLayout()
        hbox.addStretch()
        hbox.addWidget(help_button)
        layout.addLayout(hbox)

        font = QFont("Arial", 14)
        self.setFont(font)

        self.login_url_input = QLineEdit(self.config.get("login_url", ""))
        layout.addWidget(QLabel("登录页面 URL:"))
        layout.addWidget(self.login_url_input)

        self.url1_input = QLineEdit(self.config.get("url1", ""))
        layout.addWidget(QLabel("数据 URL:"))
        layout.addWidget(self.url1_input)

        self.target_date_input = QDateEdit()
        self.target_date_input.setCalendarPopup(True)
        self.target_date_input.setDate(QDate.currentDate())
        layout.addWidget(QLabel("要生成的日期:"))
        layout.addWidget(self.target_date_input)

        self.include_stock_status_input = QComboBox()
        self.include_stock_status_input.addItems(["否", "是"])
        self.include_stock_status_input.setCurrentIndex(0)  # 默认选择“否”
        layout.addWidget(QLabel("是否生成订货列:"))
        layout.addWidget(self.include_stock_status_input)

        self.finished_filter_input = QComboBox()
        self.finished_filter_input.addItems(["全部", "仅已标记完结的订单", "仅未标记完结的订单"])
        self.finished_filter_input.setCurrentIndex(0)  # 默认选择“全部”
        layout.addWidget(QLabel("生成已标记完结还是未完结的订单:"))
        layout.addWidget(self.finished_filter_input)

        self.skip_negative_qty_input = QComboBox()
        self.skip_negative_qty_input.addItems(["是", "否"])
        self.skip_negative_qty_input.setCurrentIndex(0)  # 默认选择“是”
        layout.addWidget(QLabel("跳过负库存记录:"))
        layout.addWidget(self.skip_negative_qty_input)

        self.generate_button = QPushButton("生成")
        self.generate_button.clicked.connect(self.on_generate_click)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)

    def on_generate_click(self):
        login_url = self.login_url_input.text()
        url1 = self.url1_input.text()
        base_url = self.config.get("base_url", "")
        target_date = self.target_date_input.date().toPyDate()
        include_stock_status = self.include_stock_status_input.currentText() == "是"
        finished_filter = self.finished_filter_input.currentIndex() - 1
        skip_negative_qty = self.skip_negative_qty_input.currentText() == "是"

        session = requests.Session()
        data_rows = process_data(session, login_url, url1, base_url, target_date, include_stock_status, finished_filter, skip_negative_qty)

        if data_rows:
            file_path = "//VIVA303-WORK/Viva店面共享/Viva自提单生成H.xlsx"
            try:
                write_to_excel(data_rows, file_path)
                QMessageBox.information(self, "完成", f"数据处理完成，文件已保存到: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"文件保存失败: {str(e)}")
        else:
            QMessageBox.information(self, "无记录", "选定条件下没有生成任何记录。")


    def show_about_dialog(self):
        QMessageBox.about(
            self,
            "关于",
            f'APP NAME: VIVA自提自动生成工具\nVERSION: {APP_VERSION}\nDEVELOPER: Haochu Chen\n\n'
            "Copyright © 2025 Haochu Chen\n"
            "All rights reserved.\n"
            "Unauthorized copying, modification, distribution, or use for commercial purposes is prohibited."
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(get_icon_path()))
    extractor_app = DataExtractorApp()
    extractor_app.show()
    sys.exit(app.exec_())

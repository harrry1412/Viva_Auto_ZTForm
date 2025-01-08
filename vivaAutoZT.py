import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
import json
from datetime import datetime
import csv
from PyQt5.QtWidgets import (
    QApplication, QVBoxLayout, QLineEdit, QLabel, QPushButton, QComboBox, QWidget, QMessageBox, QDateEdit
)
from PyQt5.QtCore import QDate
import sys

def write_to_csv(data_rows, filename="output.csv"):
    headers = ["空A", "销售", "单号", "空D", "产品型号", "供货商", "数量", "顾客姓名", "电话", "家具自提", "留言", "货期", "订货"]
    with open(filename, mode="w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        for row in data_rows:
            writer.writerow(row)

def process_data(session, login_url, url1, target_date, include_stock_status, finished_filter, skip_negative_qty):
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
        return

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

    base_url = "http://34.95.11.166/sales/document/document?id="
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

    write_to_csv(data_rows)
    QMessageBox.information(None, "完成", "数据处理完成，CSV 已生成！")

class DataExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("数据提取工具")
        self.resize(600, 400)
        layout = QVBoxLayout()

        self.login_url_input = QLineEdit("http://34.95.11.166/sales/account/login")
        layout.addWidget(QLabel("登录 URL:"))
        layout.addWidget(self.login_url_input)

        self.url1_input = QLineEdit("http://34.95.11.166/sales/document/index?page=1")
        layout.addWidget(QLabel("数据 URL1:"))
        layout.addWidget(self.url1_input)

        self.target_date_input = QDateEdit()
        self.target_date_input.setCalendarPopup(True)
        self.target_date_input.setDate(QDate.currentDate())
        layout.addWidget(QLabel("目标日期:"))
        layout.addWidget(self.target_date_input)

        self.include_stock_status_input = QComboBox()
        self.include_stock_status_input.addItems(["是", "否"])
        layout.addWidget(QLabel("是否计算订货状态:"))
        layout.addWidget(self.include_stock_status_input)

        self.finished_filter_input = QComboBox()
        self.finished_filter_input.addItems(["全部", "仅已完成 (1)", "仅未完成 (0)"])
        layout.addWidget(QLabel("筛选条件 (finished):"))
        layout.addWidget(self.finished_filter_input)

        self.skip_negative_qty_input = QComboBox()
        self.skip_negative_qty_input.addItems(["是", "否"])
        layout.addWidget(QLabel("跳过负库存记录:"))
        layout.addWidget(self.skip_negative_qty_input)

        self.generate_button = QPushButton("生成")
        self.generate_button.clicked.connect(self.on_generate_click)
        layout.addWidget(self.generate_button)

        self.setLayout(layout)

    def on_generate_click(self):
        login_url = self.login_url_input.text()
        url1 = self.url1_input.text()
        target_date = self.target_date_input.date().toPyDate()
        include_stock_status = self.include_stock_status_input.currentText() == "是"
        finished_filter = self.finished_filter_input.currentIndex() - 1
        skip_negative_qty = self.skip_negative_qty_input.currentText() == "是"

        session = requests.Session()
        process_data(session, login_url, url1, target_date, include_stock_status, finished_filter, skip_negative_qty)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    extractor_app = DataExtractorApp()
    extractor_app.show()
    sys.exit(app.exec_())

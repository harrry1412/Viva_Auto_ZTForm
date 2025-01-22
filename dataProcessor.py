import requests
from selenium import webdriver
from PyQt5.QtWidgets import QMessageBox
import re
import json
from datetime import datetime


class DataProcessor:
    def get_authenticated_session(self, login_url):
        """使用 Selenium 登录并返回已认证的 Requests 会话"""
        driver = webdriver.Chrome()
        driver.get(login_url)
        QMessageBox.information(None, "提示", "请在浏览器中完成登录后点击确定继续。")

        cookies = driver.get_cookies()
        driver.quit()

        session = requests.Session()
        for cookie in cookies:
            session.cookies.set(cookie['name'], cookie['value'])
        return session

    def extract_datalist(self, html_content):
        """从 HTML 中提取 datalist 数据"""
        match = re.search(r"var\s+datalist\s*=\s*(\[.*?\]);", html_content, re.DOTALL)
        if not match:
            return None
        return json.loads(match.group(1))

    def filter_data(self, datalist, target, mode, finished_filter):
        """根据模式和条件筛选数据"""
        if mode == "date":
            return [
                {
                    "OriginalID": item["OriginalID"],
                    "UserName": item.get("UserName", "无此字段"),
                    "FirstName": item.get("FirstName", "无此字段"),
                    "LastName": item.get("LastName", "无此字段"),
                    "Number": item.get("Number", "无此字段"),
                    "Created": item["Created"]
                }
                for item in datalist
                if (finished_filter not in [0, 1] or item.get("finished") == finished_filter)
                and "Created" in item
                and datetime.strptime(item["Created"], "%Y-%m-%d %H:%M:%S").date() == target
            ]
        elif mode == "orderNumber":
            return [
                {
                    "OriginalID": item["OriginalID"],
                    "UserName": item.get("UserName", "无此字段"),
                    "FirstName": item.get("FirstName", "无此字段"),
                    "LastName": item.get("LastName", "无此字段"),
                    "Number": item.get("Number", "无此字段")
                }
                for item in datalist
                if (finished_filter not in [0, 1] or item.get("finished") == finished_filter)
                and item.get("Number") == target
            ]
        else:
            raise ValueError(f"未知模式: {mode}")

    def fetch_and_format_data(self, filtered_data, session, base_url, include_stock_status, skip_negative_qty):
        """根据筛选后的数据提取详细信息并格式化为 Excel 行"""
        data_rows = []
        for data in filtered_data:
            original_id = data["OriginalID"]
            url2 = f"{base_url}{original_id}"
            try:
                response2 = session.get(url2)
                response2.raise_for_status()
                match_data = re.search(r"var\s+data\s*=\s*(\{.*?\});", response2.text, re.DOTALL)
                if not match_data:
                    continue
                data_content = json.loads(match_data.group(1))

                # 格式化基础数据
                phone_combined = self.combine_phone_numbers(data_content)
                data_row = [
                    "", data["UserName"], data["Number"], "", "", "", "",
                    f"{data['FirstName']} {data['LastName']}", phone_combined, "", "", "", ""
                ]
                data_rows.append(data_row)

                # 格式化详细项目数据
                items = data_content.get("items", [])
                if skip_negative_qty:
                    items = [item for item in items if float(item.get("Qty", 0)) >= 0]
                for item in items:
                    qty = float(item.get("Qty", 0))
                    qty_oh = float(item.get("Qty_OH", 0))
                    stock_status = ""
                    # 仅在用户选择生成订货列时计算订货状态
                    if include_stock_status:
                        stock_status = "现货" if qty_oh - qty >= 1 else "需要订货"

                    item_row = [
                        "", "", "", "", item.get("VendorPLU", ""), item.get("VendorName", ""),
                        item.get("Qty", ""), "", "", "", "", "", stock_status if include_stock_status else ""
                    ]
                    data_rows.append(item_row)

                # 添加空行分隔订单
                data_rows.append(["" for _ in range(13)])
            except Exception as e:
                print(f"提取数据失败: {str(e)}")
                continue
        return data_rows


    def combine_phone_numbers(self, data_content):
        """合并电话号码"""
        phone_numbers = [
            data_content.get("PhoneCell", ""),
            data_content.get("PhoneHome", ""),
            data_content.get("PhoneOffice", "")
        ]
        return "/".join(filter(None, phone_numbers))

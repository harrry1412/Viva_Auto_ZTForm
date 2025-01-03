import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox
import re
import time
import json
from datetime import datetime

def wait_for_user_action():
    """弹出一个Tkinter窗口，等待用户点击继续按钮"""
    def on_continue():
        nonlocal user_ready
        user_ready = True
        root.destroy()

    user_ready = False
    root = tk.Tk()
    root.title("等待用户操作")
    root.geometry("300x100")
    label = tk.Label(root, text="请完成登录后点击继续按钮", font=("Arial", 12))
    label.pack(pady=10)
    button = tk.Button(root, text="继续", command=on_continue, font=("Arial", 12), bg="green", fg="white")
    button.pack(pady=10)
    root.mainloop()

    return user_ready

def process_additional_urls(filtered_data, session):
    base_url = "http://34.95.11.166/sales/document/document?id="
    for data in filtered_data:
        original_id = data["OriginalID"]
        url2 = f"{base_url}{original_id}"
        response2 = session.get(url2)
        response2.raise_for_status()  # 检查目标页面请求是否成功
        html_content2 = response2.text

        match_data = re.search(r"var\s+data\s*=\s*(\{.*?\});", html_content2, re.DOTALL)
        if match_data:
            data_content_raw = match_data.group(1)
            data_content = json.loads(data_content_raw)  # 转换为 Python 数据结构

            # 获取 items 数组中的数据
            items = data_content.get("items", [])
            for item in items:
                item_data = {
                    "DocumentID": item.get("DocumentID", "无此字段"),
                    "VendorPLU": item.get("VendorPLU", "无此字段"),
                    "Qty": item.get("Qty", "无此字段"),
                    "Qty_OH": item.get("Qty_OH", "无此字段")
                }

                # 打印整合后的数据
                combined_data = {**data, **item_data}
                print("整合数据:", combined_data)

def login_and_extract_data(url1, login_url, target_date):
    try:
        # 使用 Selenium 打开浏览器窗口
        driver = webdriver.Chrome()  # 请确保已安装 ChromeDriver 并配置在 PATH 中

        # 打开登录页面
        driver.get(login_url)

        # 弹出窗口等待用户完成登录
        print("请在浏览器中完成登录，随后点击弹出窗口中的继续按钮...")
        user_ready = wait_for_user_action()
        if not user_ready:
            return "用户未确认继续，操作中止。", None

        # 登录完成后，获取登录后的 Cookie
        cookies = driver.get_cookies()

        # 使用 requests.Session 模拟登录后的请求
        session = requests.Session()
        for cookie in cookies:
            session.cookies.set(cookie['name'], cookie['value'])

        # 请求第一个页面，提取 var datalist = 的内容
        response1 = session.get(url1)
        response1.raise_for_status()  # 检查目标页面请求是否成功
        html_content1 = response1.text

        match_datalist = re.search(r"var\s+datalist\s*=\s*(\[.*?\]);", html_content1, re.DOTALL)
        if match_datalist:
            datalist_content_raw = match_datalist.group(1)
            datalist_content = json.loads(datalist_content_raw)  # 转换为 Python 数据结构

            # 找到 finished 为 0 且 Created 等于 target_date 的元素，并提取相关字段
            filtered_data = [
                {
                    "OriginalID": item.get("OriginalID"),
                    "UserName": item.get("UserName", "无此字段"),
                    "FirstName": item.get("FirstName", "无此字段"),
                    "LastName": item.get("LastName", "无此字段"),
                    "Number": item.get("Number", "无此字段")
                }
                for item in datalist_content
                if item.get("finished") == 0
                and "Created" in item
                and datetime.strptime(item["Created"], "%Y-%m-%d %H:%M:%S").date() == target_date
            ]

            # 处理生成的新 URL 并提取数据
            process_additional_urls(filtered_data, session)

        # 关闭浏览器
        driver.quit()

    except Exception as e:
        print(f"发生错误: {e}")

# 登录页面和目标页面的URL
login_url = "http://34.95.11.166/sales/account/login"
url1 = "http://34.95.11.166/sales/document/index?page=1"

# 设置目标日期
target_date = datetime.strptime("2025-01-02", "%Y-%m-%d").date()

login_and_extract_data(url1, login_url, target_date)

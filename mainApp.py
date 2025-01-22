import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton,
    QComboBox, QRadioButton, QDateEdit, QMessageBox, QButtonGroup
)
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QFont, QIcon
from openpyxl import Workbook
from dataProcessor import DataProcessor  # 引入 DataProcessor
import os
import json

# 全局常量
CONFIG_FILENAME = "config.json"
ICON_FILENAME = "app_icon.png"
APP_TITLE = "VIVA自提单自动生成工具 V1.0.0"


class DataExtractorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()  # 加载配置文件
        self.processor = DataProcessor()  # 实例化数据处理类
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

        # 模式切换事件
        self.date_mode_button.toggled.connect(self.update_input_fields)

        # 日期输入
        self.target_date_input = QDateEdit()
        self.target_date_input.setCalendarPopup(True)
        self.target_date_input.setDate(QDate.currentDate())

        # 单号输入
        self.target_number_input = QLineEdit("单号Test")
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

    def on_generate_click(self):
        """点击生成按钮的处理逻辑"""
        try:
            login_url = self.login_url_input.text()
            url1 = self.url1_input.text()
            base_url = self.config.get("base_url", "")
            include_stock_status = self.include_stock_status_input.currentText() == "是"
            finished_filter = self.finished_filter_input.currentIndex() - 1
            skip_negative_qty = self.skip_negative_qty_input.currentText() == "是"
            output_filename = self.output_filename_input.text().strip()

            if not output_filename:
                QMessageBox.warning(self, "警告", "输出文件名不能为空！")
                return

            if self.date_mode_button.isChecked():
                target = self.target_date_input.date().toPyDate()
                mode = "date"
            else:
                target = self.target_number_input.text()
                mode = "orderNumber"

            # 使用 DataProcessor 处理数据
            session = self.processor.get_authenticated_session(login_url)
            response = session.get(url1)
            datalist = self.processor.extract_datalist(response.text)
            filtered_data = self.processor.filter_data(datalist, target, mode, finished_filter)
            data_rows = self.processor.fetch_and_format_data(filtered_data, session, base_url, include_stock_status, skip_negative_qty)

            # 保存到 Excel
            output_filepath = f"//VIVA303-WORK/Viva店面共享/{output_filename}.xlsx"
            self.write_to_excel(data_rows, output_filepath)
            QMessageBox.information(self, "完成", f"数据处理完成，文件已保存为：{output_filepath}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"发生错误: {str(e)}")

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

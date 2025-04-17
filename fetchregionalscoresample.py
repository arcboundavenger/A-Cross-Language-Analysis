import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import random
import os

# 要爬取的 appID 列表
app_ids = [
    2358720, 1623730
]

# 设置 ChromeDriver 的路径
service = Service(ChromeDriverManager().install())

# 创建一个 Excel 文件
output_file = "regional_review_scores.xlsx"

# 初始创建 Excel 文件
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for app_id in app_ids:
        url = f"https://www.togeproductions.com/SteamScout/steamAPI.php?appID={app_id}"

        # 初始化 Selenium WebDriver
        driver = webdriver.Chrome(service=service)
        driver.get(url)

        # 随机等待页面加载，时间范围从 40 到 70 秒
        time.sleep(random.randint(40, 70))

        # 获取页面内容
        html = driver.page_source

        # 使用 pandas 读取 HTML 中的表格
        try:
            tables = pd.read_html(html)
            if tables:  # 检查是否有表格
                df = tables[0]

                # 对调第一列和第二列，包括标题
                if df.shape[1] >= 2:  # 确保有至少两列
                    # 交换列标题
                    df.columns = [df.columns[1], df.columns[0]] + list(df.columns[2:])
                    # 交换数据
                    df.iloc[:, [0, 1]] = df.iloc[:, [1, 0]].values

                # 将数据写入 Excel，表名为 appID
                df.to_excel(writer, sheet_name=str(app_id), index=False)
                print(f"数据已保存到 {output_file} 的工作表 {app_id}")
            else:
                print(f"No tables found for appID: {app_id}")
        except Exception as e:
            print(f"Error reading tables for appID {app_id}: {e}")

        # 关闭浏览器
        driver.quit()

print(f"所有数据已保存到 {output_file}")
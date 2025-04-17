import requests
import pandas as pd
import time
import random


def get_median_playtime(app_ids):
    median_playtimes = {}

    for app_id in app_ids:
        # 构造 API 请求的 URL
        url = f"https://steamspy.com/api.php?request=appdetails&appid={app_id}"

        try:
            response = requests.get(url)
            response.raise_for_status()  # 检查请求是否成功
            data = response.json()

            # 提取中位数游戏时间
            if 'median_forever' in data:
                median_playtimes[app_id] = data['median_forever'] / 60  # 转换为小时
            else:
                median_playtimes[app_id] = None  # 如果没有找到中位数

            # 创建 DataFrame
            df = pd.DataFrame(list(median_playtimes.items()), columns=['App ID', 'Median Playtime (Hours)'])

            # 保存为 Excel 文件
            output_file = 'steamspy_api_results.xlsx'
            df.to_excel(output_file, index=False)
            print(f"Current results saved to {output_file}")

            # 等待随机 1 到 2 秒
            time.sleep(random.uniform(1, 2))

        except requests.exceptions.RequestException as e:
            print(f"Error fetching data for app ID {app_id}: {e}")
            median_playtimes[app_id] = None

    return median_playtimes


# 示例 appid 列表
app_ids = [
    2140330, 1190970
] # 这里可以替换为你自己的 appid 列表
median_playtimes = get_median_playtime(app_ids)

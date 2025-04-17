import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# 从指定的 Excel 文件和工作表中读取数据
file_path = 'Gen9_5000.xlsx'  # 替换为您的文件路径
df = pd.read_excel(file_path, sheet_name='All')

# 将 releaseDate 转换为日期格式
df['releaseDate'] = pd.to_datetime(df['releaseDate'])

# 按季度分组并计算平均分
grouped = df.groupby(pd.Grouper(key='releaseDate', freq='Q'))[['ChineseReviewScore', 'EnglishReviewScore', 'reviewScore']].mean().reset_index()

# 舍弃最后的数据点（2024年第四季度）
grouped = grouped[grouped['releaseDate'] < '2024-10-01']

# 计算每个季度的游戏数量
game_counts = df.groupby(pd.Grouper(key='releaseDate', freq='Q'))['name'].count().reset_index(name='gameCount')

# 合并游戏数量和评分数据
merged = pd.merge(grouped, game_counts, on='releaseDate', how='left')

# 创建图形和坐标轴
fig, ax1 = plt.subplots(figsize=(12, 6))

# 绘制评分数据
ax1.plot(merged['releaseDate'], merged['ChineseReviewScore'], marker='o', label='Chinese Review Score', color='blue')
ax1.plot(merged['releaseDate'], merged['EnglishReviewScore'], marker='o', label='English Review Score', color='orange')
ax1.plot(merged['releaseDate'], merged['reviewScore'], marker='o', label='Overall Review Score', color='green')

# 设置主坐标轴标签
ax1.set_xlabel('Release Quarter')
ax1.set_ylabel('Scores', color='black')
ax1.tick_params(axis='y', labelcolor='black')

# 格式化 x 轴为季度
ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y Q%q'))
ax1.xaxis.set_major_locator(mdates.MonthLocator(interval=3))  # 每3个月一个刻度

# 手动设置 x 轴标签
ax1.set_xticks(merged['releaseDate'])
ax1.set_xticklabels([f"{date.year} Q{((date.month - 1) // 3) + 1}" for date in merged['releaseDate']], rotation=45)

# 创建第二个坐标轴
ax2 = ax1.twinx()
ax2.bar(merged['releaseDate'], merged['gameCount'], alpha=0.3, label='Game Count', color='gray', width=20, align='center')

# 设置第二个坐标轴标签
ax2.set_ylabel('Game Count', color='gray')
ax2.tick_params(axis='y', labelcolor='gray')

# 添加图例
fig.tight_layout()
ax1.legend(loc='upper left')
ax2.legend(loc='upper right')

# 设置图表标题
plt.title('Average Review Scores and Game Counts by Release Quarter')

# 显示图表
plt.show()
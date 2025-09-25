import pandas as pd
import matplotlib.pyplot as plt

# 读取Excel文件
file_path = 'games_studied.xlsx'
df = pd.read_excel(file_path, sheet_name='All')

# 转换为数值类型
df['price'] = pd.to_numeric(df['price'], errors='coerce')
df['reviewScore'] = pd.to_numeric(df['reviewScore'], errors='coerce')
df['ScoreGap'] = pd.to_numeric(df['ScoreGap'], errors='coerce')

# 定义价格区间
bins_price = [0, 10, 20, 30, 40, 50, 60, 70]
labels_price = ['[0, 10)', '[10, 20)', '[20, 30)', '[30, 40)', '[40, 50)',
                '[50, 60)', '[60, 70)']
df['price_group'] = pd.cut(df['price'], bins=bins_price, labels=labels_price, right=False)

# 创建图形
fig, axs = plt.subplots(1, 2, figsize=(12, 6))

# 左图：ScoreGap按ReviewScore分组
bins_score = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
labels_score = ['[0, 10)', '[10, 20)', '[20, 30)', '[30, 40)', '[40, 50)',
                '[50, 60)', '[60, 70)', '[70, 80)', '[80, 90)', '[90, 100)']
df['reviewScore_group'] = pd.cut(df['reviewScore'], bins=bins_score, labels=labels_score, right=False)

# 绘制左图箱线图
df.boxplot(column='ScoreGap', by='reviewScore_group', ax=axs[0])
# 计算并标注平均值
means_score = df.groupby('reviewScore_group')['ScoreGap'].mean()
for i, mean in enumerate(means_score):
    axs[0].scatter([i + 1], [mean], color='red', zorder=5)
axs[0].set_xlabel('Global Review Score', fontsize=12)
axs[0].set_ylabel('Score Gap (ENG - CHN)', fontsize=12)
axs[0].tick_params(axis='x', rotation=45)
# 添加字母标记 a
axs[0].text(0.95, 0.95, 'a', transform=axs[0].transAxes, fontsize=14, fontweight='bold',
            horizontalalignment='right', verticalalignment='top')

# 右图：ScoreGap按Price分组
df.boxplot(column='ScoreGap', by='price_group', ax=axs[1])
# 计算并标注平均值
means_price = df.groupby('price_group')['ScoreGap'].mean()
for i, mean in enumerate(means_price):
    axs[1].scatter([i + 1], [mean], color='red', zorder=5)
axs[1].set_xlabel('Price Interval', fontsize=12)
axs[1].set_ylabel('Score Gap (ENG - CHN)', fontsize=12)
axs[1].tick_params(axis='x', rotation=45)
# 添加字母标记 b - 修正了这里，使用axs[1].transAxes而不是axs[0].transAxes
axs[1].text(0.95, 0.95, 'b', transform=axs[1].transAxes, fontsize=14, fontweight='bold',
            horizontalalignment='right', verticalalignment='top')

# 移除所有标题
for ax in axs:
    ax.set_title('')  # 移除子标题
plt.suptitle('')      # 移除主标题

plt.tight_layout()
fig.savefig('figure5.png', dpi=600, bbox_inches='tight')
plt.show()
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

df.boxplot(column='ScoreGap', by='reviewScore_group', ax=axs[0])
axs[0].set_xlabel('Review Score Group', fontsize=12)
axs[0].set_ylabel('Score Gap', fontsize=12)
axs[0].tick_params(axis='x', rotation=45)

# 右图：ScoreGap按Price分组
df.boxplot(column='ScoreGap', by='price_group', ax=axs[1])
axs[1].set_xlabel('Price Group', fontsize=12)
axs[1].set_ylabel('Score Gap', fontsize=12)
axs[1].tick_params(axis='x', rotation=45)

# 移除所有标题
for ax in axs:
    ax.set_title('')  # 移除子标题
plt.suptitle('')      # 移除主标题

plt.tight_layout()
fig.savefig('figure5.png', dpi=600, bbox_inches='tight')
plt.show()
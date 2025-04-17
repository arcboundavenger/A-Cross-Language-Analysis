import pandas as pd
import matplotlib.pyplot as plt

# 读取Excel文件，指定工作表名称为 'All'
file_path = 'Gen9_5000.xlsx'  # 请将此处替换为您的文件名
df = pd.read_excel(file_path, sheet_name='All')

# 将价格和评分列转换为数值类型
df['price'] = pd.to_numeric(df['price'], errors='coerce')
df['reviewScore'] = pd.to_numeric(df['reviewScore'], errors='coerce')
df['ScoreGap'] = pd.to_numeric(df['ScoreGap'], errors='coerce')

# 定义价格区间
bins_price = [0, 10, 20, 30, 40, 50, 60, 70]
labels_price = ['[0, 10)', '[10, 20)', '[20, 30)', '[30, 40)', '[40, 50)',
                '[50, 60)', '[60, 70)']

# 创建价格区间列
df['price_group'] = pd.cut(df['price'], bins=bins_price, labels=labels_price, right=False)

# 绘制第一个图形：三个箱型图横向排列
fig1, axs = plt.subplots(1, 3, figsize=(18, 6))

# 统一纵轴范围
y_min, y_max = -10, 110

def plot_box_with_mean(ax, column, title):
    df.boxplot(column=column, by='price_group', ax=ax)
    ax.set_xlabel('Price Group', fontsize=10)
    ax.set_ylabel('')  # 去掉纵轴标题
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45, fontsize=9)
    ax.set_ylim(y_min, y_max)  # 统一y轴范围

    # 计算并标注平均值
    means = df.groupby('price_group')[column].mean()
    for i, mean in enumerate(means):
        ax.scatter([i + 1], [mean], color='red', label='Mean' if i == 0 else "", zorder=5)

# 中文评分箱型图
plot_box_with_mean(axs[0], 'ChineseReviewScore', 'Chinese Review Score')

# 英文评分箱型图
plot_box_with_mean(axs[1], 'EnglishReviewScore', 'English Review Score')

# reviewScore箱型图
plot_box_with_mean(axs[2], 'reviewScore', 'Review Score')

# 添加图例
axs[0].legend()

plt.suptitle('')  # 移除默认的标题
plt.tight_layout()  # 调整布局以适应标签

# 绘制第二个图形：ScoreGap和ReviewScore根据分组
fig2, axs2 = plt.subplots(1, 2, figsize=(12, 6))

# ScoreGap根据ReviewScore分组的箱型图
axs2[0].set_title('Score Gap by Review Score Group', fontsize=14)
bins_score = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100]
labels_score = ['[0, 10)', '[10, 20)', '[20, 30)', '[30, 40)', '[40, 50)',
                '[50, 60)', '[60, 70)', '[70, 80)', '[80, 90)', '[90, 100)']
df['reviewScore_group'] = pd.cut(df['reviewScore'], bins=bins_score, labels=labels_score, right=False)
df.boxplot(column='ScoreGap', by='reviewScore_group', ax=axs2[0])
axs2[0].set_xlabel('Review Score Group', fontsize=12)
axs2[0].set_ylabel('Score Gap', fontsize=12)
axs2[0].tick_params(axis='x', rotation=45)

# 计算并标注平均值
means_score = df.groupby('reviewScore_group')['ScoreGap'].mean()
for i, mean in enumerate(means_score):
    axs2[0].scatter([i + 1], [mean], color='red', zorder=5)

# ScoreGap根据价格分组的箱型图
axs2[1].set_title('Score Gap by Price Group', fontsize=14)
df.boxplot(column='ScoreGap', by='price_group', ax=axs2[1])
axs2[1].set_xlabel('Price Group', fontsize=12)
axs2[1].set_ylabel('Score Gap', fontsize=12)
axs2[1].tick_params(axis='x', rotation=45)

# 计算并标注平均值
means_price = df.groupby('price_group')['ScoreGap'].mean()
for i, mean in enumerate(means_price):
    axs2[1].scatter([i + 1], [mean], color='red', zorder=5)

plt.suptitle('')  # 移除默认的标题
plt.tight_layout()  # 调整布局以适应标签

# 只调用一次 plt.show()
plt.show()
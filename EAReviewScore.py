import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 读取 Excel 文件
file_path = 'Gen9_5000.xlsx'  # 替换为您的文件路径
df = pd.read_excel(file_path, sheet_name='All')

# 筛选 EA? 列为 1 和 0 的数据
df_ea_0 = df[df['EA?'] == 0]
df_ea_1 = df[df['EA?'] == 1]

# 设置图形的大小
plt.figure(figsize=(12, 6))

# 计算所有评分数据的最大最小值，用于统一y轴范围
all_scores = np.concatenate([
    df_ea_0['ChineseReviewScore'],
    df_ea_1['ChineseReviewScore'],
    df_ea_0['EnglishReviewScore'],
    df_ea_1['EnglishReviewScore']
])
y_min, y_max = np.floor(min(all_scores)), np.ceil(max(all_scores))

# 绘制 Chinese Review Score 的箱线图
plt.subplot(1, 2, 1)
plt.boxplot([df_ea_0['ChineseReviewScore'], df_ea_1['ChineseReviewScore']],
            labels=['Non-Early Access', 'Early Access'])
plt.title('Chinese Review Score by Early Access Status')
plt.ylabel('Scores')
plt.ylim(10, 110)  # 设置统一的y轴范围
plt.grid(axis='y')
# 计算并标注平均值
means = [df_ea_0['ChineseReviewScore'].mean(), df_ea_1['ChineseReviewScore'].mean()]
for i, mean in enumerate(means):
    plt.plot(i + 1, mean, marker='o', color='red', label='Mean' if i == 0 else "")
plt.legend()

# 绘制 English Review Score 的箱线图
plt.subplot(1, 2, 2)
plt.boxplot([df_ea_0['EnglishReviewScore'], df_ea_1['EnglishReviewScore']],
            labels=['Non-Early Access', 'Early Access'])
plt.title('English Review Score by Early Access Status')
plt.ylabel('Scores')
plt.ylim(10, 110)  # 设置统一的y轴范围
plt.grid(axis='y')
# 计算并标注平均值
means = [df_ea_0['EnglishReviewScore'].mean(), df_ea_1['EnglishReviewScore'].mean()]
for i, mean in enumerate(means):
    plt.plot(i + 1, mean, marker='o', color='red', label='Mean' if i == 0 else "")
plt.legend()

# 调整布局
plt.tight_layout()

# 显示图形
plt.show()
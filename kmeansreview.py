# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from mpl_toolkits.mplot3d import Axes3D
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
import numpy as np

# 设置字体以支持中文
matplotlib.rcParams['font.family'] = 'SimHei'  # 使用黑体
matplotlib.rcParams['axes.unicode_minus'] = False  # 处理负号显示问题

# 读取 Excel 文件
file_path = 'Gen9_5000.xlsx'
df = pd.read_excel(file_path, sheet_name='All')

# 将相关列转换为数值型，如果无法转换则设置为 NaN
df['ScoreGap'] = pd.to_numeric(df['ScoreGap'], errors='coerce')
df['price'] = pd.to_numeric(df['price'], errors='coerce')
df['reviewScore'] = pd.to_numeric(df['reviewScore'], errors='coerce')
df['medianPlaytime'] = pd.to_numeric(df['medianPlaytime'], errors='coerce')

# 过滤掉包含 NaN 的行
df = df.dropna(subset=['ScoreGap', 'price', 'reviewScore', 'medianPlaytime'])

# 选择要聚类的特征
X = df[['ScoreGap', 'price', 'reviewScore']]

# 数据标准化
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)

# K-means 聚类（设置聚类数为4）
optimal_k = 4
kmeans = KMeans(n_clusters=optimal_k, random_state=42)
df['cluster'] = kmeans.fit_predict(X_scaled)

# 可视化聚类结果（3D 散点图）
fig = plt.figure(figsize=(12, 8))
ax = fig.add_subplot(111, projection='3d')

# 使用 scatter 方法绘制 3D 散点图
scatter = ax.scatter(df['ScoreGap'], df['price'], df['reviewScore'], c=df['cluster'], cmap='viridis', alpha=0.6)

# 设置标题和坐标轴标签
ax.set_title('K-means Clustering Result (3D View)')
ax.set_xlabel('Score Gap')
ax.set_ylabel('Price')
ax.set_zlabel('Review Score')

# 添加色点图例
colors = plt.cm.viridis(np.linspace(0, 1, optimal_k))
labels = [f'Cluster {i}' for i in range(optimal_k)]
for idx, (color, label) in enumerate(zip(colors, labels)):
    ax.scatter([], [], [], color=color, label=label, alpha=0.6)

ax.legend(title="Clusters", loc='upper left')

plt.tight_layout()
plt.show()

# 输出到 Excel
output_file = 'clustered_results_with_medianPlaytime.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for cluster_id in range(optimal_k):
        cluster_data = df[df['cluster'] == cluster_id][['steamId', 'name', 'ScoreGap', 'price', 'reviewScore', 'medianPlaytime']]
        cluster_data.to_excel(writer, sheet_name=f'Cluster_{cluster_id}', index=False)

print(f"输出成功，文件名为：{output_file}")
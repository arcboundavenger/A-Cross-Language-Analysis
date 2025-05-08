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
file_path = 'games_studied.xlsx'
df = pd.read_excel(file_path, sheet_name='All')

# 将相关列转换为数值型
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

# 绘制3D散点图
scatter = ax.scatter(
    df['ScoreGap'],
    df['price'],
    df['reviewScore'],
    c=df['cluster'],
    cmap='viridis',
    alpha=0.6,
    s=50  # 点的大小
)

# 设置坐标轴标签
ax.set_xlabel('Score Gap', labelpad=10)
ax.set_ylabel('Price', labelpad=10)
ax.set_zlabel('Review Score', labelpad=10)

# 添加色点图例
colors = plt.cm.viridis(np.linspace(0, 1, optimal_k))
labels = [f'Cluster {i}' for i in range(optimal_k)]
legend_elements = [plt.Line2D([0], [0], marker='o', color='w',
                             markerfacecolor=colors[i], markersize=10,
                             label=labels[i])
                  for i in range(optimal_k)]
ax.legend(handles=legend_elements, title="Clusters", loc='upper right')

# 调整视角（俯仰角30度，方位角45度）
ax.view_init(elev=30, azim=45)

# 调整布局并保存图像
plt.tight_layout()
fig.savefig('figure7.png', dpi=600, bbox_inches='tight', transparent=False)
plt.show()

# 输出到 Excel
output_file = 'clustered_results.xlsx'
with pd.ExcelWriter(output_file) as writer:
    for cluster_id in range(optimal_k):
        cluster_data = df[df['cluster'] == cluster_id][['steamId', 'name', 'ScoreGap', 'price', 'reviewScore']]
        cluster_data.to_excel(writer, sheet_name=f'Cluster_{cluster_id}', index=False)

print(f"聚类结果已保存至: {output_file}")
print(f"3D可视化图已保存为: figure7.png (600 DPI)")
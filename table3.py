import pandas as pd
import numpy as np
from scipy import stats

# 读取数据
file_path = 'games_studied_regional_score.xlsx'  # 替换为你的文件路径
data = pd.read_excel(file_path, sheet_name='Sheet1')

# 提取全球平均值行
global_means = data.iloc[-1, 1:].values
global_means = pd.to_numeric(global_means, errors='coerce')
global_mean_value = np.nanmean(global_means)

# 创建结果DataFrame
results = pd.DataFrame(columns=['Language', 'Mean', 'Global Mean', 'z-statistic', 'p-value'])

# 逐行计算
for index, row in data.iterrows():
    if index == len(data) - 1:
        continue
    language = row['Language']
    region_data = pd.to_numeric(row[1:], errors='coerce').dropna().values

    if len(region_data) < 2:
        print(f"跳过 {language}，因为没有足够的数据进行z检验。")
        continue

    language_mean = np.nanmean(region_data)
    language_std = np.nanstd(region_data, ddof=1)  # 计算样本标准差
    n1 = len(region_data)  # 样本 1 的大小
    n2 = len(global_means)  # 样本 2 的大小
    global_std = np.nanstd(global_means, ddof=1)  # 计算全球均值的样本标准差

    # 计算 z 统计量
    z_statistic = (language_mean - global_mean_value) / np.sqrt((language_std ** 2 / n1) + (global_std ** 2 / n2))

    # 计算大于全球均值的概率（右尾 p 值）
    p_value = 1 - stats.norm.cdf(z_statistic)

    # 存储结果
    results = pd.concat([results, pd.DataFrame({
        'Language': [language],
        'Mean': [language_mean],
        'Global Mean': [global_mean_value],
        'z-statistic': [z_statistic],
        'p-value': [p_value]
    })], ignore_index=True)

    # 判断显著性
    alpha = 0.05
    if p_value < alpha and z_statistic > 0:
        print(f"{language}: 拒绝零假设，样本均值大于全球均值")
    else:
        print(f"{language}: 无法拒绝零假设，样本均值不大于全球均值")

# 输出结果表格
print(results)

# 保存结果为Excel文件
output_file_path = 'Table 3.xlsx'
results.to_excel(output_file_path, index=False)
print(f"结果已保存到 {output_file_path}")
import pandas as pd
import numpy as np
from scipy import stats

# 读取数据
file_path = 'games_studied_regional_score.xlsx'
data = pd.read_excel(file_path, sheet_name='Sheet1')

# 提取全球平均值行
global_means = data.iloc[-1, 1:].values
global_means = pd.to_numeric(global_means, errors='coerce')
global_mean_value = np.nanmean(global_means)

# 创建结果DataFrame
results = pd.DataFrame(columns=[
    'Language', 'Mean', 'Global Mean', 'z-statistic',
    'p-value (raw)', 'p-value (adjusted)', 'Significance', 'Direction'
])

# 计算比较次数（排除最后一行全球平均值）
num_comparisons = len(data) - 1


# 显著性标记函数
def get_significance_stars(adjusted_p):
    if adjusted_p < 0.001:
        return '***'
    elif adjusted_p < 0.01:
        return '**'
    elif adjusted_p < 0.05:
        return '*'
    else:
        return ''


# 逐行计算
for index, row in data.iterrows():
    if index == len(data) - 1:
        continue  # 跳过全球平均值行

    language = row['Language']
    region_data = pd.to_numeric(row[1:], errors='coerce').dropna().values

    if len(region_data) < 2:
        print(f"Skipped {language}: insufficient data")
        continue

    # 计算统计量
    language_mean = np.nanmean(region_data)
    language_std = np.nanstd(region_data, ddof=1)
    n1 = len(region_data)
    n2 = len(global_means)
    global_std = np.nanstd(global_means, ddof=1)

    # 计算Z统计量和双边p值
    z_statistic = (language_mean - global_mean_value) / np.sqrt((language_std ** 2 / n1) + (global_std ** 2 / n2))
    raw_p = 2 * (1 - stats.norm.cdf(abs(z_statistic)))  # 双边检验

    # Bonferroni校正
    adjusted_p = min(1, raw_p * num_comparisons)

    # 判断方向和显著性
    direction = "Higher" if z_statistic > 0 else "Lower"
    stars = get_significance_stars(adjusted_p)

    # 存储结果
    results = pd.concat([results, pd.DataFrame({
        'Language': [language],
        'Mean': [language_mean],
        'Global Mean': [global_mean_value],
        'z-statistic': [z_statistic],
        'p-value (raw)': [raw_p],
        'p-value (adjusted)': [adjusted_p],
        'Significance': [stars],
        'Direction': [direction]
    })], ignore_index=True)

# 输出结果
print(results)

# 保存结果
output_path = 'Table 2.xlsx'
results.to_excel(output_path, index=False)
print(f"\nResults saved to {output_path}")
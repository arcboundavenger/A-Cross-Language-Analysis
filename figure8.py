import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import chi2_contingency, fisher_exact
from statsmodels.stats.multitest import multipletests


def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheets = [sheet for sheet in xls.sheet_names if sheet != 'Sheet1']
    dfs = [pd.read_excel(xls, sheet_name=sheet)[['is_recommended', 'language', 'dominant_emotion']]
           for sheet in sheets]
    return pd.concat(dfs, ignore_index=True)


def process_data(df):
    valid_df = df[df['is_recommended'].isin(['positive', 'negative'])]
    emotions = ['Anger', 'Disgust', 'Anticipation', 'Fear', 'Joy', 'Sadness', 'Trust', 'Surprise']

    def process_subset(subset):
        return pd.crosstab(
            index=subset['language'],
            columns=subset['dominant_emotion'],
            normalize='index'
        ).reindex(columns=emotions, fill_value=0).reindex(['english', 'schinese'], fill_value=0) * 100

    return (process_subset(valid_df[valid_df['is_recommended'] == 'positive']).T,
            process_subset(valid_df[valid_df['is_recommended'] == 'negative']).T)


# 新增：三级星号标注函数
def get_significance_stars(p):
    """返回统计学显著性标记 (***p<0.001, **p<0.01, *p<0.05)"""
    if p < 0.001:
        return '***'
    elif p < 0.01:
        return '**'
    elif p < 0.05:
        return '*'
    else:
        return ''


def export_to_excel(valid_df, output_path="analysis_results.xlsx"):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 创建说明工作表
        legend_df = pd.DataFrame({
            'Symbol': ['*', '**', '***'],
            'Meaning': ['p < 0.05', 'p < 0.01', 'p < 0.001']
        })
        legend_df.to_excel(writer, sheet_name='Legend', index=False)

        for sentiment in ['positive', 'negative']:
            subset = valid_df[valid_df['is_recommended'] == sentiment]

            # 整体检验
            contingency_all = pd.crosstab(subset['language'], subset['dominant_emotion'])
            chi2, p_overall, _, _ = chi2_contingency(contingency_all)
            overall_df = pd.DataFrame({
                'Test Type': ['Chi-square test'],
                'Chi-square': [f"{chi2:.3f}"],
                'p-value': [f"{p_overall:.5f}"],
                'Significance': [get_significance_stars(p_overall)]
            })

            # 详细对比：先收集 p 值列表
            detail_data = []
            raw_pvals = []
            emotions = contingency_all.columns.tolist()

            for emotion in emotions:
                contingency = pd.crosstab(
                    subset['language'],
                    subset['dominant_emotion'] == emotion
                ).reindex(index=['english', 'schinese'], columns=[True, False], fill_value=0)

                _, p = fisher_exact(contingency)
                raw_pvals.append(p)

            # Bonferroni 校正
            corrected_pvals = multipletests(raw_pvals, method='bonferroni')[1]

            # 再构造最终表格
            for i, emotion in enumerate(emotions):
                en_rate = subset[subset['language'] == 'english']['dominant_emotion'].value_counts(normalize=True).get(
                    emotion, 0) * 100
                cn_rate = subset[subset['language'] == 'schinese']['dominant_emotion'].value_counts(normalize=True).get(
                    emotion, 0) * 100

                detail_data.append([
                    emotion,
                    f"{en_rate:.1f}%",
                    f"{cn_rate:.1f}%",
                    f"{corrected_pvals[i]:.5f}",
                    get_significance_stars(corrected_pvals[i])
                ])

            detail_df = pd.DataFrame(detail_data,
                                     columns=['Emotion', 'English Rate', 'Chinese Rate',
                                              'Adjusted p-value', 'Significance'])

            # 写入
            sheet_name = f"{sentiment.capitalize()} Reviews"
            overall_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
            detail_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=5)

            # 调整列宽
            worksheet = writer.sheets[sheet_name]
            for col in ['A', 'B', 'C', 'D', 'E']:
                worksheet.column_dimensions[col].width = 18
            worksheet['A1'] = f"{sheet_name} Statistical Analysis"


def visualize(pos_data, neg_data):
    emotions = pos_data.index.tolist()
    x = np.arange(len(emotions))
    bar_width = 0.35

    fig = plt.figure(figsize=(20, 8))

    # 第一个子图（正面评价）
    ax1 = plt.subplot(1, 2, 1)
    plt.bar(x - bar_width / 2, pos_data['english'], width=bar_width, label='English', color='#66c2a5')
    plt.bar(x + bar_width / 2, pos_data['schinese'], width=bar_width, label='Chinese', color='#fc8d62')
    # 获取当前坐标轴范围
    xmin, xmax = ax1.get_xlim()
    ymin, ymax = ax1.get_ylim()
    # 在绘图区域右上角添加"a"标签
    plt.text(xmax * 0.99, ymax * 1.33, 'a', fontsize=14, fontweight='bold', ha='right', va='top')
    plt.xticks(x, emotions, rotation=45, ha='right')
    plt.ylabel('Percentage (%)')
    plt.ylim(0, 100)
    # 图例放在左边
    plt.legend(loc='upper left', bbox_to_anchor=(0, 1))

    # 第二个子图（负面评价）
    ax2 = plt.subplot(1, 2, 2)
    plt.bar(x - bar_width / 2, neg_data['english'], width=bar_width, label='English', color='#66c2a5')
    plt.bar(x + bar_width / 2, neg_data['schinese'], width=bar_width, label='Chinese', color='#fc8d62')
    # 获取当前坐标轴范围
    xmin, xmax = ax2.get_xlim()
    ymin, ymax = ax2.get_ylim()
    # 在绘图区域右上角添加"b"标签
    plt.text(xmax * 0.99, ymax * 1.18, 'b', fontsize=14, fontweight='bold', ha='right', va='top')
    plt.xticks(x, emotions, rotation=45, ha='right')
    plt.ylim(0, 100)
    # 图例放在左边
    plt.legend(loc='upper left', bbox_to_anchor=(0, 1))

    plt.tight_layout()

    # 保存图像为figure8.png（DPI=600）
    fig.savefig('figure8.png', dpi=600, bbox_inches='tight')
    plt.show()


if __name__ == "__main__":
    df = load_data("emotion_scores.xlsx")
    valid_df = df[df['is_recommended'].isin(['positive', 'negative'])]

    # 导出Excel结果
    export_to_excel(valid_df, "emotion_analysis.xlsx")
    print("分析结果已保存至 emotion_analysis.xlsx")

    # 可视化
    pos_percent, neg_percent = process_data(df)
    visualize(pos_percent, neg_percent)
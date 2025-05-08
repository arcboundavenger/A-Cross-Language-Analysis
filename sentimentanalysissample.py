import pandas as pd
from ollama import generate
import json
import os
from tqdm import tqdm
from collections import defaultdict

# 初始化情感标签及默认值
EMOTION_TYPES = ['Anger', 'Disgust', 'Anticipation', 'Fear',
                 'Joy', 'Sadness', 'Trust', 'Surprise']
DEFAULT_EMOTIONS = {k: 0.0 for k in EMOTION_TYPES}

EMOTION_SENTIMENT_MAPPING = {
    'Anger': 'negative',
    'Disgust': 'negative',
    'Anticipation': 'positive',
    'Fear': 'negative',
    'Joy': 'positive',
    'Sadness': 'negative',
    'Trust': 'positive',
    'Surprise': 'positive'
}


def init_excel_file(file_path):
    """初始化带八种情感字段的Excel文件"""
    columns = ['review_id', 'content', 'language', 'is_recommended',
               'sentiment', 'confidence', 'dominant_emotion'] + EMOTION_TYPES
    pd.DataFrame(columns=columns).to_excel(file_path, index=False)


def analyze_sentiment(text):
    """改进的提示词工程（强制二分类）"""
    emotion_definitions = """
    情感强度定义（0-1范围）：
    Anger（愤怒）: 表达攻击性/不满的程度（如：垃圾游戏→0.95）
    Disgust（厌恶）: 排斥/反感程度（如：恶心→0.9）
    Anticipation（期待）: 对未来的期望值（如：等更新→0.8） 
    Fear（恐惧）: 担忧/害怕程度（如：封号风险→0.7）
    Joy（快乐）: 积极愉悦程度（如：太好玩了→0.95）
    Sadness（悲伤）: 失落/难过程度（如：好友退游→0.85）
    Trust（信任）: 对产品/官方的认可度（如：官方良心→0.9）
    Surprise（惊讶）: 意外感受程度（如：没想到这么好→0.75）"""

    prompt = f"""{emotion_definitions}

    请分析以下评论的八维度情感强度，返回包含各情感数值的JSON：
    {{
      "sentiment": "positive/negative",  # 关键修改：删除中性选项
      "confidence": 0-1,
      "emotions": {{
        "Anger": 0.0-1.0,
        "Disgust": 0.0-1.0,
        "Anticipation": 0.0-1.0,
        "Fear": 0.0-1.0,
        "Joy": 0.0-1.0,
        "Sadness": 0.0-1.0,
        "Trust": 0.0-1.0,
        "Surprise": 0.0-1.0
      }}
    }}
    评论内容：{text}"""

    try:
        response = generate(
            model='deepseek-r1:8b',
            prompt=prompt,
            format='json',
            options={'temperature': 0.1}
        )
        result = json.loads(response['response'])

        # 结果校验
        for emo in EMOTION_TYPES:
            val = result['emotions'].get(emo, 0.0)
            result['emotions'][emo] = min(max(float(val), 0.0), 1.0)

        # 确定主导情感
        emotions = result['emotions']
        max_score = max(emotions.values())
        candidates = [k for k, v in emotions.items() if v == max_score]

        # 直接根据主导情感确定最终极性（参考网页7的二分类逻辑）
        dominant_emotion = candidates[0]
        if len(candidates) > 1:
            dominant_emotion = max(candidates, key=lambda x: EMOTION_SENTIMENT_MAPPING[x] == 'positive')

        # 强制覆盖为二分类结果（参考网页3的极性分类）
        final_sentiment = EMOTION_SENTIMENT_MAPPING[dominant_emotion]

        return {
            'sentiment': final_sentiment,  # 只返回positive/negative
            'confidence': result['confidence'],
            'emotions': result['emotions'],
            'dominant_emotion': dominant_emotion
        }

    except Exception as e:
        return {
            'sentiment': 'negative',  # 错误时默认负面
            'confidence': 0.0,
            'emotions': DEFAULT_EMOTIONS,
            'dominant_emotion': 'error'
        }

def process_game_reviews(file_path):
    """改进的主处理流程"""
    excel_file = pd.ExcelFile(file_path)
    init_excel_file('emotion_scores.xlsx')

    with pd.ExcelWriter('emotion_scores.xlsx', engine='openpyxl', mode='a') as writer:
        for sheet_name in tqdm(excel_file.sheet_names, desc="处理游戏"):
            if sheet_name == '0_Summary':
                continue

            df = excel_file.parse(sheet_name)

            # 新增：转换推荐字段为情感标签
            df['is_recommended'] = df['is_recommended'].map({
                True: 'positive',
                False: 'negative',
                'TRUE': 'positive',  # 兼容字符串类型
                'FALSE': 'negative'
            }).astype('category')

            # 情感分析
            tqdm.pandas(desc=f"情感分析 - {sheet_name}")
            results = df['content'].progress_apply(
                lambda x: pd.Series(analyze_sentiment(x))
            )

            # 构建结果数据
            emotion_df = pd.json_normalize(results['emotions'])
            final_df = pd.concat([
                df[['review_id', 'content', 'language', 'is_recommended']],
                results[['sentiment', 'confidence', 'dominant_emotion']],
                emotion_df
            ], axis=1)

            # 保存结果
            final_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

if __name__ == "__main__":
    process_game_reviews('steam_reviews_top50.xlsx')
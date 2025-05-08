import steamreviews
import pandas as pd
import time
import random
from tqdm import tqdm
import re

# 配置请求参数
BASE_PARAMS = {
    'filter': 'toprated',
    'language': 'all',
    'num_per_page': 1000,
    'json_query': {
        'strOrder': 'toprated',
        'bIgnoreUnvoted': False
    }
}

# 游戏ID列表
APP_IDS = [
    1623730, 1517290, 1551360, 1987080, 2050650, 1919590, 1196590, 1451940,
    529340, 1966900, 2344520, 860510, 1084600, 1515210, 668580, 1850570,
    1985810, 2239150, 1371980, 780310, 1338770, 2427700, 1388880, 2138710,
    1034140, 990630, 1475810, 1465460, 2163330, 1954200, 1601580, 2114740,
    1875830, 1335790, 1509510, 2144740, 973810, 1497440, 1328840
]
def clean_illegal_chars(s):
    if isinstance(s, str):
        return re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', s)
    return s

def fetch_reviews(app_id):
    """获取单个游戏的前50条点赞最多的评论"""
    time.sleep(random.uniform(3, 5))  # 小心别被封IP

    try:
        review_dict, _ = steamreviews.download_reviews_for_app_id(
            app_id=str(app_id),
            chosen_request_params=BASE_PARAMS
        )

        reviews = []
        for review_id, data in review_dict.get('reviews', {}).items():
            reviews.append({
                'review_id': review_id,
                'language': data.get('language', 'unknown'),
                'is_recommended': data.get('voted_up'),
                'votes_up': data.get('votes_up', 0),
                'votes_funny': data.get('votes_funny', 0),
                'weighted_score': data.get('weighted_vote_score', 0),
                'playtime_at_review': f"{data.get('author', {}).get('playtime_at_review', 0) / 60:.1f}h",
                'content': clean_illegal_chars(data.get('review', '')[:2000]),
                'created_at': pd.to_datetime(data.get('timestamp_created', 0), unit='s'),
                'steam_purchase': data.get('steam_purchase', False)
            })

        # ✅ 核心改动：按 votes_up 降序排序，取前50条
        reviews_sorted = sorted(reviews, key=lambda x: x['votes_up'], reverse=True)
        top_reviews = reviews_sorted[:50]

        return pd.DataFrame(top_reviews)

    except Exception as e:
        print(f"\n[Error] AppID {app_id}: {str(e)}")
        return None

# 写入Excel
with pd.ExcelWriter('steam_reviews_top50.xlsx', engine='openpyxl') as writer:
    success_count = 0
    for app_id in tqdm(APP_IDS, desc="Downloading Top Reviews"):
        df = fetch_reviews(app_id)
        if df is not None:
            df.to_excel(writer, sheet_name=str(app_id), index=False)
            success_count += 1
        time.sleep(random.uniform(1, 2))

    # 摘要Sheet
    summary_df = pd.DataFrame({
        'app_id': APP_IDS,
        'status': ['Success' if str(app_id) in writer.sheets else 'Failed' for app_id in APP_IDS],
        'reviews_saved': [50 if str(app_id) in writer.sheets else 0 for app_id in APP_IDS]
    })
    summary_df.to_excel(writer, sheet_name='0_Summary', index=False)

print(f"\n完成！成功抓取 {success_count} 个游戏Top 50评论，已保存为 steam_reviews_top50.xlsx")

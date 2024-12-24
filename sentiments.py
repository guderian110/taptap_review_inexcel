from snownlp import SnowNLP
from pathlib import Path
import pandas as pd


# this_dir = Path(__file__).resolve().parent
# csv_dir = this_dir / "review_content.csv"

# data = pd.read_csv(csv_dir)

def analyze_sentiment(text):
        if isinstance(text, list):
            # 如果是列表，分句并计算每个句子的情感得分
            sentiments = [SnowNLP(sentence).sentiments for sentence in text]
            return sum(sentiments) / len(sentiments) if sentiments else 0.5  # 返回平均值
        elif isinstance(text, str):
            # 如果是字符串，直接计算情感得分
            s = SnowNLP(text)
            return s.sentiments
# data['sentiment_score'] = data['comment'].apply(analyze_sentiment)

# data.to_csv('review_content_with_sentiment.csv', index=False, encoding='utf-8')

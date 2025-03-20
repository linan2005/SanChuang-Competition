import json
import jieba
import re
import numpy as np
import pandas as pd
from collections import defaultdict
from sklearn.feature_extraction.text import TfidfVectorizer


# 配置参数
INITIAL_KEYWORDS_PATH = "initial_keywords.json"  # 初始关键词文件路径
COMMENTS_PATH = "3.14.xlsx"                 # 评论数据文件路径
OUTPUT_PATH = "topic_keywords.txt"              # 输出文件路径
STOPWORDS_PATH = "停用词.txt"                # 中文停用词文件路径

# 加载初始关键词
with open(INITIAL_KEYWORDS_PATH, 'r', encoding='utf-8') as f:
    topic_keywords = json.load(f)

# 加载评论数据
df = pd.read_excel(COMMENTS_PATH)
comments = df['评价内容'].tolist()

# 加载停用词
with open(STOPWORDS_PATH, 'r', encoding='utf-8') as f:
    stopwords = set([line.strip() for line in f])

# ----------------------
# 1. 数据预处理：清洗与分词
# ----------------------
def preprocess(text):
    # 去除结构化前缀（如"画面品质："）
    text = re.sub(r'\w+：', '', text)
    # 分词并过滤
    words = jieba.lcut(text)
    return [word for word in words if len(word) > 1 and word not in stopwords]

processed_comments = [preprocess(comment) for comment in comments]

# ----------------------
# 2. 基于TF-IDF提取高频词
# ----------------------
# 将分词结果拼接为字符串
texts = [' '.join(words) for words in processed_comments]

# 计算TF-IDF
tfidf = TfidfVectorizer(max_features=100)
tfidf_matrix = tfidf.fit_transform(texts)
feature_names = tfidf.get_feature_names_out()
tfidf_scores = np.asarray(tfidf_matrix.sum(axis=0)).ravel()

# 按TF-IDF得分排序
tfidf_keywords = {
    word: score for word, score in zip(feature_names, tfidf_scores)
}

# ----------------------
# 3. 动态扩展关键词（基于规则匹配）
# ----------------------
enhanced_keywords = defaultdict(list)

for topic, keywords in topic_keywords.items():
    # 添加初始关键词
    enhanced_keywords[topic].extend(keywords)
    
    # 基于TF-IDF扩展
    topic_tfidf = {}
    for word in tfidf_keywords:
        if any(kw in word for kw in keywords):  # 粗略匹配
            topic_tfidf[word] = tfidf_keywords[word]
    # 取TF-IDF最高的前5个词
    top_tfidf = sorted(topic_tfidf.items(), key=lambda x: -x[1])[:5]
    enhanced_keywords[topic].extend([word for word, _ in top_tfidf])

# 去重并过滤
final_keywords = {}
for topic, words in enhanced_keywords.items():
    unique_words = list(set([w for w in words if w in tfidf_keywords]))
    final_keywords[topic] = unique_words[:5]  # 每个主题最多保留10个词

# ----------------------
# 4. 保存结果
# ----------------------
with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    for topic, words in final_keywords.items():
        line = f"{topic}：" + "、".join(words) + "\n"
        f.write(line)

print(f"关键词库已生成至 {OUTPUT_PATH}")



# 假设 final_keywords 是模型提取的关键词库
# tfidf_keywords 是 TF-IDF 提取的高频词

# 将 final_keywords 展平为一个集合
model_keywords = set()
for topic, words in final_keywords.items():
    model_keywords.update(words)

# 计算覆盖率
tfidf_words = set(tfidf_keywords.keys())
coverage = len(model_keywords.intersection(tfidf_words)) / len(tfidf_words) * 100

print(f"模型提取的关键词覆盖率：{coverage:.2f}%")
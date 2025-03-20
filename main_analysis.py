import json
import pandas as pd
import re
from snownlp import SnowNLP

# 配置参数
KEYWORDS_PATH = "updated_keywords.json"  # 关键词文件路径
COMMENTS_PATH = "3.14.xlsx"             # 评论数据文件路径
OUTPUT_PATH = "binary_variables.xlsx"   # 输出文件路径

# 加载关键词
with open(KEYWORDS_PATH, "r", encoding="utf-8") as f:
    topic_keywords = json.load(f)

# 加载评论数据
df = pd.read_excel(COMMENTS_PATH)

# 统计总评论数
total_comments = len(df)

# 删除“此用户未填写评价内容”的评论
df = df[df["评价内容"] != "此用户未填写评价内容"]

# 统计筛选后的有效评论数
valid_comments = len(df)

# 计算有效数据占比
valid_ratio = valid_comments / total_comments * 100

# 输出统计信息
print(f"总评论数: {total_comments}")
print(f"有效评论数: {valid_comments}")
print(f"有效数据占比: {valid_ratio:.2f}%")

# ----------------------
# 情感分析
# ----------------------
def get_sentiment(text):
    """
    使用 SnowNLP 判断情感倾向。
    - 返回值：情感得分（0~1），大于 0.5 为正向，否则为负向。
    """
    return SnowNLP(text).sentiments

# 添加情感评分列
df["sentiment_score"] = df["评价内容"].apply(get_sentiment)

# 筛选情感评分正向的评论（情感得分 > 0.5）
df_positive = df[df["sentiment_score"] > 0.5]

# 统计正向评论数
positive_comments = len(df_positive)

# 计算正向评论占比
positive_ratio = positive_comments / valid_comments * 100

# 输出正向评论统计信息
print(f"正向评论数: {positive_comments}")
print(f"正向评论占比: {positive_ratio:.2f}%")

# 提取正向评论列表
comments = df_positive["评价内容"].tolist()

# ----------------------
# 关键词匹配
# ----------------------
def match_keywords(comment, keywords):
    """
    匹配评论中的关键词。
    - 返回值：匹配到的主题列表。
    """
    matched_topics = []
    for topic, words in keywords.items():
        if any(word in comment for word in words):
            matched_topics.append(topic)
    return matched_topics

# ----------------------
# 生成 1/0 变量
# ----------------------
# 初始化结果表
results = []

for idx, row in df_positive.iterrows():
    # 关键词匹配
    matched_topics = match_keywords(row["评价内容"], topic_keywords)

    # 生成 1/0 变量
    result_row = {
        "会员": row.get("会员", ""),  # 将会员列写入
        "时间": row.get("时间", ""),  # 将时间列写入
        "comment": row["评价内容"],
        "sentiment_score": row["sentiment_score"]  # 添加 sentiment_score 列
    }
    for topic in topic_keywords.keys():
        result_row[topic] = 1 if topic in matched_topics else 0
    results.append(result_row)

# 将结果保存为 DataFrame
df_results = pd.DataFrame(results)

# ----------------------
# 写入 Excel
# ----------------------
with pd.ExcelWriter(OUTPUT_PATH) as writer:
    df_results.to_excel(writer, index=False, sheet_name="Binary Variables")
    pd.DataFrame(topic_keywords.items(), columns=["Topic", "Keywords"]).to_excel(
        writer, index=False, sheet_name="Keywords"
    )
    # 保存统计信息
    stats = {
        "总评论数": total_comments,
        "有效评论数": valid_comments,
        "有效数据占比": valid_ratio,
        "正向评论数": positive_comments,
        "正向评论占比": positive_ratio
    }
    pd.DataFrame([stats]).to_excel(writer, index=False, sheet_name="Statistics")

print(f"1/0 变量结果已保存至 {OUTPUT_PATH}")
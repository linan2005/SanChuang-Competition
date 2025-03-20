import pandas as pd
import statsmodels.api as sm

# 加载数据
df = pd.read_excel("binary_variables.xlsx", sheet_name="Binary Variables")

# 将时间列转换为日期格式
df["时间"] = pd.to_datetime(df["时间"])

# 按天统计评论量（作为销量）
daily_sales = df.groupby(df["时间"].dt.date).size().reset_index(name="销量")

# 按天统计关键词出现次数
daily_keywords = df.groupby(df["时间"].dt.date)[list(df.columns[4:])].sum().reset_index()

# 按天统计情感得分均值（作为总体满意度）
daily_sentiment = df.groupby(df["时间"].dt.date)["sentiment_score"].mean().reset_index(name="总体满意度")

# 合并销量、关键词和满意度数据
df_regression = pd.merge(daily_sales, daily_keywords, on="时间")
df_regression = pd.merge(df_regression, daily_sentiment, on="时间")

# ----------------------
# 回归分析：关键词对销量的影响
# ----------------------
# 自变量：关键词列
X = df_regression[list(df.columns[4:])]
# 因变量：销量
y = df_regression["销量"]

# 添加常数项
X = sm.add_constant(X)

# 构建线性回归模型
model = sm.OLS(y, X).fit()

# 输出回归结果
print("回归分析结果（关键词对销量的影响）：")
print(model.summary())

# ----------------------
# 回归分析：关键词对总体满意度的影响
# ----------------------
# 自变量：关键词列
X_satisfaction = df_regression[list(df.columns[4:])]
# 因变量：总体满意度
y_satisfaction = df_regression["总体满意度"]

# 添加常数项
X_satisfaction = sm.add_constant(X_satisfaction)

# 构建线性回归模型
model_satisfaction = sm.OLS(y_satisfaction, X_satisfaction).fit()

# 输出回归结果
print("\n回归分析结果（关键词对总体满意度的影响）：")
print(model_satisfaction.summary())
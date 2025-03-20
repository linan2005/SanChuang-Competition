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

# 合并销量和关键词数据
df_regression = pd.merge(daily_sales, daily_keywords, on="时间")

# 确保所有列都是数值型（排除非数值型列，如“时间”）
df_regression_numeric = df_regression.drop(columns=["时间"])

# ----------------------
# 多元线性回归分析
# ----------------------
# 自变量：关键词列
X = df_regression_numeric.drop(columns=["销量"])
# 因变量：销量
y = df_regression_numeric["销量"]

# 添加常数项（截距）
X = sm.add_constant(X)

# 构建多元线性回归模型
model = sm.OLS(y, X).fit()

# 输出回归结果
print("多元线性回归分析结果：")
print(model.summary())

# 输出回归方程
equation = "销量 = {:.2f}".format(model.params['const'])
for feature in X.columns[1:]:
    equation += " + {:.2f} * {}".format(model.params[feature], feature)
print("\n回归方程：")
print(equation) 
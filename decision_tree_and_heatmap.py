import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.tree import DecisionTreeRegressor, plot_tree
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error, r2_score

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置字体为 SimHei（黑体）
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
sns.set(font='SimHei')  # 设置 Seaborn 的字体

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
# 决策树分析
# ----------------------
# 自变量：关键词列
X = df_regression_numeric.drop(columns=["销量"])
# 因变量：销量
y = df_regression_numeric["销量"]

# 划分训练集和测试集
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# 构建决策树模型
tree_model = DecisionTreeRegressor(max_depth=3, random_state=42)
tree_model.fit(X_train, y_train)

# 预测
y_pred = tree_model.predict(X_test)

# 评估模型
mse = mean_squared_error(y_test, y_pred)
r2 = r2_score(y_test, y_pred)
print(f"决策树模型评估：")
print(f"均方误差 (MSE): {mse:.2f}")
print(f"R² 分数: {r2:.2f}")

# 可视化决策树
plt.figure(figsize=(20, 10),dpi=300)
plot_tree(tree_model, feature_names=X.columns, filled=True, rounded=True)
plt.title("决策树模型")
plt.show()

# ----------------------
# 热力图：关键词与销量的相关性
# ----------------------
# 计算相关性矩阵
correlation_matrix = df_regression_numeric.corr()

# 绘制热力图
plt.figure(figsize=(10, 8))
sns.heatmap(correlation_matrix, annot=True, cmap="coolwarm", fmt=".2f")
plt.title("关键词与销量的相关性热力图")
plt.show()
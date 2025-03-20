from sklearn.datasets import load_iris
from sklearn.tree import DecisionTreeClassifier, export_graphviz
import graphviz

# 加载数据集
iris = load_iris()
X = iris.data
y = iris.target

# 创建决策树模型
clf = DecisionTreeClassifier()
clf.fit(X, y)

# 导出决策树为Graphviz格式
dot_data = export_graphviz(clf, out_file=None, 
                           feature_names=iris.feature_names,  
                           class_names=iris.target_names,  
                           filled=True, rounded=True,  
                           special_characters=True)  

# 使用graphviz生成图像
graph = graphviz.Source(dot_data)
# 设置分辨率，例如设置为300dpi
graph.attr(dpi='300')
# 保存为PDF格式，PDF通常可以保持高质量
graph.render("iris_tree", format='pdf', cleanup=True)
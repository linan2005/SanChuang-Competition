import pandas as pd
import matplotlib.pyplot as plt
import re

# 设置图片清晰度
plt.rcParams['figure.dpi'] = 300

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 读取 Excel 文件
excel_file = pd.ExcelFile('jdcomment.xlsx')

# 获取所有表名
sheet_names = excel_file.sheet_names

# 遍历每个 sheet
for sheet_name in sheet_names:
    print(f'正在分析表 {sheet_name}...')
    # 获取当前 sheet 的数据
    data_jingdong = excel_file.parse(sheet_name)
        # 分解商品字段
    def split_product_info(product):
        # 处理换行符
        product = product.replace('\n', ' ')
        # 定义日期的正则表达式模式
        date_pattern = r'\d{4}-\d{2}-\d{2}'
        # 查找日期
        match = re.search(date_pattern, product)
        if match:
            date = match.group()
            # 根据日期位置分割配置信息和地理位置
            date_index = product.index(date)
            config_info = product[:date_index].strip()
            location = product[date_index + len(date):].strip()
        else:
            # 如果未找到日期，给出默认值
            config_info = product
            date = '未知日期'
            location = '未知位置'
        return config_info, date, location

    data_jingdong[['配置信息', '日期', '地理位置']] = data_jingdong['商品属性'].apply(lambda x: pd.Series(split_product_info(x)))





    # 按月统计京东店铺的销售数量；按商品属性统计商品的销售情况；按评分统计评价人数并计算商品在京东的评分
    level = data_jingdong['级别']
    star = data_jingdong['评价星级']
    eva_time = data_jingdong['时间']
    style = data_jingdong['配置信息']

    # 各月份的用户评论条数
    data_jingdong['时间'] = pd.to_datetime(data_jingdong['时间'])
    data_jingdong['月份'] = data_jingdong['时间'].dt.month
    monthly_comments = data_jingdong['月份'].value_counts().sort_index()
    print(f'表 {sheet_name} 各月份的用户评论条数：')
    print(monthly_comments)


    # 绘制各月份用户评论条数折线图
    plt.figure(figsize=(10, 6))
    monthly_comments.plot(kind='line', marker='o')
    plt.title(f'表 {sheet_name} 各月份的用户评论条数')
    plt.xlabel('月份')
    plt.ylabel('评论条数')
    plt.xticks(rotation=0)
    plt.show()

    # 各属性购买数量
    group_data = data_jingdong.groupby('配置信息')['会员'].count()
    print(f'表 {sheet_name} 各属性购买数量：')
    print(group_data)

   

    # # 绘制各属性购买数量饼图
    # plt.figure(figsize=(8, 8))
    # group_data.plot(kind='pie', autopct='%1.1f%%', startangle=90)
    # plt.title(f'表 {sheet_name} 各属性购买数量占比')
    # plt.ylabel('')  # 隐藏 y 轴标签
    # plt.show()

    # 绘制各属性购买数量饼图
    plt.figure(figsize=(8, 8))
    # 绘制饼图，设置 labeldistance 为 None 避免标签与图例冲突
    patches, texts, autotexts = plt.pie(group_data, autopct='%1.1f%%', startangle=90)    
    plt.title(f'表 {sheet_name} 各属性购买数量占比')
    plt.ylabel('')  # 隐藏 y 轴标签
    # 添加图例
    plt.legend(patches, group_data.index, loc="best")
    plt.show()


    # 绘制柱状图
    plt.figure(figsize=(10, 6))
    ax = group_data.plot(kind='bar')

    # 添加数据标签
    for p in ax.patches:
        ax.annotate(str(p.get_height()), (p.get_x() + p.get_width() / 2., p.get_height()),
                    ha='center', va='center', xytext=(0, 5), textcoords='offset points')

    # 设置图表标题和坐标轴标签
    plt.title(f'表 {sheet_name} 各商品属性的购买数量')
    plt.xlabel('商品属性')
    plt.ylabel('购买数量')

    # 旋转 x 轴标签以避免重叠
    plt.xticks(rotation=45)

    # 显示图表
    plt.show()


    # 提取最高的商品属性和销售数量
    pname = group_data.idxmax()
    pnum = group_data.max()
    print(f'表 {sheet_name} 销量最高的商品是<{pname}>,一共销售了{pnum}份')

    # 按照会员和非会员分别统计人数
    plus_num = data_jingdong.groupby('级别')['会员'].count()[0]
    print(f'表 {sheet_name} PLUS会员占比{plus_num/len(data_jingdong)*100:.2f}%')

    # 取出 star5
    data_jingdong['score'] = [int(i[-1]) for i in data_jingdong['评价星级']]
    ave_score = data_jingdong['score'].mean()
    print(f'表 {sheet_name} 这件商品的平均评分是{ave_score:.2f}')

    print('-' * 50)
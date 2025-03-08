import pandas as pd
import matplotlib.pyplot as plt
from pyecharts import options as opts
from pyecharts.charts import Map
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

    # 统计地理位置出现的次数并重置索引
    location_counts = data_jingdong['地理位置'].value_counts().reset_index()
    location_counts.columns = ['地理位置', '数量']

    # 定义需要排除的地区列表
    exclude_areas = ['海外']

    # 筛选掉海外和港澳地区的数据
    location_counts = location_counts[~location_counts['地理位置'].isin(exclude_areas)]

    # 定义自治区，直辖市，台湾字典，用于补全名称
    autonomous_regions = {
        '内蒙古': '内蒙古自治区',
        '新疆': '新疆维吾尔自治区',
        '广西': '广西壮族自治区',
        '宁夏': '宁夏回族自治区',
        '西藏': '西藏自治区',
        '中国台湾': '台湾省',
        '北京': '北京市',
        '天津': '天津市',
        '上海': '上海市',
        '重庆': '重庆市',
        '港澳': '香港特别行政区',
    }

    # 定义一个函数来修改地理位置名称
    def modify_location(location):
        # 如果是自治区，补全名称
        if location in autonomous_regions:
            return autonomous_regions[location]
        # 其他情况添加 '省' 字
        return location + '省'

    # 应用函数修改地理位置名称
    location_counts['地理位置'] = location_counts['地理位置'].apply(modify_location)
    maxValue = location_counts['数量'].max()  # 店铺最大值
    location = list(zip(location_counts['地理位置'], location_counts['数量']))
    print(location)

    # 使用 pyecharts 绘制地图
    c= (
        Map()
        .add("销量", location, "china")
        .set_global_opts(
            title_opts=opts.TitleOpts(title=f"{sheet_name} 各地理位置销量分布"),
            visualmap_opts=opts.VisualMapOpts(max_=int(maxValue)),
        )
        .render(f"{sheet_name}_location_map.html")
    )



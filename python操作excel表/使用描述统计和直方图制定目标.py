import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw

df = pd.read_excel('F:\\测试用表格\\描述统计.xlsx')
df.columns = ['序号', '员工姓名', '月销售额']
df = df.drop(columns=['序号', '员工姓名'])  # 删除“序号”列和“员工姓名”列
df_describe = df.astype('float').describe()  # 计算数据的个数，平均值，最大值和最小值等描述统计数据
df_cut = pd.cut(df['月销售额'], bins=7, precision=2)  # 将“月销售额”列的数据分成7个均等的区间
cut_count = df['月销售额'].groupby(df_cut).count()
df_all = pd.DataFrame()  # 创建一个空的DataFrame用于绘总数据
print(cut_count)
df_all['计数'] = cut_count  # 将月销售额的区间及区间的人数写入前面创建的DataFrame中
df_all_new = df_all.reset_index()  # 将索引重置为数字序号
df_all_new['月销售额'] = df_all_new['月销售额'].apply(lambda x: str(x))  # 将’月销售额‘列的数据转换为字符串类型
fig = plt.figure()  # 创建绘图窗口
plt.rcParams['font.sans-serif'] = ['SimHei']  # 解决中文乱码问题
n, bins, patches = plt.hist(df['月销售额'], bins=7, edgecolor="black", linewidth=0.5)  # 使用月销售额列的数据绘制直方图
plt.xticks(bins)  # 将直方图x轴的刻度标签设置为个区间的端点值
plt.title('月销售额频率分析')  # 设置直方图的图表标题
plt.xlabel('月销售额')  # 设置直方图x轴的标题
plt.ylabel("频数")  # 设置直方图y轴的标题
app = xw.App(visible=False)
workbook = app.books.open('F:\\测试用表格\\描述统计.xlsx')
worksheet = workbook.sheets['业务员销售额统计表']
worksheet.range('E2').value = df_describe  # 将计算出的个数，平均值，最大值和最小值等数据写入工作表
worksheet.range('H2').value = df_all_new  # 将月销售额的区间及区间的人数写入工作表
worksheet.pictures.add(fig, name='图片1', update=True, left=400, top=200)  # 将绘制的直方图转换为图片并写入工作表
worksheet.autofit()
workbook.save('F:\\测试用表格\\描述统计1.xlsx')  # 另存工作簿
workbook.close()
app.quit()

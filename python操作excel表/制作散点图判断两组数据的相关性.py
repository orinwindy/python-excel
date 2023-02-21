import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\汽车速度和刹车距离表.xlsx')
figure=plt.figure()
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False
x=df['汽车速度（km/h）']#指定数据中的“汽车速度（km/h）”列为x坐标的值
y=df['刹车距离（m）']#指定数据中的“刹车距离（m）”列为y坐标的值
plt.scatter(x,y,s=400,color='red',marker='o',edgecolors='black')#制作散点图,s为点面积
plt.xlabel('汽车速度（km/h）',fontdict={'family':'Microsoft YaHei','color':'black','size':20},labelpad=20)#添加并设置x轴标题
plt.ylabel('刹车距离（m）',fontdict={'family':'Microsoft YaHei','color':'black','size':20},labelpad=20)
plt.title('汽车速度与刹车距离关系图',fontdict={'family':'Microsoft YaHei','color':'black','size':30},loc='center')#添加并设置图表标题
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\汽车速度和刹车距离表.xlsx')#打开要插入图表的工作簿
worksheet=workbook.sheets[0]
worksheet.pictures.add(figure,name='图片1',update=True,left=200)#在工作表中插入制作的散点图
workbook.save('F:\\测试用表格\\散点图.xlsx')
workbook.close()
app.quit()
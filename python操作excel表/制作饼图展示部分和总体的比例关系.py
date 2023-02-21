import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\饼图.xlsx')
figure=plt.figure()
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False
x=df['产品名称']
y=df['销售额']
plt.pie(y,labels=x,labeldistance=1.1,autopct='%.2f%%',pctdistance=0.8,startangle=90,radius=1.0,explode=[0,0,0,0,0,0.3,0])#制作饼图并分离饼图块,第10行代码中用pie()函数根据这
# 些数据制作饼图，并将饼图中的第6个饼图块分离出来
plt.title(label='产品销售额占比图',fontdict={'color':'black','size':30},loc='center')#添加并设置图表标题
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\饼图.xlsx')
worksheet=workbook.sheets[0]
worksheet.pictures.add(figure,name='图片1',update=True,left=200)
workbook.save()
workbook.close()
app.quit()
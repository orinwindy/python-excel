import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\销售业绩表.xlsx')
figure=plt.figure()
plt.rcParams['font.sans-serif']=["SimHei"]#为图表中的中文文本设置默认字体，以避免中文显示乱码问题
plt.rcParams['axes.unicode_minus']=False#解决坐标值为负数时无法正常显示负号的问题
x=df['月份']
y=df['销售额']
plt.scatter(x,y,s=500,color='red',marker='*')
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\销售业绩表.xlsx')
worksheet=workbook.sheets['销售业绩']
worksheet.pictures.add(figure,left=500)
workbook.save()
workbook.close()
app.quit()
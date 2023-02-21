import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
df = pd.read_excel('F:\\测试用表格\\方差分析.xlsx')
df=df[['A型号','B型号','C型号','D型号','E型号']]
figure=plt.figure()#创建绘图窗口
plt.rcParams['font.sans-serif']=['SimHei']#解决中文乱码问题
df.boxplot(grid=False)#绘制箱型图并删除网格线
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\方差分析.xlsx')
worksheet=workbook.sheets['单因素方差分析']
worksheet.pictures.add(figure,name='图片1',update=True,left=500,top=10)#将绘制的箱型图插入工作表
workbook.save('F:\\测试用表格\\箱形图.xlsx')
workbook.close()
app.quit()
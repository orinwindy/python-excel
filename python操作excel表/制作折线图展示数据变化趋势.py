import pandas as pd
import matplotlib.pyplot as plt
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\月销售表.xlsx')
figure=plt.figure()#创建一个绘图窗口
plt.rcParams['font.sans-serif']=['SimHei']#解决汉字乱码问题
plt.rcParams['axes.unicode_minus']=False#解决坐标值为负数时无法正常显示负号的问题
x=df['月份']
y=df['销售额']
plt.plot(x,y,color='red',linewidth=3,linestyle='solid')
plt.title(label='月销售额趋势图',fontdict={'color':'black','size':30},loc='center')
for a,b in zip(x,y):
    plt.text(a,b+0.2,(a,'%.0f'%b),ha='center',va='bottom',fontsize=10)#添加并设置数据标签
plt.axis('off')#隐藏坐标轴
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\月销售表.xlsx')
worksheet=workbook.sheets['Sheet1']
worksheet.pictures.add(figure,name='图片1',update=True,left=200)#在工作表中插入制作的折线图
workbook.save('F:\\测试用表格\\折线图.xlsx')
workbook.close()
app.quit()
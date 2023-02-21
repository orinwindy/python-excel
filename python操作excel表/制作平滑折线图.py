import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy import interpolate
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\月销售表.xlsx')
figure=plt.figure()
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False
x=df['月份']
y=df['销售额']
xnew=np.arange(1,12,0.1)
func=interpolate.interp1d(x,y,kind='cubic')#插值模块
ynew=func(xnew)
plt.plot(xnew,ynew,color='red',linewidth=3,linestyle='solid')#制作平滑折线图
plt.title(label='月销售额趋势图',fontdict={'color':'black','size':30},loc='center')
plt.xlabel('月份',fontdict={'family':'SimSun','color':'black','size':20},labelpad=20)
plt.ylabel('销售额',fontdict={"family":'SimSun','color':'black','size':20},labelpad=20)
plt.xlim(0,12)#设置图表x轴的取值范围
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\月销售表.xlsx')
worksheet=workbook.sheets['Sheet1']
worksheet.pictures.add(figure,name='图片1',update=True,left=200)
workbook.save('F:\\测试用表格\\平滑折线图.xlsx')
workbook.close()
app.quit()

import pandas as pd
import matplotlib.pyplot as plt
df=pd.read_excel('F:\\测试用表格\\销售业绩表.xlsx')
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False#解决坐标值为负数时无法正常显示负号的问题
x=df['月份']
y=df['销售额']
plt.bar(x,y,color='black')
plt.title(label="各月销售额对比图",fontdict={'family':'KaiTi','color':'red','size':30},loc='left')
plt.xlabel('月份',fontdict={'family':'SimSun','color':'black','size':15},labelpad=5)#添加并设置x轴标题，labelpad:图表标题到坐标轴的距离
plt.ylabel('销售额',fontdict={'family':'SimSun','color':'black','size':15},labelpad=5)#添加并设置y轴标题
plt.show()
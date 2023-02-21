import pandas as pd
import matplotlib.pyplot as plt
df=pd.read_excel('F:\\测试用表格\\销售业绩表.xlsx')
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False#解决坐标值为负数时无法正常显示负号的问题
x=df['月份']
y=df['销售额']
plt.bar(x,y,label='销售额')
plt.legend(loc='upper left',fontsize=20)
plt.show()
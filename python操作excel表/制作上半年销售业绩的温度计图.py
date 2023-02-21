import pandas as pd
import matplotlib.pyplot as plt
df=pd.read_excel('F:\\测试用表格\\温度计图.xlsx')
sum=0
for i in range(6):
    sum=df['销售业绩（万元）'][i]+sum
goal=df['销售业绩（万元）'][13]
percentage=sum/goal# 计算全年的实际销售业绩占目标销售业绩的百分比
plt.bar(1,1,color='yellow')# 制作柱形图展示全年的目标销售业绩，设置填充颜色为黄色
plt.bar(1,percentage,color='cyan')# 制作柱形图展示全年的实际销售业绩，设置填充颜色为青色
plt.xlim(0,2)
plt.ylim(0,1.2)
plt.text(1,percentage-0.01,percentage,ha='center',va='top',fontdict={'color':'black','size':20})
plt.show()
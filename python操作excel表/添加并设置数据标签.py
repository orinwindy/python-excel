import pandas as pd
import matplotlib.pyplot as plt
df=pd.read_excel('F:\\测试用表格\\销售业绩表.xlsx')
plt.rcParams['font.sans-serif']=['SimHei']
plt.rcParams['axes.unicode_minus']=False
x=df['月份']
y=df["销售额"]
plt.plot(x,y,color='red',linewidth=3,linestyle='solid')
for a,b in zip(x,y):
    plt.text(a,b,b,fontdict={'family':'KaiTi','color':'red','size':20})
plt.show()
# text()函数为折线图添加数据标签，因为这个函数只是为折线图上的
# 某一个数据点添加数据标签，所以还需要配合使用第9行代码中的for语句为整个折线图的
# 所有数据点添加数据标签。
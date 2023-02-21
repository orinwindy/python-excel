import pandas as pd
from sklearn import linear_model
df = pd.read_excel('F:\\测试用表格\\回归分析.xlsx',header=None)
df=df[2:]#删除前两行数据
df.columns=['月份','电视台广告费','视频门户广告费','汽车当月销售额']#重命名数据列
x=df[['视频门户广告费','电视台广告费']]#获取“视频门户广告费”列和“电视台广告费”列的数据作为自变量
y=df['汽车当月销售额']#获取“汽车当月销售额”列的数据作为因变量
model=linear_model.LinearRegression()#创建一个线性回归模型
model.fit(x,y)#用自变量和因变量数据对线性回归模型进行训练，拟合出线性回归方程
R2=model.score(x,y)
print(R2)#输入R平方的值
#R2值的取值范围为0～1，越接近1，说明方程的拟合程度越高。这里计算出的R2值比
# 较接近1，说明方程的拟合程度较高，可以用此方程来进行预测。
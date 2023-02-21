import pandas as pd
from sklearn import linear_model
df=pd.read_excel('F:\\测试用表格\\回归分析.xlsx',header=None)
df=df[2:]#删除前两行数据
df.columns=['月份','电视台广告费','视频门户广告费','汽车当月销售额']
x=df[['视频门户广告费','电视台广告费']]
y=df['汽车当月销售额']
model=linear_model.LinearRegression()
model.fit(x,y)
coef=model.coef_#获取自变量的系数
model_intercept=model.intercept_#获取截距
result='y={}*x1+{}*x2{}'.format(coef[0],coef[1],model_intercept)#获取线性回归方程
print('线性回归方差为:','\n',result)#输出线性回归方程
a=30#设置视频门户广告费
b=20#设置电视台广告费
y=coef[0]*a+coef[1]*b+model_intercept#根据线性回归方程计算汽车销售额
print(y)#输出计算出的汽车销售额
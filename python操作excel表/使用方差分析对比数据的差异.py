import pandas as pd
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm  # 导入statsmodels.stats.anova模块中的anova_lm()函数
import xlwings as xw
df=pd.read_excel('F:\\测试用表格\\方差分析.xlsx')
df = df[['A型号','B型号','C型号','D型号','E型号']]#选取“A型号”’B型号‘“C型号”，“D型号”“E型号”列的数据用于分析
df_melt=df.melt()#将列名转换为列数据,重构DataFrame
df_melt.columns=['Treat','Value']#重命名列
df_describe=pd.DataFrame()#创建空DataFrame用于汇总数据
df_describe['A型号']=df['A型号'].describe()#计算“A型号”轮胎刹车距离的平均值，最大值，最小值等
df_describe['B型号']=df['B型号'].describe()
df_describe['C型号']=df['C型号'].describe()
df_describe['D型号']=df['D型号'].describe()
df_describe['E型号']=df['E型号'].describe()
model=ols('Value~C(Treat)',data=df_melt).fit()#对样本数据进行最小二乘线性拟合计算,注意是~不是-
anova_table=anova_lm(model,typ=3)#对样本数据进行方差分析
app=xw.App(visible=False)
workbook=app.books.open('F:\\测试用表格\\方差分析.xlsx')
worksheet=workbook.sheets['单因素方差分析']#选中工作表“单因素方差分析”
worksheet.range('H2').value=df_describe.T#将计算出的平均值，最大值和最小值等数据转置行列并写入工作表
worksheet.range('H14').value='方差分析'#在工作表中写入‘方差分析’
worksheet.range('H15').value=anova_table#将方差分析的结果写入工作表
workbook.save()
workbook.close()
app.quit()

# 我们需要关心单元格L17中的数值，它相当于用
# Excel的单因素方差分析功能计算出的P-value，代表观测到的显著性水平。通常情况下，
# 该值≤0.01表示有极显著的差异，该值在0.01～0.05之间表示有显著差异，该值≥0.05表示
# 没有显著差异。这里的P-value为0.00674≤0.01，说明5种型号轮胎的平均刹车距离有极显
# 著的差异
import pandas as pd
df = pd.read_excel('相关性分析.xlsx',index_col='代理商编号')#index_col索引列
result=df.corr()["年销售额(万元)"]#计算年销售额与其他变量之间的皮埃逊相关系数
print(result)
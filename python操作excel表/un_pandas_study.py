import pandas as pd

# s = pd.Series(['丁一','王二','张三'])
# print(s)

# a = pd.DataFrame([[1, 2], [3, 4], [5, 6]], columns=['data', 'score'],index=['A','B','C'])
# print(a)

# 通过列表创建DateFrame结构
# a=[1,2, 3]
# b = [4, 5, 6]
# date = pd.DataFrame()
# date['a'] = a
# date['b'] = b
# print(date)

# 通过字典创建DateFrame,默认以键名为列索引
# b = pd.DataFrame({'a': [1, 3, 5], 'b': [2, 4, 6]}, index=['x', 'y', 'z'])
# print(b)
# 也可以键名为行索引
# b=pd.DataFrame.from_dict({'a': [1, 3, 5], 'b': [2, 4, 6]},orient='index',columns=['好','号','浩'])
# print(b)

# 通过二维数组创建DataFrame
# import numpy as np
#
# a = np.arange(12).reshape(3, 4)
# b = pd.DataFrame(a, index=[1, 2, 3], columns=["A", 'B', 'C', "D"])
# print(b)

# DataFrame索引的修改
# a = pd.DataFrame([[1, 2], [3, 4], [5, 6]], columns=['date', 'score'], index=['A', 'B', 'C'])
# # print(a)
# a.index.name='公司'
# print(a)
# 注意这里需要重新赋值给a，因为替换后会新建一个DataFrame，
# a = a.rename(index={'A': '万科','B': '阿里','C': '百度'}, columns={'date': '日期','score': '分数'})
# print(a)
# 如果不想重新赋值可以用inplace参数一步到位
# a.rename(index={'A': '万科','B': '阿里','C': '百度'}, columns={'date': '日期','score': '分数'},inplace=True)
# print(a)
# 重置行索引，即把当前行索引变为一列
# a.reset_index(inplace=True)
# print(a)
# # 设置某列为行索引
# a.set_index('date',inplace=True)
# print(a)

data=pd.read_excel('income.xlsx',sheet_name=0)
print(data)


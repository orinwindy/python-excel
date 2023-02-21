import pandas as pd

data = pd.read_excel("name.xlsx")#打开表格文件
data['列1'] = data['列2'].apply(lambda x: x.split('/')[0].strip())#把列2的数据用'/'分隔开,取第一段并去空格存到'列1'中（若列1不存在则新建列1）
writer = pd.ExcelWriter('temp.xlsx')#存储对象与存储表格文件名
for i in data['检索列'].unique():  # 去重（例如年份，国家）
    data[data['条件'] == i].to_excel(writer, sheet_name=i)#找到符合条件的数据存入存储对象中。（相同的年份在一张表上，不同的年份在不同表上）
writer.close()


print(data[data['t'].str.contains('包含的字符')])#该行代码过滤出't'列中含有”包含的字符“的表格。

type_list = set(z for i in data['t'] for z in i.split(' '))#把列标题为’t‘的每一行按空格分块生成列表。set可进行去重
for ty in type_list:
    data[data['t'].str.contains(ty)].to_excel(writer,sheet_name=ty)#以去重后的每一个列表元素作为每个sheet名，把包含每个列表元素的整行存到相应sheet中

import os
import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
file_path='F:\\测试用表格'
file_list=os.listdir(file_path)
for i in file_list:
    if os.path.splitext(i)[1]=='.xlsx':
        workbook=app.books.open(file_path+'\\'+i)
        worksheet=workbook.sheets
        for j in worksheet:
            values=j.range('A1').expand('table').options(pd.DataFrame).value
            print(values)
            values['销售利润']=values['销售利润'].astype('float')#转换“销售利润”列的数据类型
            result=values.groupby('销售区域').sum()# 根据“销售区域”列对数据进行分类汇总，汇总运算方式为求和,还可以使用其他
# 函数完成其他类型的汇总运算。常用的有：用mean()函数求平均值，用count()函数统计个
# 数，用max()函数求最大值，用min()函数求最小值
            j.range('J1').value=result['销售利润']# 将各个销售区域的销售利润汇总结果写入当前工作表
        workbook.save()
        workbook.close()
app.quit()
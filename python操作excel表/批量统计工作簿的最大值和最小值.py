import os
import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
file_path=''
file_list=os.listdir(file_path)
for j in file_list:
    if os.path.splitext(j)[1]=='.xlsx':
        workbook=app.books.open(file_path+'\\'+j)
        worksheet=workbook.sheets
        for i in worksheet:
            values=i.range('A1').expand('table').options(pd.DataFrame).value
            max=values['销售利润'].max()
            min=values['销售利润'].min()
            i.range('I1').value='最大销售利润'#统计结果写入的表格区域
            i.range('J1').value=max
            i.range('I2').value='最小销售利润'
            i.range('J2').value=min
            workbook.save()
            workbook.close()
app.quit()
#除了sum()、mean()、count()、max()、min()等函数，还可以用value
# counts()函数统计
# 重复值的个数，用product()函数计算乘积，用std()函数计算标准差，等等。
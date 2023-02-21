import os
import xlwings as xw
import pandas as pd
file_path='产品记录表'
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    worksheet=workbook.sheets['规格表']#指定要处理的工作表
    values=worksheet.range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value#读取表中数据
    new_values=values['规格'].str.split('*',expand=True)#根据“*”号拆分规则列
    values['长(mm)']=new_values[0]# 将拆分出的第1部分数据添加到标题为“长(mm)”的列中
    values['宽(mm)']=new_values[1]
    values['高(mm)']=new_values[2]
    values.drop(columns=['规格'],inplace=True)#删除’规格‘列
    worksheet['A1'].options(index=False).value=values
    worksheet.autofit()# 根据数据内容自动调整工作表的行高和列宽
    workbook.save()
    workbook.close()
app.quit()


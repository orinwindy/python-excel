import os
import xlwings as xw
import pandas as pd
file_path='F:\\测试用表格'
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith("~$"):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    worksheet=workbook.sheets["规格表"]
    values=worksheet.range('A1').options(pd.DataFrame,header=1,index=False,expand="table").value
    values['规格']=values['长(mm)'].astype("str")+'*'+values['宽(mm)'].astype('str')+'*'+values['高(mm)'].astype('str')#合并列数据
    values.drop(columns=['长(mm)'],inplace=True)#删除标题为“长(mm)”的列
    values.drop(columns=['宽(mm)'],inplace=True)
    values.drop(columns=['高(mm)'],inplace=True)
    worksheet.clear()#清除工作表“规格表”中原有的数据
    print(values)
    worksheet['A1'].options(index=False).value=values#将处理好的数据写入工作表
    worksheet.autofit()
    workbook.save()
    workbook.close()
app.quit()
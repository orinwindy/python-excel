import os
import xlwings as xw
file_path="F:\\测试用表格"
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    worksheet=workbook.sheets['Sheet1']#指定要修改的工作表名
    value=worksheet['A2'].expand('table').value
    for index,val in enumerate(value):
        print(val[2])
        val[2]=val[2]*(1+0.05)#修改第三个单元格的数据，上调5%
        value[index]=val#替换整行数据
    worksheet['A2'].expand('table').value=value#将替换后的数据写入表格
    workbook.save()
    workbook.close()
app.quit()
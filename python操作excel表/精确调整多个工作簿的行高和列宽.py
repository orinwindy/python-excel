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
    for j in workbook.sheets:
        value=j.range('A1').expand('table')
        value.column_width = 12#列宽
        value.row_height=20#行高
    workbook.save()
    workbook.close()
app.quit()
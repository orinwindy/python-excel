import os
import xlwings as xw

file_path=""
file_list=os.listdir(file_path)
sheet_name="特定的工作表名称"
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith("~$"):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    for j in workbook.sheets:
        if j.name ==sheet_name:
            j.api.PrintOut()
            break
    app.quit()
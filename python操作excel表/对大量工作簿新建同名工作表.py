import os
import xlwings as xw

file_path = "f:\\科技研发部考勤"
file_list = os.listdir(file_path)
sheet_name = "2023年本月份考勤表"
app = xw.App(visible=False, add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths = os.path.join(file_path, i)
    workbook = app.books.open(file_paths)
    sheet_names = [j.name for j in workbook.sheets]
    if sheet_name not in sheet_names:
        workbook.sheets.add(sheet_name)
        workbook.save()
app.quit()

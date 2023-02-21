import os
import xlwings as xw

file_path = "f:\\科技研发部考勤"
file_list = os.listdir(file_path)
# print(file_list)
sheet_name = "2023年本月份考勤表"
app = xw.App(visible=False, add_book=False)
for i in file_list:
    if i.startswith("~$"):
        continue
    file_paths = os.path.join(file_path, i)
    workbook = app.books.open(file_paths)
    for j in workbook.sheets:
        if j.name == sheet_name:
            j.delete()
            break
    workbook.save()
app.quit()

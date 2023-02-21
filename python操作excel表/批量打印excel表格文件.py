import os
import xlwings as xw

file_path = "f:\\"#要打印的文件夹路径
file_list = os.listdir(file_path)
app = xw.App(visible=False, add_book=False)
for i in file_list:
    if i.startswith("~$"):
        continue
    file_paths=os.path.join(file_path,i)
    workbook = app.books.open(file_paths)
    workbook.api.PrintOut()
app.quit()
# 该脚本将工作表名称改为与工作簿名称一致
import xlwings as xw
import os

file_path="f:\\科技研发部考勤"
file_list = os.listdir(file_path)
app=xw.App(visible=False, add_book=False)
for i in file_list:
    print(i)
    # print((type(i)))
    workbook = app.books.open(file_path+"\\"+i)
    worksheet = workbook.sheets[0]
    # print(worksheets)
    worksheet.name = os.path.splitext(i)[0]
    # print(os.path.splitext(i)[0])
    workbook.save()
app.quit()
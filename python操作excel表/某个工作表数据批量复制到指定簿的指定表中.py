import os
import xlwings as xw

app = xw.App(visible=False, add_book=False)
file_path = "f:\\新建文件夹"
file_list = os.listdir(file_path)
workbook = app.books.open("f:\\源数据.xlsx")  # 打开源数据簿
worksheet = workbook.sheets["Sheet1"]
value = worksheet.range('A1').expand('table')
start_cell = (2, 1)
end_cell = (value.shape[0], value.shape[1])
cell_area = worksheet.range(start_cell, end_cell).value
for i in file_list:
    if os.path.splitext(i)[1] == ".xlsx":
        try:
            workbooks = xw.books.open(file_path + "\\" + i)
            sheet = workbooks.sheets['Sheet1']
            scope = sheet.range("A1").expand()
            sheet.range(scope.shape[0]+1, 1).value = cell_area#+1,1代表复制过来的数据放在第二行第一列，与源数据的（2，1）位置对应相等
            workbooks.save()
        finally:
            workbooks.close()
workbook.close()
app.quit()

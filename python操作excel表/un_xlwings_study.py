import xlwings as xw

# 新建工作簿，visible可见性，add_book表示新建1个sheet
app = xw.App(visible=True,add_book=False)
# 新建一张表
# workbook = app.books.add()
# 保存工作簿
# workbook.save("d:\\1.xlsx")
# 关闭工作簿
# workbook.close()
# 退出Excel程序
# app.quit()

# 打开工作簿
workbook=app.books.open(r'd:\1.xlsx')

# 操作工作表和单元格。
worksheet = workbook.sheets["Sheet1"]
worksheet.range("A1").value="编号"
workbook.save("d:\\1.xlsx")
workbook.close()
app.quit()
import xlwings as xw
import pandas as pd
app=xw.App(visible=True,add_book=False)
workbook=app.books.open('')
worksheet=workbook.sheets
for i in worksheet:
    values=i.range('A1').expand('table').options(pd.DataFrame).value
    sums=values['采购金额'].sum()
    i.range('F1').value=sums#将当前工作表中数据的求和结果写入当前工作表的单元格F1中
workbook.save()
workbook.close()
app.quit()
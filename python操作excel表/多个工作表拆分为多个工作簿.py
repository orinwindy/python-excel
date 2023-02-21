#本代码将多个表的每个表拆分为一个工作簿

import xlwings as xw
workbook_name="f:\\"#指定要拆分的来源工作簿
app=xw.App(visible=False,add_book=False)
header=None
all_data=[]
workbook = app.books.open(workbook_name)
for i in workbook.sheets:
    workbook_split=app.books.add()
    sheet_split=workbook_split.sheets[0]
    i.api.Copu(Before=sheet_split.api)
    workbook_split.save('{}'.format(i.name))
app.quit()

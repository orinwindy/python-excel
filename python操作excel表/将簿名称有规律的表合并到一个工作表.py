import os
import xlwings as xw
workbook_name = ""
sheet_names=[str(sheet)+'月'for sheet in range(1,7)]#列出要合并工作表的工作簿
new_sheet_name="上半年统计表"
app=xw.App(visible=False,add_book=False)
header=None
all_data=[]
workbook=app.books.open(workbook_name)
for i in workbook.sheets:
    if new_sheet_name in i.name:
        i.delete()#判断工作簿中是否已经存在：上半年统计表 的工作表
        #若存在就删除已存在的表
new_worksheet=workbook.sheets.add(new_sheet_name)
title_copyed=False
for j in workbook.sheets:
    if j.name in sheet_names:
        if title_copyed==False:
            j['A1'.api.EntireRow.Copy(Destination=new_worksheet["A1"].api)]#将要合并的工作表的列标题复制到新增工作表“上半年统计表中”
            title_copyed=True
        row_num=new_worksheet['A1'].current_region.last_cell.row
        j['A1'].current_region.offset(1,0).api.Copy(Destination=new_worksheet["A{}".format(row_num+1)].api)
new_worksheet.autofit()
workbook.save()
app.quit()
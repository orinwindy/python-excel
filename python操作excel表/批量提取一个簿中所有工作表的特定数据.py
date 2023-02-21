import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
workbook=app.books.open("")#打开工作簿
worksheet=workbook.sheets
data=[]
for i in worksheet:
    values=i.range('A1').expand().options(pd.DataFrame).value#读取当前工作表的所有数据
    filtered=values[values["采购物品"]=='复印纸']#提取“采购物品”为“复印纸”的行数据
    if not filtered.empty:#判断提出的行数据是否为空
        data.append(filtered)#将提取到的行数据追加到列表中
new_workbook=xw.books.add()
new_worksheet=new_workbook.sheets.add('复印纸')#在工作簿中新建一个为“复印纸”的工作表
new_worksheet.range('A1').value=pd.concat(data,ignore_index=False)#将提取出的行数据写入工作表“复印纸”中
new_workbook.save()
workbook.close()
app.quit()
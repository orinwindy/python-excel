#本脚本对A2列的数据进行去重并保存为书名xlsx
import xlwings as xw
app=xw.App(visible=True,add_book=False)
workbook=app.books.open('f:\\测试用表格\\测试.xlsx')
data=[]
for i,worksheet in enumerate(workbook.sheets):
    values=worksheet['A2'].expand('down').value
    data=data+values#将提取出的书名数据添加到前面创建的列表中
data=list(set(data))#对列表中的书名数据进行去重操作
data.insert(0,'书名')#在去重后的书名数据前添加列标题“书名”
new_workbook=xw.books.add()
new_worksheet=new_workbook.sheets.add('书名')
new_worksheet['A1'].options(transpose=True).value=data#将处理好的书名数据写入新工作表
new_worksheet.autofit()
new_workbook.save('f:\\测试用表格\\书名.xlsx')
new_workbook.close()
app.quit()

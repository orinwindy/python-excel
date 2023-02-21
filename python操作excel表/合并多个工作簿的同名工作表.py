import os
import xlwings as xw
file_path="f:\\"#源工作簿路径
file_list=os.listdir(file_path)
sheet_name='产品销售统计'#要合并的表名称
app=xw.App(visible=False,add_book=False)
header=None#定义变量header，初始值为一个空对象，后续用于存放要合并的工作表中的数据列标题
all_data=[]
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    for j in workbook.sheets:
        if j.name==sheet_name:
            if header==None:
                header=j['A1:I1'].value#如果未存放，则读取列标题并赋给变量header
            values=j['A2'].expand('table').value#数据内容存放
            all_data=all_data+values#存放所有数据
new_workbook=xw.Book()
new_worksheet=new_workbook.sheets.add(sheet_name)
new_worksheet['A1'].value=header#列标题存入表格
new_worksheet['A2'].value=all_data#数据存入表格
new_worksheet.autofit()
new_workbook.save("f:\\")
app.quit()
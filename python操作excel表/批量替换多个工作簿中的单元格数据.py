import os
import xlwings as xw
file_path=''
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith("~$"):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=xw.books.open(file_paths)
    for j in workbook.sheets:
        value=j['A2'].expand('table').value
        for index,val in enumerate(value):#按行遍历工作表数据
            if val[0]=='背包':#判断当前第一个单元格的数据是否为“背包”
                val[0]='双肩包'#若是，则替换为该值
        j['A2'].expand('table').value=value#放入表内

    workbook.save()
    workbook.close()
app.quit()
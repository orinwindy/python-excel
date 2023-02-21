import os
import xlwings as xw
file_path='F:\\测试用表格'
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    for j in workbook.sheets:
        value=j['A2'].expand('table').value
        print(value)
        for index,val in enumerate(value):# 按行遍历工作表数据
            if val==['背包',16,65]:# 判断行数据是否为“背包”、16、65
                value[index]=['双肩包',36,79]# 如果是，则将该行数据替换为新的数据
        j['A2'].expand('table').value=value# 将完成替换的数据写入工作表
    workbook.save()
    workbook.close()
app.quit()
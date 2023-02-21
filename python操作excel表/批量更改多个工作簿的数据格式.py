import os
import xlwings as xw

file_path=''
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    for j in workbook.sheets:
        row_num=j['A1'].current_region.last_cell.row#获取工作表数据区域最后一行的行号
        j['A2:A{}'.format(row_num)].number_format='m/d'#将A列的数据全部改为“/月/日”
        j['D2:D{}'.format(row_num)].number_format='¥#，##0.00'#将D列的数据格式全部更改
    workbook.save()
    workbook.close()
app.quit()
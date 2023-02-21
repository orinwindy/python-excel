import os
import xlwings as xw
newContent=[['双肩包','64','110'],['腰包','23','58']]#给出要追加的行数据
app=xw.apps.add()
file_path='f:\\测试用表格'
file_list=os.listdir(file_path)
for i in file_list:
    if os.path.splitext(i)[1]=='.xlsx':
        workbook=app.books.open(file_path+'\\'+i)
        worksheet=workbook.sheets['Sheet1']#指定要追加行数的工作表
        values=worksheet.range('A1').expand()
        # print(values)
        number=values.shape[0]#获取原有数据的行数
        worksheet.range(number+1,1).value=newContent#将前面指定的行数据追加到原有数据的下方
        workbook.save()
        workbook.close()
app.quit()
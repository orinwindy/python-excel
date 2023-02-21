import os
import xlwings as xw

app = xw.App(visible=False, add_book=False)
file_path = "f:\\科技研发部考勤"  # 目标工作簿所在的文件路径
file_list = os.listdir(file_path)
workbook = app.books.open('F:\\科技研发部考勤\\2023年科技研发部1月份考勤表.xlsm')  # 打开来源工作簿
worksheet = workbook.sheets
for i in file_list:
    if os.path.splitext(i)[1] == '.xlsm':
        workbooks = app.books.open(file_path + '\\' + i)
        for j in worksheet:
            contents = j.range('A1:AI8').value
            name = j.name
            workbooks.sheets.add(name=name, after=len(workbooks.sheets))  # 在目标工作簿中新增同名工作表
            workbooks.sheets[name].range('A1:AI8').value = contents  # 内容复制，（注:格式无法复制）
        workbooks.save()
app.quit()

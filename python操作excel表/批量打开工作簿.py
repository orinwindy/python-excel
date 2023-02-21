# 文件路径：f"f:\\科技研发部考勤\\2023年科技研发部{i}月份考勤表.xlsx

import os
import xlwings as xw

file_path= "f:\\科技研发部考勤"
file_list = os.listdir(file_path)
app=xw.App(visible=True,add_book=False)
for i in file_list:
    if os.path.splitext(i)[1]==".xlsm":
        name = os.path.splitext(i)[0]
        app.books.open(file_path+'\\'+i)
        print(f"{name}打开成功！")
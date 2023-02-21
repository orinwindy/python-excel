import xlwings as xw

app=xw.App(visible=True,add_book=False)
for i in range(1,13):
    workbook = app.books.add()
    # 该步骤需要确保目录存在，不存在会报错
    workbook.save(f"f:\\科技研发部考勤\\2023年科技研发部{i}月份考勤表.xlsm")
    workbook.close()

app.quit()
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")

# excel.Visible = True     # 可以让excel 可见

# 这里填写要修改的Excel文件的绝对路径
workbook = excel.Workbooks.Open(r"F:\pyCharmproject\python操作excel表\income.xlsx")

# 得到 2017 表单
sheet = workbook.Sheets('2017')

# 修改表单第一行第一列单元格内容
# com接口，单元格行号、列号从1开始
sheet.Cells(1,1).Value="你好"

# 保存内容
workbook.Save()

# 关闭该Excel文件
workbook.Close()

# excel进程退出
excel.Quit()

# 释放相关资源
sheet = None
book = None
excel.Quit()
excel = None
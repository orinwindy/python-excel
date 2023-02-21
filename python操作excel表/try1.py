import xlwings as xw  # 导入xlwings模块

file_path = 'income.xlsx'  # 给出来源工作簿的文件路径
sheetname = '2018'  # 给出要拆分的工作表的名称
app = xw.App(visible=True, add_book=False)  # 启动Excel程序
workbook = app.books.open(file_path)  # 打开来源工作簿
worksheet = workbook.sheets[sheetname]  # 选中要拆分的工作表
value = worksheet.range('A2').expand('table').value  # 读取要拆分的
# 工作表中的所有数据
print(len(value))
print(value[0][1])
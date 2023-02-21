import xlrd

# 实例化工作蒲对象
book = xlrd.open_workbook("income.xlsx")
# print(f"包含表单的数量{book.nsheets}")
# print(f"表单的名分别为：{book.sheet_names()}")
sheet1 = book.sheet_by_name('2018')
sheet2 = book.sheet_by_name("2017")
sheet3 = book.sheet_by_name("2016")
# 数据显示
# print(f"表单名：{sheet.name} ")
# print(f"表单索引：{sheet.number}")
# print(f"表单行数：{sheet.nrows}")
# print(f"表单列数：{sheet.ncols}")
# print(f"第一行的内容是：{sheet.row_values(rowx=0)}")
# print(f"第一列的内容是：{sheet.col_values(colx=0)}")

# 计算全年的总和
incomes = sheet1.col_values(colx=1,start_rowx=1)#收入在第二列，从第二行开始加
print(f"2018年的总收入为：{sum(incomes)}")
#找出带星标的月份
months = sheet1.col_values(colx=0)
tosubstract = 0#该值存取星标月份的工资总和
for row,month in enumerate(months):
    if type(month) is str and month.endswith('*'):
        income = sheet1.cell_value(row,1)
        # print(month,income)
        tosubstract += income
print(f"2018年去星标收入为：{int(sum(incomes)-tosubstract)}")
a= int(sum(incomes)-tosubstract)

incomes = sheet2.col_values(colx=1,start_rowx=1)#收入在第二列，从第二行开始加
print(f"2017年的总收入为：{sum(incomes)}")
#找出带星标的月份
months = sheet2.col_values(colx=0)
tosubstract = 0#该值存取星标月份的工资总和
for row,month in enumerate(months):
    if type(month) is str and month.endswith('*'):
        income = sheet2.cell_value(row,1)
        # print(month,income)
        tosubstract += income
print(f"2017年去星标收入为：{int(sum(incomes)-tosubstract)}")
b= int(sum(incomes)-tosubstract)

incomes = sheet3.col_values(colx=1,start_rowx=1)#收入在第二列，从第二行开始加
print(f"2016年的总收入为：{sum(incomes)}")
#找出带星标的月份
months = sheet3.col_values(colx=0)
tosubstract = 0#该值存取星标月份的工资总和
for row,month in enumerate(months):
    if type(month) is str and month.endswith('*'):
        income = sheet3.cell_value(row,1)
        # print(month,income)
        tosubstract += income
print(f"2016年去星标收入为：{int(sum(incomes)-tosubstract)}")
c= int(sum(incomes)-tosubstract)
print(f"三年去星标收入为：{a+b+c}")
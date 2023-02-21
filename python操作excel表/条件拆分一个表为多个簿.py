#思路整理：给出源数据路径-给出被拆分表格名称-启动excel程序-打开源工作簿-选中要拆分的工资表-读取要拆分的数据-创建空字典存储数据-遍历读取数据，字典的键为商品名称，字典的值为各商品的项数据
#-新建工作簿-源数据表列标题复制过来-源数据数据放到列标题下A2开始的区域。保存表格-退出excel文件
import xlwings as xw

file_path = "f:\\新建文件夹\\源数据.xlsx"  # 源工资簿的文件路径
sheet_name = 'Sheet1'  # 源工作表名称
app = xw.App(visible=True, add_book=False)
workbook = app.books.open(file_path)  # 打开源数据工作簿
worksheet = workbook.sheets(sheet_name)  # 选中源数据工资表
value = worksheet.range('A2').expand('table').value  # 读取要拆分的工作表中的所有数据
data = dict()
# print(len(value))#列表的元素个数，也即是行数
# print(value)#显示得到信息value为一个列表，列表的每一个元素是存储每一行信息的列表
for i in range(len(value)):
    product_name = value[i][1]
    if product_name not in data:
        data[product_name] = []
    data[product_name].append((value[i]))
for key, value in data.items():
    new_workbook = xw.books.add()
    new_worksheet = new_workbook.sheets.add(key)
    new_worksheet["A1"].value=worksheet["A1:H1"].value#将源数据的列标题复制到新建簿中
    new_worksheet["A2"].value=value
    new_workbook.save("{}.xlsx".format(key))
app.quit()

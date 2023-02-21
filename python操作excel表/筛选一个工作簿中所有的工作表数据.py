# 本脚本把簿中所有按采购物品分成很多表
import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
workbook=app.books.open('F:\\测试用表格\\ces.xlsx')
worksheet=workbook.sheets
table=pd.DataFrame()
for i,j in enumerate(worksheet):
    values=j.range('A1').options(pd.DataFrame,header=1,index=False,expand='table').value
    data=values.reindex(columns=['采购物品','采购日期','采购数量','采购金额'])#调整列的顺序，将’采购物品‘移到第一列
    table=pd.concat([table,data],ignore_index=True)#将调整列顺序后的数据合并到前面创建的DataFrame中
table=table.groupby('采购物品')#根据‘采购物品’列筛选数据
new_workbook=xw.books.add()
for idx, group in table:#遍历筛选好的数据，其中idx对应物品名称，group对应该物品的说有明细数据
    new_worksheet=new_workbook.sheets.add(idx)
    new_worksheet['A1'].options(index=False).value=group
    last_cell=new_worksheet['A1'].expand('table').last_cell#获取当前工作表数据区域右下角的单元格
    last_row=last_cell.row#获取数据区域最后一行的行号
    last_column=last_cell.column
    last_column_letter=chr(64+last_column)#将数据区域最后一列的列号（数字）转换为该列的列标
    sum_cell_name='{}{}'.format(last_column_letter,last_row+1)
    sum_last_row_name='{}{}'.format(last_column_letter,last_row)
    formula='=SUM({}2:{})'.format(last_column_letter,sum_last_row_name)
    new_worksheet[sum_cell_name].formula=formula
    new_worksheet.autofit()
new_workbook.save('F:\\测试用表格\\采购分类表.xlsx')
workbook.close()
app.quit()
#将“SUM”改为“AVERAGE”就是求平均值，改为“MAX”就是求
#最大值


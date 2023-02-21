import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
workbook=app.books.open('产品销售统计表.xlsx')
worksheet=workbook.sheets
for i in worksheet:
    values=i.range('A1').expand('table').options(pd.DataFrame).value# 读取当前工作表的数据并转换为DataFrame格式
    result=values.sort_values(by='销售利润')#对销售利润列进行升序排序,若要实现降序则在括号里加入ascending=False即可
    i.range('A1').value=result#将排序结果写入当前工作表，替换原有数据
workbook.save()
workbook.close()
app.quit()
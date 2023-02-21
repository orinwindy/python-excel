import xlwings as xw
app=xw.App(visible=True,add_book=False)
workbook=app.books.open("F:\\测试用表格\\员工销售业绩统计表.xlsx")
for i in workbook.sheets:
    chart=i.charts.add(left=200,top=0,width=355,height=211)
    chart.set_source_data(i['A1'].expand('table'))
    chart.chart_type='bar_clustered'#制作条形图
workbook.save('F:\\测试用表格\\条形图.xlsx')
workbook.close()
app.quit()
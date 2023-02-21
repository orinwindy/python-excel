import xlwings as xw
app=xw.App(visible=True,add_book=False)
workbook=app.books.open('F:\\测试用表格\\员工销售业绩统计表.xlsx')
for i in workbook.sheets:
    chart=i.charts.add(left=200,top=0,width=355,height=211)#设置图表的位置和尺寸
    chart.set_source_data(i['A1'].expand())#读取工作表中要制作图表的数据
    chart.chart_type='column_clustered'#制作柱形图   #bar_clustered条形图   #line折线图    #area面积图    #pie饼图  #doughnut圆环图    #xy_scatter散点图  #radar雷达图
workbook.save('F:\\测试用表格\\柱形图.xlsx')
workbook.close()
app.quit()
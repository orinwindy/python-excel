import os
import xlwings as xw
file_path='F:\\测试用表格'
file_list=os.listdir(file_path)
app=xw.App(visible=False,add_book=False)
for i in file_list:
    if i.startswith('~$'):
        continue
    file_paths=os.path.join(file_path,i)
    workbook=app.books.open(file_paths)
    for j in workbook.sheets:
        j['A1:H1'].api.Font.Name='宋体'# 设置工作表标题行的字体为“宋体”
        j['A1:H1'].api.Font.Size =10# 设置工作表标题行的字号为“10”磅
        j['A1:H1'].api.Font.Bold=True# 加粗工作表标题行
        j['A1:H1'].api.Font.Color=xw.utils.rgb_to_int((255,255,255))# 设置工作表标题行的字体颜色为“白色”
        j['A1:H1'].color=xw.utils.rgb_to_int((0,0,0))# 设置工作表标题行的单元格填充颜色为“黑色”
        j['A1:H1'].api.HorizontalAlignment=xw.constants.HAlign.xlHAlignCenter# 设置工作表标题行的水平对齐方式为“居中”
        j['A1:H1'].api.VerticalAlignment=xw.constants.VAlign.xlVAlignCenter# 设置工作表标题行的垂直对齐方式为“居中”
        j['A2'].expand('table').api.Font.Name='宋体'# 设置工作表的正文字体为“宋体”
        j['A2'].expand('table').api.Font.Size=10# 设置工作表的正文字号为“10”磅
        j['A2'].expand('table').api.HorizontalAlignment=xw.constants.HAlign.xlHAlignLeft# 设置工作表正文的水平对齐方式为“靠左”
        j['A2'].expand('table').api.VerticalAlignment=xw.constants.HAlign.xlHAlignCenter# 设置工作表正文的垂直对齐方式为“居中”
        for cell in j['A1'].expand('table'):# 从单元格A1开始为工作表添加合适粗细的边框
            for b in range(7,12):
                cell.api.Borders(b).LineStyle=1#设置单元格的边框线型
                cell.api.Borders(b).Weight=2#设置单元格的边框粗细
    workbook.save()
    workbook.close()
app.quit()
import openpyxl

name2Age = [
    ['张飞' ,  38 ] ,
    ['赵云' ,  27 ] ,
    ['许褚' ,  36 ] ,
    ['典韦' ,  38 ] ,
    ['关羽' ,  39 ] ,
    ['黄忠' ,  49 ] ,
    ['徐晃' ,  43 ] ,
    ['马超' ,  23 ]
]

book = openpyxl.Workbook()

sh = book.active

sh.title = '年龄表'

sh["A1"]="name"
sh["B1"]="age"

for row in name2Age:
    sh.append(row)

book.save("列表or元组写入excel.xlsx")
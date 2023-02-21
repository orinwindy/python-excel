import os
import xlwings as xw
import pandas as pd
app=xw.App(visible=False,add_book=False)
file_path=''
file_list=os.listdir(file_path)
collection=[]
for i in file_list:
    if os.path.splitext(i)[1]=='.xlsx':
        workbook=app.books.open(file_path+'\\'+i)
        worksheet=workbook.sheets['销售记录表']
        values=worksheet.range('A1').expand('table').options(pd.DataFrame).value
        filtered=values[['销售区域','销售利润']]#只保留'销售区域'合'销售利润'这两列数据
        collection.append(filtered)
        workbook.close()
new_values=pd.concat(collection,ignore_index=False).set_index('销售区域')
new_values['销售利润']=new_values['销售利润'].astype('float')
result=new_values.groupby('销售区域').sum()
new_workbook=app.books.add()
sheet=new_workbook.sheets[0]
sheet.range('A1').value=result
new_workbook.save('汇总.xlsx')
app.quit()
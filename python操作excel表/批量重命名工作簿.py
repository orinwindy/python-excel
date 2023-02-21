# 该代码能将指定文件夹下的文件名部分文字中的科技研发部修改为商务部，同时能忽略临时文件的存在

import os
file_path = "f:\\科技研发部考勤"
file_list = os.listdir(file_path)
old_book_name = "科技研发部"
new_book_name = "商务部"
for i in file_list:
    if i.startswith('~$'):
        continue
    new_file = i.replace(old_book_name,new_book_name)
    old_file_path = os.path.join(file_path,i)
    new_file_path = os.path.join(file_path,new_file)
    os.rename(old_file_path,new_file_path)
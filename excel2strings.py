# -*- coding: utf-8 -*-
import  xdrlib ,sys, xlrd, os, time
reload(sys)
sys.setdefaultencoding( "utf-8" )
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)

# file：Excel文件路径
# colnameindex：表头列名所在行的所以  ，
# by_index：表的索引
def excel_table_byindex(file, colnameindex=0, by_index=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.col_values(colnameindex) #某一行数据
    list = colnames[2:]  #此处可能需要更改
    return list

# 创建strings文件
# base_list: 基础列表
# other_list: 对应的翻译
# dir_name: 文件夹名
# language: 语言名
def create_strings( base_list, other_list, dir_name, language='en'):
    file_dir_name = dir_name+'/'+language+'.lproj'
    os.mkdir(file_dir_name)
    f=open(file_dir_name+'/'+'Localizable.strings','w')
    for index, string in enumerate(base_list):
        f.write('"'+str(string)+'" = '+'"'+str(other_list[index])+'";\n')
    f.close()

#  file: 文件路径
# dir_name: 文件夹名
def main(file, dir_name=''):
    # 如果没有输入文件夹名字，文件夹名为'exc2string+时间戳'
   if dir_name == '':
       dir_name = 'excel2strings' + str(time.time())

   # 此处可以自定义
   base_tables = excel_table_byindex(file, 0)
   zh_tables = excel_table_byindex(file, 1)
   en_tables = excel_table_byindex(file, 2)
   ko_tables = excel_table_byindex(file, 3)
   ja_tables = excel_table_byindex(file, 4)

   os.mkdir(dir_name)
   # 此处可以自定义
   create_strings(base_tables, en_tables, dir_name, 'en')
   create_strings(base_tables, zh_tables, dir_name, 'zh-Hans')
   create_strings(base_tables, ja_tables, dir_name, 'ja')
   create_strings(base_tables, ko_tables, dir_name, 'ko')

# argv[1]: excel文件路径
if __name__=="__main__":
    file_name = sys.argv[1]
    main(file_name)

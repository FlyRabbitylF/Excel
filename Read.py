import xlwings as xw

wb=xw.Book(r'C:\Users\lyh\Desktop\pythonProject\Excel\Test.xlsx')

shtList=[]

#  获取所有工作表名字
for i in range(0,wb.sheets.count):
    shtList.append(wb.sheets[i].name)

# 单元格读取 赋值
sh=wb.sheets['sheet1']
sh.range('a1').value
print(sh.range('a1').value)
sh.range('d1').value='d1'

sh.range('a1').expand()

# 按列获取代码 区域不可间断
for i in sh.range('a1').expand().value:
    print(i)

# 获取行数
print(sh.range('a1').expand().rows.count)

# 获取列数
print(sh.range('a1').expand().columns.count)

# 获取元素个数
print(sh.range('a1').expand().count)

# 获取某个单元格上面有内容的单元格的行数
print(sh.range('a6666').end('up').row)
sh.range('a6666').end('up').row

# 获取某个单元格下面有内容的单元格的行数
print(sh.range('a1').end('down').row)

# 已使用的最大行数 列数
rows=wb.sheets['sheet1'].used_range.rows.count
cols=wb.sheets['sheet1'].used_range.columns.count

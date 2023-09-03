import xlwings as xw

wb=xw.Book(r'C:\Users\lyh\Desktop\pythonProject\Excel\Test.xlsx')

# 新建工作表
# 在最前面插入
# sht=wb.sheets.add()

# 最后插入工作表
# sht=wb.sheets.add(after=wb.sheets.count)

# 工作表命名
# sht.name='NewSheet'

# 名称/索引引用获取工作表名字
# print(wb.sheets[0].name)
# print(wb.sheets['hello'].name)

# 遍历工作表
# for sh in wb.sheets:
#     print(sh.name)
#
# for i in range(0,wb.sheets.count):
#     print(wb.sheets[i].name)

# 复制工作表
# wb.sheets['hello'].copy()

# 删除工作表
# wb.sheets['hello'].delete()

# 拆分
# for sht in wb.sheets:
#     print(sht.name)
#     sht.copy()
#     xw.books[xw.books.count-1].save(r'C:\Users\lyh\Desktop\pythonProject\Excel/' + sht.name +'.xlsx')
# xw.books[xw.books.count-1].close()

# 合并
# arr=[]
# for sht in wb.sheets:
#     if sht.name!='All':
#         brr=[]
#         brr=sht.range('a2').expand().value
#         arr+=brr
# wb.sheets['All'].range('a2').value=arr
#
# for i in arr:
#     print(i)
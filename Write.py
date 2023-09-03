import xlwings as xw

wb=xw.Book(r'C:\Users\lyh\Desktop\pythonProject\Excel\Test.xlsx')

sh=wb.sheets['sheet1']
# rng=sh.range('e1')
# # 横向赋值
# # rng.value=[10,11,12]
# # 纵向赋值
# rng.value=[[1],[2],[3]]
# # 区域赋值
# rng.value=[[1,2,3],[4,5,6],[7,8,9]]
# # 清空单元格
# rng.resize(3,3).clear_contents()

# 获取区域所有内容
arr=sh.range('a1:c30').value

#将获取区域的内容放到新的地方
rng=sh.range('e1')
rng.value=arr
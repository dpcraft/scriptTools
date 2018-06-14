# -*- coding: utf-8-*-
from win32com.client import Dispatch

info_path = 'F:\PythonWorkplace\input\info.xlsx'
print(info_path)
# 创建Excel对象
excel = Dispatch('Excel.Application')
excel.visible = 0        #不显示Excel
excel.DisplayAlerts = 0  #关闭系统警告
excel.ScreenUpdating = 0 #关闭屏幕刷新

# 打开Excel文件
workbook = excel.workbooks.Open(info_path)

# 激活第1个工作表
worksheet = workbook.worksheets[0]
worksheet.Activate()
# 获取当前工作表总行数
row_max = excel.Range("A56636").End(-4162).Row

# 直接A列获取数据(该方法可以获取得到数据列表)
# values = excel.Range("A1:A%s" % row_max).Value

# 也可以遍历
for row in range(1, row_max + 1):
    print(worksheet.Cells(row, 1).Value)
# 关闭Excel文件，不保存(若保存，使用True即可)
workbook.Close(False)

# 退出Excel
excel.Quit()



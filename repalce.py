# -*- coding: utf-8-*-

import win32com
from win32com.client import Dispatch

# 模板文件保存路径，此处使用的是绝对路径，相对路径未测试过
template_path = 'F:\PythonWorkplace\input\print.docx'
# excel 文件路径
info_path = 'F:\PythonWorkplace\input\info.xlsx'
# 另存文件路径，需要提前建好文件夹，不然会出错
store_path = 'F:\PythonWorkplace\output\\'
# 模板中需要被替换的文本。   u''中的u表示unicode字符，用于中文支持
NewStr = u'姓名'

# 启动word
# w = win32com.client.Dispatch('Word.Application')
# 或者使用下面的方法，使用启动独立的进程：
w = win32com.client.DispatchEx('Word.Application')

# 后台运行，不显示，不警告
w.Visible = 0
w.DisplayAlerts = 0
# 打开新的文件
doc = w.Documents.Open(FileName=template_path)
# worddoc = w.Documents.Add() # 创建新的文档

# 正文文字替换
w.Selection.Find.ClearFormatting()
w.Selection.Find.Replacement.ClearFormatting()

# 迭代替换名字，并以名字为名另存文件

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
for row in range(1, row_max + 1):
    print(worksheet.Cells(row, 1).Value)
    i = str.strip(worksheet.Cells(row, 1).Value)
    OldStr, NewStr = NewStr, i
    w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)
    doc.SaveAs(store_path + i + '.docx')
    # doc.PrintOut()     直接打印，未测试

doc.Close()
w.Quit()

# 关闭Excel文件，不保存(若保存，使用True即可)
workbook.Close(False)

# 退出Excel
excel.Quit()


# for i in lst:
#     OldStr, NewStr = NewStr, i
#     w.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)
#     doc.SaveAs(store_path + i + '.docx')
#     # doc.PrintOut()     直接打印，未测试
#
# doc.Close()
# w.Quit()



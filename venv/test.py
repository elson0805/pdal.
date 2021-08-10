import os
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

# 檔案路徑
file_path = str('C:\\Users\\PDAL\\Desktop\\VB-GTCA022')
# 製作一半的BOM表儲存路徑
onwork_BOM_open = str(file_path + "\\BOM表\\")
# BOM表儲存路徑
BOM_output_path = str(file_path + "\\auto\\BOM_output-GTCA022\\")

# wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
# sheets1 = wb1.get_sheet_names()  # 获取sheet页
# sheet1 = wb1.get_sheet_by_name('Sheet1')

def creat_excel():
    fn = 'C:\\Users\\PDAL\\Desktop\\test.xlsx'
    wb = openpyxl.Workbook()
    # wb.create_sheet("工作表1", 0) # 新增工作表並指定放置位置
    # wb.create_sheet("工作表2", 1)

    wb.save(fn)


def copy_worksheet():
    fn = 'C:\\Users\\PDAL\\Desktop\\catia_bom.xlsx'
    wb = openpyxl.load_workbook(fn)
    # 取得目前作用中的工作表
    actSheet = wb.active
    print(actSheet.title)
    target = wb.copy_worksheet(actSheet)
    target.title = 'sheet2'
    wb.save(fn)


def worksheet():
    wb = openpyxl.load_workbook(fn)
    wb.active = 0
    ws = wb.active
    print('B2內容： ', ws['B2'].value)
    ws['B2'].value =  20
    # print('B2內容： ', ws['B2'].value)
    # ws.cell(column=2, row=3).value = 999
    wb.save(fn) # 若給予不同檔名代表另存新檔的意思


def decide_Row():  # 判斷資料數目
    wb1 = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx", data_only=False)

    Rng1 = {"what": "Quantity", "After": "ActiveCell", "LookIn": "xlFormulas",
            "LookAt": "xlPart", "SearchOrder": "xlByRows", "SearchDirection": "xlNext",
            "MatchCase": False, "MatchByte": False, "SearchFormat": False}

    sheet = wb1['工作表1']

    for row in sheet.iter_rows(min_row=5, max_col=1, max_row=99, values_only=True):
        cunt = 0
        for value in row:
            if value == "":
                cunt -= 1
                break
            cunt += 1



decide_Row()


import csv
import win32com.client as win32
import openpyxl
from openpyxl.styles import Font, colors, Alignment
import xlwings as xw
import time

# 檔案路徑
file_path = str('C:\\Users\\PDAL\\Desktop\\VB-GTCA022')
# 模具規範路徑
die_rule_path = str(file_path + "\\auto\\die_rule\\")
# 製作一半的BOM表儲存路徑
onwork_BOM_open = str(file_path + "\\BOM表\\")
# BOM表儲存路徑
BOM_output_path = str(file_path + "\\auto\\BOM_output-GTCA022\\")

# # wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
# # sheets1 = wb1.get_sheet_names()  # 获取sheet页
# # sheet1 = wb1.get_sheet_by_name('Sheet1')
#
# def creat_excel():
#     fn = 'C:\\Users\\PDAL\\Desktop\\test.xlsx'
#     wb = openpyxl.Workbook()
#     # wb.create_sheet("工作表1", 0) # 新增工作表並指定放置位置
#     # wb.create_sheet("工作表2", 1)
#
#     wb.save(fn)
#
#
# def copy_worksheet():
#     fn = 'C:\\Users\\PDAL\\Desktop\\catia_bom.xlsx'
#     wb = openpyxl.load_workbook(fn)
#     # 取得目前作用中的工作表
#     actSheet = wb.active
#     print(actSheet.title)
#     target = wb.copy_worksheet(actSheet)
#     target.title = 'sheet2'
#     wb.save(fn)
#
#
# def worksheet():
#     wb = openpyxl.load_workbook(fn)
#     wb.active = 0
#     ws = wb.active
#     print('B2內容： ', ws['B2'].value)
#     ws['B2'].value =  20
#     # print('B2內容： ', ws['B2'].value)
#     # ws.cell(column=2, row=3).value = 999
#     wb.save(fn) # 若給予不同檔名代表另存新檔的意思
#
#
# def decide_Row():  # 判斷資料數目
#     wb1 = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx", data_only=False)
#
#     Rng1 = {"what": "Quantity", "After": "ActiveCell", "LookIn": "xlFormulas",
#             "LookAt": "xlPart", "SearchOrder": "xlByRows", "SearchDirection": "xlNext",
#             "MatchCase": False, "MatchByte": False, "SearchFormat": False}
#
#     sheet = wb1['工作表1']
#
#     for row in sheet.iter_rows(min_row=5, max_col=1, max_row=99, values_only=True):
#         cunt = 0
#         for value in row:
#             if value == "":
#                 cunt -= 1
#                 break
#             cunt += 1


data_size = []
cunt = 46

app = xw.App(visible=True, add_book=False)  # 程式可見，只打開不新建工作薄
app.display_alerts = False  # 警告關閉
app.screen_updating = True  # 螢幕更新關閉


def test():
    wb = app.books.open(onwork_BOM_open + "BOM_空白頁.xlsx")
    wb.save()  # 儲存檔案
    # wb.close()  # 關閉檔案
    # app.quit()  # 關閉程式


def output_bom():
    wb1 = load_workbook(BOM_output_path + "catia_bom.xlsx")
    wb2 = load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")
    (page) = decide_Page(cunt)
    # decide_Size(cunt, page)


def decide_Page(cunt):
    wb = load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")
    page = int(cunt / 30)
    if page < 1:
        page = 0
    for i in range(page, 2):
        sheet = wb['Sheet1']
        target = wb.copy_worksheet(sheet)
        target.title = 'Sheet' + str(i + 1)

    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30

    wb.save(BOM_output_path + "BOM_空白頁.xlsx")

    return page


def decide_Size(cunt, page):
    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page

    for j in range(1, page + 2):
        wb = openpyxl.load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")
        ws = wb["Sheet" + str(j)]
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Size", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            row = ws['G' + str(i + 4)]
            cell = row
            data_size.append(cell.value)
            # print(data_size)
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")
            ws = wb.active
            kss1 = {"What": "規格", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=3, value=cell.value)
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page


def create_catia_bom():
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.ActiveDocument
    product1 = document.Product
    wb1 = app.books.open(die_rule_path + "rule.xlsx")

    assemblyConvertor1 = product1.getItem("BillOfMaterial")
    arrayOfVariantOfBSTR1 = ["Quantity", "Part Number", "Material_Data", "Heat Treatment", "Product Description",
                             "Page", "Size"]

    assemblyConvertor1Variant = assemblyConvertor1
    # assemblyConvertor1Variant.SetSecondaryFormat(arrayOfVariantOfBSTR1)
    assemblyConvertor1Variant.SetCurrentFormat(arrayOfVariantOfBSTR1)

    # 含數據內容之BOM表(複製用)儲存路徑
    assemblyConvertor1.Print("XLS", BOM_output_path + "catia_bom.xlsx", product1)


create_catia_bom()

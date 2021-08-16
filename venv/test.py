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

from copy import copy
from openpyxl import load_workbook, Workbook

src_file = BOM_output_path + "catia_bom.xlsx"
tag_file = onwork_BOM_open + "BOM_空白頁.xlsx"
sheet_name = "Sheet2"

def replace_xls(src_file, tag_file, sheet_name):
    #        src_file是源xlsx文件，tag_file是目标xlsx文件，sheet_name是目标xlsx里的新sheet名称

    print("Start sheet %s copy from %s to %s" % (sheet_name, src_file, tag_file))
    wb = load_workbook(src_file)
    wb2 = load_workbook(tag_file)

    ws = wb.get_sheet_by_name(wb.get_sheet_names()[0])
    ws2 = wb2.create_sheet(sheet_name.decode('utf-8'))

    max_row = ws.max_row  # 最大行数
    max_column = ws.max_column  # 最大列数

    wm = zip(ws.merged_cells)  # 开始处理合并单元格
    if len(wm) > 0:
        for i in range(0, len(wm)):
            cell2 = str(wm[i]).replace('(<MergeCell ', '').replace('>,)', '')
            print("MergeCell : %s" % cell2)
            ws2.merge_cells(cell2)

    for m in range(1, max_row + 1):
        ws2.row_dimensions[m].height = ws.row_dimensions[m].height
        for n in range(1, 1 + max_column):
            if n < 27:
                c = chr(n + 64).upper()  # ASCII字符,chr(65)='A'
            else:
                if n < 677:
                    c = chr(divmod(n, 26)[0] + 64) + chr(divmod(n, 26)[1] + 64)
                else:
                    c = chr(divmod(n, 676)[0] + 64) + chr(divmod(divmod(n, 676)[1], 26)[0] + 64) + chr(
                        divmod(divmod(n, 676)[1], 26)[1] + 64)
            i = '%s%d' % (c, m)  # 单元格编号
            if m == 1:
                #				 print("Modify column %s width from %d to %d" % (n, ws2.column_dimensions[c].width ,ws.column_dimensions[c].width))
                ws2.column_dimensions[c].width = ws.column_dimensions[c].width
            try:
                getattr(ws.cell(row=m, column=c), "value")
                cell1 = ws[i]  # 获取data单元格数据
                ws2[i].value = cell1.value  # 赋值到ws2单元格
                if cell1.has_style:  # 拷贝格式
                    ws2[i].font = copy(cell1.font)
                    ws2[i].border = copy(cell1.border)
                    ws2[i].fill = copy(cell1.fill)
                    ws2[i].number_format = copy(cell1.number_format)
                    ws2[i].protection = copy(cell1.protection)
                    ws2[i].alignment = copy(cell1.alignment)
            except AttributeError as e:
                print("cell(%s) is %s" % (i, e))
                continue

    wb2.save(BOM_output_path + "BOM_空白頁.xlsx")

    wb2.close()
    wb.close()


def decide_Page(cunt):
    wb2 = app.books.open(onwork_BOM_open + "BOM_空白頁.xlsx")
    wb = openpyxl.load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")
    page = int(cunt / 30)
    if page < 1:
        page = 0
    for i in range(page, 2):
        j = i + 1
        sheet = wb.worksheets[0]
        target = wb.copy_worksheet(sheet)
        target.title = 'Sheet' + str(j)

    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30


def test():
    wb2 = app.books.open(onwork_BOM_open + "BOM_空白頁.xlsx")
    wb = xw.Book(onwork_BOM_open + "BOM_空白頁.xlsx")
    sheet = wb.sheets['Sheet1']

    for i in range(1, 2):
        # 将sheet1工作表复制到该工作簿的最后一个工作表后面
        sheet2 = wb.sheets[-1]
        sheet.api.Copy(After=sheet2.api)

        wb.sheets[i].name = 'Sheet2' #重命名工作表


test()
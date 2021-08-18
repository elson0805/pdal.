import csv
import win32com.client as win32
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Border, Side, Font, colors, Alignment
from openpyxl.utils import get_column_letter
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

data = []
data_size = []
cunt = 46
page = 1

Frame_Commissioner = "專案人員"
Frame_guest_number = 654452
Finished_product_Name = "矽鋼片連續模"
Company_Name = "金屬中心"


app = xw.App(visible=True, add_book=False)  # 程式可見，只打開不新建工作薄
app.display_alerts = False  # 警告關閉
app.screen_updating = True  # 螢幕更新關閉

from copy import copy
from openpyxl import load_workbook, Workbook


def test():
    output_bom()
    information_bom(page)
    save()


def output_bom():
    wb1 = app.books.open(BOM_output_path + "catia_bom.xlsx")
    (cunt) = decide_Row()  # 搜尋資料數目
    (page) = decide_Page(cunt)  # BOM表頁數
    decide_Size(cunt, page)  # 尺寸
    decide_NO(cunt, page)  # 件號
    decide_name(cunt, page)  # 名稱
    decide_Quantity(cunt, page)  # 數量
    decide_material(cunt, page)  # 材質
    decide_Heat_treatment(cunt, page)  # 熱處理
    decide_description(cunt, page)  # 規格
    decide_Pa(cunt, page)  # 頁碼
    # decide_cost(cunt, page)  # 價格
    # draw_block(cunt, page)  # 備註

    return page


def information_bom(page):
    for j in range(1, page + 2):
        wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
        ws = wb.get_sheet_by_name("Sheet" + str(j))

        data = ["製    表", "製表日期", "頁    數", "模具編號", "品    號", "品    名"]
        kss = {"What": "製    表", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
               "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
               "SearchFormat": False}
        for Data in data:
            if Data == "製    表":
                ws.cell(row=4, column=7, value=Frame_Commissioner)

            elif Data == "製表日期":
                kss["What"] = "製表日期"
                localtime = time.localtime()
                result = time.strftime("%Y-%m-%d %I:%M:%S %p", localtime)
                ws.cell(row=5, column=7, value=result)

            elif Data == "頁    數":
                kss["What"] = "頁    數"
                ws.cell(row=3, column=7, value=j)

            elif Data == "模具編號":
                kss["What"] = "模具編號"
                ws.cell(row=3, column=3, value=Frame_guest_number)

            elif Data == "品    號":
                kss["What"] = "品    號"
                ws.cell(row=4, column=3, value=Frame_guest_number)

            elif Data == "品    名":
                kss["What"] = "品    名"
                ws.cell(row=5, column=3, value=Finished_product_Name)

        ws.cell(row=5, column=3, value=Company_Name)
        wb.save(BOM_output_path + "BOM_空白頁.xlsx")
    Adjustment(page)


def save():
    wb = openpyxl.Workbook(BOM_output_path + "BOM_空白頁.xlsx")
    wb.save(BOM_output_path + "BOM.xlsx")
    # FileName = write_BOM_location
    # FileFormat = xlNormal, Password = ""
    # WriteResPassword = ""
    # ReadOnlyRecommended = False
    # CreateBackup = False


def decide_Row():  # 判斷資料數目
    wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")

    Rng1 = {"what": "Quantity", "After": "ActiveCell", "LookIn": "xlFormulas",
            "LookAt": "xlPart", "SearchOrder": "xlByRows", "SearchDirection": "xlNext",
            "MatchCase": False, "MatchByte": False, "SearchFormat": False}

    ws = wb['工作表1']  # 獲取Sheet

    cunt = 0
    for row in ws['A5':'A99']:
        for cell in row:
            cunt += 1
        if cell.value is None:
            cunt -= 1
            break
        data.append(cell.value)
    # print(data)

    return cunt


def decide_Page(cunt):
    wb2 = app.books.open(onwork_BOM_open + "BOM_空白頁.xlsx")
    wb = load_workbook(onwork_BOM_open + "BOM_空白頁.xlsx")

    page = int(cunt / 30)
    if page < 1:
        page = 0
    for i in range(page, 2):
        j = i + 1
        target = wb['Sheet1']
        target1 = wb.copy_worksheet(target)
        target1.title = 'Sheet' + str(j)

        wb.save(BOM_output_path + "BOM_空白頁.xlsx")

        wb2.close()

    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30

    return page


def decide_Size(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    time.sleep(1)
    wb2.close()
    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1

    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Size", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            row = ws['G' + str(k + 4)]
            cell = row
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name('Sheet' + str(j))
            kss1 = {"What": "規格", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=3).value = cell.value
            # ==========================貼上BOM表資料==========================

            loops += 1
        if page0 == 0:
            pagenumb = cunt - 30 * page
        page0 -= 1

        wb.save(BOM_output_path + "BOM_空白頁.xlsx")

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        wb2.close()


def decide_NO(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    wb = load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    for j in range(1, page + 2):
        ws = wb.get_sheet_by_name("Sheet" + str(j))
        for i in range(1, pagenumb + 1):
            kss = {"What": "件號", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            ws.cell(row=(i + 6), column=1).value = i  # 依照順序填入編號

        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page
        wb.save(BOM_output_path + "BOM_空白頁.xlsx")
        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        wb2.close()


def decide_name(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    time.sleep(1)
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1

    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Part Number", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            cell = ws.cell(row=(k + 4), column=2)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "名稱", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=2).value = c
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def decide_Quantity(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1
    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Quantity", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            cell = ws.cell(row=(k + 4), column=1)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "數量", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=6).value = cell.value
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def decide_material(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1
    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(str(str(BOM_output_path) + "catia_bom.xlsx"))
            ws = wb.active
            kss = {"What": "Material_Data", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": "False"}

            cell = ws.cell(row=(k + 4), column=3)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "材質", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=4).value = c
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def decide_Heat_treatment(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1

    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            # ==========================複製BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Heat Treatment", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            cell = ws.cell(row=(k + 4), column=4)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "熱處理", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=5).value = cell.value
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def decide_description(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1
    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Product Description", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            cell = ws.cell(row=(k + 4), column=5)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "規格", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=3).value = c
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def decide_Pa(cunt, page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    wb2.close()

    loops = 0
    if page < 1:
        pagenumb = cunt
    if page >= 1:
        pagenumb = 30
    page0 = page
    k = 1
    for j in range(1, page + 2):
        for i in range(1, pagenumb + 1):
            wb = openpyxl.load_workbook(BOM_output_path + "catia_bom.xlsx")
            ws = wb.active
            kss = {"What": "Page", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                   "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                   "SearchFormat": False}

            cell = ws.cell(row=(k + 4), column=6)
            c = cell.value
            data_size.append(cell.value)
            k += 1
            # ==========================複製BOM表資料==========================

            # ==========================貼上BOM表資料==========================
            wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")
            ws = wb.get_sheet_by_name("Sheet" + str(j))
            kss1 = {"What": "頁碼", "After": "ActiveCell", "LookIn": "xlFormulas", "LookAt": "xlPart",
                    "SearchOrder": "xlByRows", "SearchDirection": "xlNext", "MatchCase": False, "MatchByte": False,
                    "SearchFormat": False}

            ws.cell(row=(i + 6), column=7).value = cell.value
            wb.save(BOM_output_path + "BOM_空白頁.xlsx")
            # ==========================貼上BOM表資料==========================

            loops += 1
        page0 -= 1
        if page0 == 0:
            pagenumb = cunt - 30 * page

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


def Adjustment(page):
    wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
    time.sleep(1)
    wb2.close()
    wb = openpyxl.load_workbook(BOM_output_path + "BOM_空白頁.xlsx")

    # ==========================調整欄寬至適當大小==========================
    all_ws = wb.sheetnames

    dims = {}
    for ws in all_ws:
        ws = wb[ws]
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column] = max(dims.get(cell.column, 0), len(str(cell.value)))
        print(dims)
    for ws in all_ws:
        ws = wb[ws]
        for col, value in dims.items():
            ws.column_dimensions[get_column_letter(col)].width = value + 3
    dims.clear()
    # ==========================調整欄寬至適當大小==========================
    for i in range(1, page + 2):
        ws = wb.get_sheet_by_name("Sheet" + str(i))
        # # ==========================文字置中==========================
        AtoH = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        for j in AtoH:
            j = str(j)
            for k in range(2, 37):
                k = str(k)
                ws[j + k].font = Font(name=u'標楷體', bold=False, italic=False, size=12)
                ws[j + k].alignment = Alignment(horizontal='center', vertical='center')

        # ==========================更改字型==========================
        wb.save(BOM_output_path + "BOM_空白頁.xlsx")

        wb2 = app.books.open(BOM_output_path + "BOM_空白頁.xlsx")
        time.sleep(1)
        wb2.close()


information_bom(page)

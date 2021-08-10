import csv
import win32com.client as win32
import openpyxl

output_file_root = str()
import_file_root = str()
strip_parameters_file_root = str('C:\\Users\\PDAL\\Desktop\\VB-GTCA022\\strip_parameter.csv')
Mode_status = str('閉模')
input_root = str('C:\\Users\\PDAL\\Desktop\\VB-GTCA022\\auto\\catia_input-GTCA022\\')
# 檔案路徑
file_path = str('C:\\Users\\PDAL\\Desktop\\VB-GTCA022')
# 儲存路徑 (output 零件)
save_path = str(file_path + '\\auto\\catia_output-GTCA022\\')
# 母檔輸入路徑 (input Data)
open_path = str(file_path + "\\auto\\catia_input-GTCA022\\")
# 模具規範路徑
die_rule_path = str(file_path + "\\auto\\die_rule\\")
# 2D出圖路徑
drafting_output_path = str(file_path + "\\auto\\drafting_output-GTCA022\\")
# 標準零件路徑
standard_path = str(file_path + "\\auto\\Standard_Assembly\\")
# 製作一半的BOM表儲存路徑
onwork_BOM_open = str(file_path + "\\BOM表\\")
# BOM表儲存路徑
BOM_output_path = str(file_path + "\\auto\\BOM_output-GTCA022\\")
serch_result = float()
all_part_name = ['']

with open(strip_parameters_file_root) as csvFile:
    rows = csv.reader(csvFile)
    parameter_list = tuple(tuple(rows)[0])
    strip_parameter_list = parameter_list

now_plate_line_number = 1
plate_line_limiting_point_number = 0
F_CB_name = [[None] * 99 for i in range(99)]
F_CB_name[8][2] = "CB_8-30"


def limiting_assembly():  # 限位柱
    catapp = win32.Dispatch('CATIA.Application')

    SAVE_AS_1()  # 更改內限位柱
    SAVE_AS_2()  # 更改螺栓
    limiting()  # 組裝內限位柱
    CB()  # 更改內限位柱

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def SAVE_AS_1():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Documents
    partDocument1 = document.Open(open_path + "limiting_post.CATPart")
    part1 = partDocument1.Part

    # ===============================測試用值===============================
    lower_die_cavity_plate_height = 40
    AS_Length = 40
    AS_Diameter = 20
    # ===============================測試用值===============================

    # ===============================建點1===============================
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(0, 0, lower_die_cavity_plate_height)
    bodies1 = part1.Bodies
    body1 = bodies1.Item("PartBody")
    body1.InsertHybridShape(hybridShapePointCoord1)
    part1.InWorkObject = hybridShapePointCoord1
    # ===============================建點1===============================

    # ===============================建點2===============================
    if lower_die_cavity_plate_height >= 19:
        hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(0, 0, AS_Length - 11 + float(
            strip_parameter_list[1]) - 0.1)
    else:
        hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(0, 0, AS_Length - 8 + float(
            strip_parameter_list[1]) - 0.1)
    body1.InsertHybridShape(hybridShapePointCoord2)
    part1.InWorkObject = hybridShapePointCoord2
    # ===============================建點2===============================

    # ===============================將點隱藏===============================
    selection1 = partDocument1.Selection
    visPropertySet1 = selection1.VisProperties

    selection1.Add(hybridShapePointCoord1)
    selection1.Add(hybridShapePointCoord2)

    visPropertySet1 = visPropertySet1.Parent
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # ===============================將點隱藏===============================

    part1.Update()
    parameters1 = part1.Parameters

    length1 = parameters1.Item("D")  # 改變尺寸 D
    length1.Value = AS_Diameter

    length2 = parameters1.Item("height")  # 改變尺寸 height
    length2.Value = lower_die_cavity_plate_height + float(strip_parameter_list[1]) - 0.1

    # ===============================改變尺寸 H===============================
    length3 = parameters1.Item("d1")
    length4 = parameters1.Item("d2")
    length5 = parameters1.Item("h")
    if AS_Diameter < 19:
        length3.Value = 11
        length4.Value = 7
        length5.Value = 14
    else:
        length3.Value = 15
        length4.Value = 9
        length5.Value = 16
    # ===============================改變尺寸 H===============================

    part1.Update()
    product1 = partDocument1.getItem("Part1")  # 改part名稱
    product1.PartNumber = "limiting_post"

    partDocument1 = catapp.ActiveDocument
    partDocument1.SaveAs(save_path + "limiting_post.CATPart")

    specsAndGeomWindow1 = catapp.ActiveWindow  # 關閉視窗
    specsAndGeomWindow1.Close()


def SAVE_AS_2():
    catapp = win32.Dispatch('CATIA.Application')

    # ===============================測試用值===============================
    AS_Diameter = 20
    AS_Length = 40
    lower_pad_height = 16
    # ===============================測試用值===============================

    ss = int(AS_Length) + 1  # 設定變數

    document = catapp.Documents

    # ================依照內限位柱直徑判斷螺栓使用大小(M8 or M6)===============
    if AS_Diameter >= 19:
        partDocument1 = document.Open(standard_path + "\\Bolt\\CB_8.CATPart")
        qq = 8
        ss = AS_Length - 16 + (lower_pad_height * 1 / 2)
        ss = -int(-ss)
    else:
        partDocument1 = document.Open(standard_path + "\\Bolt\\CB_6.CATPart")
        qq = 6
        ss = AS_Length - 14 + (lower_pad_height * 1 / 2)
        ss = -int(-ss)
    # ================依照內限位柱直徑判斷螺栓使用大小(M8 or M6)===============

    part1 = partDocument1.Part
    part1.Update()

    # =====================螺栓長度===============================
    partDocument1 = catapp.ActiveDocument
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    strParam1 = parameters1.Item("CB_M_L")
    iSize = strParam1.GetEnumerateValuesSize()  # STRING裡面選擇數量
    myArray = [iSize - 1] * 31
    myArray[iSize - 1] = "8-200"
    strParam1.GetEnumerateValues(myArray)  # 抓取STRING的數值放入矩陣之中

    x = 0

    # 測試用數據
    myArray = {1: "8-8", 2: "8-10", 3: "8-12", 4: "8-15", 5: "8-16", 6: "8-18", 7: "8-20", 8: "8-22", 9: "8-25",
               10: "8-30", 11: "8-35", 12: "8-40", 13: "8-45", 14: "8-50", 15: "8-55", 16: "8-60", 17: "8-70",
               18: "8-75", 19: "8-80", 20: "8-85", 21: "8-90", 22: "8-95", 23: "8-100", 24: "8-110", 25: "8-120",
               26: "8-130", 27: "8-140", 28: "8-150", 29: "8-160", 30: "8-200"}

    while ss != 0 and x == 0:
        for j in range(1, iSize):
            limiting_bolt = str(qq) + "-" + str(ss)

            if myArray[j] == limiting_bolt:
                x = str(limiting_bolt)
                F_CB_name[qq][2] = "CB_" + str(qq) + "-" + str(ss)
                part1.Update()
        if x == 0:
            ss -= 1
        else:
            break
    # =====================螺栓長度===============================

    part1.Update()
    product1 = partDocument1.getItem("CB_" + str(qq))
    product1.PartNumber = F_CB_name[qq][2]

    partDocument1.SaveAs(save_path + str(F_CB_name[qq][2]) + ".CATPart")

    # ================關閉視窗================
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    # ================關閉視窗================


def limiting():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    plate_line_number = 1

    for g in range(1, plate_line_number + 1):

        for i in range(1, plate_line_limiting_point_number + 1):
            # ================匯入檔案================
            arrayOfVariantOfBSTR1 = [0] * 9
            arrayOfVariantOfBSTR1[0] = save_path + "limiting_post.CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ================匯入檔案================

            product1 = product1.ReferenceProduct

            # ================進行拘束================
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/limiting_post." + str(i) + "/!Start_point")
            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!plate_line_1_limiting_point_" + str(i))
            constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
            length1 = constraint1.dimension
            length1.Value = 0

            reference3 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + "_" + ".1/!Product1/lower_die_" + str(g) + ".1/")
            constraint2 = constraints1.AddMonoEltCst(0, reference3)
            # ================進行拘束================
            # product1.Update()


def CB():
    catapp = win32.Dispatch('CATIA.Application')

    # ss = int(Form19.Text6.Text) + 1 - 8  # 設定螺栓長度(不包誇螺帽) ss=全長-螺帽
    AS_Diameter = 20

    if AS_Diameter >= 19:  # 判斷使用螺栓大小(M8 or M6)
        qq = 8
        ss = 8  # 沉頭深-頭部厚
    else:
        qq = 6
        ss = 8  # 沉頭深-頭部厚

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    M = 0

    # =====================螺栓判斷(搜尋)===============================
    selection1 = document.Selection
    selection1.Clear()
    selection1.Search("Name=*" + str(F_CB_name[qq][2]) + "_*")
    M = selection1.Count
    selection1.Clear()
    # =====================螺栓判斷(搜尋)===============================

    for i in range(1, plate_line_limiting_point_number + 1):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1 = [0] * 9
        arrayOfVariantOfBSTR1[0] = save_path + str(F_CB_name[qq][2]) + ".CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        M += 1
        productS1 = product1.ReferenceProduct

        # ================進行拘束================
        constraints1 = productS1.Connections("CATIAConstraints")
        reference1 = productS1.CreateReferenceFromName(
            "Product1/" + str(F_CB_name[qq][2]) + "." + str(M) + "_up." + "/!Start_Point")
        reference2 = productS1.CreateReferenceFromName(
            "Product1/limiting_post." + str(i) + "/!End_point")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0

        reference3 = productS1.CreateReferenceFromName(
            "Product1/" + str(F_CB_name[qq][2]) + "." + str(M) + "/!End_point")
        reference4 = productS1.CreateReferenceFromName(
            "Product1/limiting_post." + str(i) + "/!Start_Point")
        constraint2 = constraints1.AddBiEltCst(1, reference3, reference4)
        length2 = constraint2.dimension
        length2.Value = ss - 1.1  # 導角+0.1
        # ================進行拘束================
        # product1.Update()


def hide1():
    catapp = win32.Dispatch('CATIA.Application')


def hide2():
    catapp = win32.Dispatch('CATIA.Application')


def hide3():
    catapp = win32.Dispatch('CATIA.Application')


def hide4():
    catapp = win32.Dispatch('CATIA.Application')


limiting_assembly()

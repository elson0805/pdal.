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

CB_data = [[None] * 4 for i in range(4)]
CB_data[1][1] = "CB_12-45"
CB_data[1][2] = "CB_12-25"
CB_data[1][3] = "CB_12-65"

BoltQuantity = [0] * 4
BoltQuantity[1] = 6
BoltQuantity[2] = 4
BoltQuantity[3] = 4


def Lockingassemble():  # 螺栓組立
    # bending_bolt()
    splint()
    stop_plate()
    lower_die()
    # guide_plate()
    # emboss_forming_punch_left()
    # emboss_forming_punch_right()
    up_plate_Bolt()
    stop_plate_Bolt()
    # half_cut_punch_assemble()
    # Bend_shaping_form_bolt_assemble()
    # hide()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()

    insert_assemble()  # 入子螺栓
    product1.Update()


def bending_bolt():
    g = None


def splint():
    catapp = win32.Dispatch('CATIA.Application')

    # plate_line_number = [] * 99
    # plate_line_number[99] = 1

    plate_line_number = 1

    for g in range(1, plate_line_number + 1):
        M = 0
        # CB_data = [[None] * 99 for i in range(99)]
        # CB_data[1][1] = "CB_12-45"

        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][1]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================

        # BoltQuantity = [0] * 99
        # BoltQuantity[1] = 6
        for i in range(1, BoltQuantity[1] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_" + str(CB_data[1][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][1]) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)

            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][1]) + "." + str(M) + "/!Start_Point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)

            product1.Update()


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')

    # plate_line_number = [] * 99
    # plate_line_number[99] = 1

    plate_line_number = 1

    for g in range(1, plate_line_number + 1):
        M = 0
        # CB_data = [[None] * 99 for i in range(99)]
        # CB_data[1][2] = "CB_12-25"

        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][2]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================

        # BoltQuantity = [0] * 99
        # BoltQuantity[2] = 4
        for i in range(1, BoltQuantity[2] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_" + str(CB_data[1][2]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][2]) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)

            reference4 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][2]) + "." + str(M) + "/!Start_Point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)

            product1.Update()


def lower_die():
    catapp = win32.Dispatch('CATIA.Application')

    # plate_line_number = [] * 99
    # plate_line_number[99] = 1

    plate_line_number = 1

    for g in range(1, plate_line_number + 1):
        M = 0
        # CB_data = [[None] * 99 for i in range(99)]
        # CB_data[1][3] = "CB_12-65"

        # =====================螺栓判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=*Bolt_" + str(CB_data[1][3]) + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================

        # BoltQuantity = [0] * 99
        # BoltQuantity[3] = 4
        for i in range(1, BoltQuantity[3] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_" + str(CB_data[1][3]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1/Stop_plate_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Locking_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][3]) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)

            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Locking_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/Bolt_" + str(CB_data[1][3]) + "." + str(M) + "/!Start_Point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)
            product1.Update()


def guide_plate():
    catapp = win32.Dispatch('CATIA.Application')


def emboss_forming_punch_left():
    catapp = win32.Dispatch('CATIA.Application')


def emboss_forming_punch_right():
    catapp = win32.Dispatch('CATIA.Application')


def up_plate_Bolt():
    catapp = win32.Dispatch('CATIA.Application')

    # =====================數值設定=========================
    CB_length = str(46)
    CB_M = str(8)
    point_name1 = "Start_Point"
    point_name2 = "End_Point"
    # =====================數值設定=========================

    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================

    element_name1 = "Bolt_CB_" + str(CB_M) + "-" + str(CB_length)  # 數值設定

    now_plate_line_number = 2
    for g in range(1, now_plate_line_number + 1):
        up_pad_Bolt_Hole = [g] * 9
        up_pad_Bolt_Hole[g] = 0
        for i in range(1, int(up_pad_Bolt_Hole[g] / 2 + 1)):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + element_name1 + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1)
            # =====================匯入檔案到組立=========================

            for for_start in range(1, 2 + 1):
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name2
                assemble_name2 = str(
                    "Product1/up_plate_" + str(g) + ".1" + "/!up_pad_" + str(g) + "_Bolt_point_" + str(i * 2 - 1))
                assemble(assemble_name1, assemble_name2)
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name1
                assemble_name2 = str(
                    "Product1/up_plate_" + str(g) + ".1" + "/!up_pad_" + str(g) + "_Bolt_point_" + str(i * 2))
                assemble(assemble_name1, assemble_name2)
            CB_Counter1 += 1


def stop_plate_Bolt():
    catapp = win32.Dispatch('CATIA.Application')

    # =====================數值設定=========================
    CB_length = str(int(strip_parameter_list[20]) - 11 + 16)
    CB_M = str(8)
    point_name1 = "Start_Point"
    point_name2 = "End_Point"
    # =====================數值設定=========================

    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================

    element_name1 = "Bolt_CB_" + str(CB_M) + "-" + str(CB_length)  # 數值設定

    now_plate_line_number = 2
    for g in range(1, now_plate_line_number + 1):
        up_pad_Bolt_Hole = [g] * 9
        up_pad_Bolt_Hole[g] = 0
        for i in range(1, int(up_pad_Bolt_Hole[g] / 2 + 1)):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + element_name1 + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1)
            # =====================匯入檔案到組立=========================

            for for_start in range(1, 2):
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name2
                assemble_name2 = str(
                    "Product1/Stop_plate_" + str(g) + ".1" + "/!Stop_plate_" + str(g) + "_Bolt_point_" + str(
                        i * 2 - 1))
                assemble(assemble_name1, assemble_name2)
                assemble_name1 = "Product1/" + element_name1 + "." + CB_Counter1 + "/!" + point_name1
                assemble_name2 = str(
                    "Product1/Stop_plate_" + str(g) + ".1" + "/!Stop_plate_" + str(g) + "_Bolt_point_" + str(i * 2))
                assemble(assemble_name1, assemble_name2)
            CB_Counter1 += 1


def half_cut_punch_assemble():
    catapp = win32.Dispatch('CATIA.Application')


def Bend_shaping_form_bolt_assemble():
    catapp = win32.Dispatch('CATIA.Application')


def hide():
    catapp = win32.Dispatch('CATIA.Application')


def insert_assemble():  # 入子螺栓組立
    catapp = win32.Dispatch('CATIA.Application')

    lower_die_cavity_plate_height = 40  # 測試用數值
    # =====================數值設定=========================
    CB_length = lower_die_cavity_plate_height - 13
    CB_length += 16
    CB_M = 8
    # =====================數值設定=========================
    (CB_length) = lifter_guide_save_CB(CB_M, CB_length)  # 建檔副程式呼叫
    # print(CB_length)

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    # =====================搜尋現有的螺栓數量=========================
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    selection1.Clear()
    selection1.Search("Name=Bolt_CB_" + str(CB_M) + str(CB_length) + "_*")
    CB_Counter1 = selection1.Count
    selection1.Clear()
    CB_Counter1 += 1
    # =====================搜尋現有的螺栓數量=========================

    now_plate_line_number = 1
    for g in range(1, now_plate_line_number + 1):
        pad_Bolt_Hole = [g] * 6
        pad_Bolt_Hole[g] = 7
        for i in range(1, pad_Bolt_Hole[g] + 1):
            # =====================匯入檔案到組立=========================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_CB_" + str(CB_M) + "-" + str(CB_length) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # =====================匯入檔案到組立=========================

            # =====================組立拘束宣告=========================
            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_pad_" + str(g) + ".1/!Product1/lower_pad_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Bolt_CB_" + str(CB_M) + "-" + str(CB_length) + "." + str(CB_Counter1) + "/!Start_Point")
            reference3 = product1.CreateReferenceFromName(
                "Product1/lower_pad_" + str(g) + ".1" + "/!pad_" + str(g) + "_Bolt_point_" + str(i))
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0
            CB_Counter1 += 1
    # =====================組立拘束宣告=========================


def lifter_guide_save_CB(M, length):
    catapp = win32.Dispatch('CATIA.Application')
    documents1 = catapp.Documents

    plate_bolt = 0  # 初始化值

    partDocument1 = documents1.Open(standard_path + "\\Bolt\\CB_" + str(M) + ".CATPart")  # 開啟檔案

    product1 = partDocument1.getItem("CB_" + str(M))
    part1 = partDocument1.Part
    parameters1 = part1.Parameters

    # =====================螺栓長度=========================
    strParam1 = parameters1.Item("CB_M_L")
    iSize = strParam1.GetEnumerateValuesSize()  # STRING 裡面選擇數量
    myArray = [iSize - 1] * 31
    myArray[iSize - 1] = "8-200"
    strParam1.GetEnumerateValues(myArray)  # 抓取 STRING 的數值放入

    # =====================找尋適合的螺栓=========================

    # 測試用數據
    myArray = {1: "8-8", 2: "8-10", 3: "8-12", 4: "8-15", 5: "8-16", 6: "8-18", 7: "8-20", 8: "8-22", 9: "8-25",
               10: "8-30", 11: "8-35", 12: "8-40", 13: "8-45", 14: "8-50", 15: "8-55", 16: "8-60", 17: "8-70",
               18: "8-75", 19: "8-80", 20: "8-85", 21: "8-90", 22: "8-95", 23: "8-100", 24: "8-110", 25: "8-120",
               26: "8-130", 27: "8-140", 28: "8-150", 29: "8-160", 30: "8-200"}

    while length != 0 and plate_bolt == 0:
        length = int(length)
        length -= 1
        plate_bolt_test_name_1 = str(M) + "-" + str(length)
        for Array_count in range(1, iSize):
            if myArray[Array_count] == plate_bolt_test_name_1:
                plate_bolt = plate_bolt_test_name_1
    # =====================找尋適合的螺栓=========================

    CB_name = "CB_" + str(plate_bolt_test_name_1)
    # =====================螺栓長度=========================

    # =====================找尋現有螺栓的尺寸=========================
    # p = save_path + "Bolt_" + CB_name + ".*"
    p = ""
    while p != "":
        if p == "Bolt_" + CB_name + ".CATPart":
            part1.Updeta()
            partDocument1.Close()
    # =====================找尋現有螺栓的尺寸=========================

    # =====================參數宣告及變更=========================
    strParam1 = parameters1.Item("CB_M_L")  # 參數宣告
    strParam1.Value = plate_bolt  # 變更
    product1.PartNumber = "Bolt_" + CB_name  # 改part名字(非檔名)
    part1.Update()
    partDocument1.SaveAs(save_path + "Bolt_" + CB_name + ".CATPart")
    partDocument1.Close()
    # =====================參數宣告及變更=========================
    return length  # 回傳數值


def assemble(assemble_name1, assemble_name2):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    # =====================組立拘束宣告=========================
    constraints1 = product1.Connections("CATIAConstraints")
    reference1 = product1.CreateReferenceFromName(assemble_name1)
    reference2 = product1.CreateReferenceFromName(assemble_name2)
    constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
    length1 = constraint1.dimension
    length1.Value = 0
    # =====================組立拘束宣告=========================
    product1.Update()


Lockingassemble()

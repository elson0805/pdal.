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
g = now_plate_line_number
total_op_number = 9
now_op_number = 7
n = now_op_number

bolt_name = [0] * 9

insert_interferance_count = [[[0] * 6 for j in range(6)] * 6 for k in range(6)]

plate_line_allotype_cut_line_number = [[0] * 99 for i in range(99)]
plate_line_allotype_cut_line_number[1][2] = 0
plate_line_allotype_cut_line_number[1][3] = 0
plate_line_allotype_cut_line_number[1][4] = 0
plate_line_allotype_cut_line_number[1][6] = 0
plate_line_allotype_cut_line_number[1][7] = 0


def allotype_punch_assemble():  # 異型沖頭
    catapp = win32.Dispatch('CATIA.Application')

    Standard_Part()  # 標準零件

    for now_op_number in range(1, total_op_number + 1):
        for g in range(1, 2):
            n = now_op_number
            op_number = 10 * n
            plate_line_allotype_cut_line_number = [[0] * 99 for i in range(99)]
            plate_line_allotype_cut_line_number[g][n] = 0
            if plate_line_allotype_cut_line_number[g][n] > 0:
                assemble(now_op_number, op_number, total_op_number)  # 組立

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def Standard_Part():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Documents
    for i in range(1, 2 + 1):
        a = [1] * 3
        a[1] = 8
        a[2] = 8
        partDocument1 = document.Open(standard_path + "\\Bolt\\CB_" + str(a[i]) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        # ===============================螺栓長度===============================
        strParam1 = parameters1.Item("CB_M_L")
        iSize = strParam1.GetEnumerateValuesSize()  # STRING裡面選擇數量
        myArray = [iSize - 1] * 99
        myArray[iSize - 1] = "8-200"
        strParam1.GetEnumerateValues(myArray)  # 抓取STRING的數值放入矩陣之中

        plate_height = [1] * 3
        plate_height[1] = int(float(strip_parameter_list[20]) - 11) + 8 * 3
        plate_height[2] = int(float(strip_parameter_list[14]) - 11) + 28

        b = plate_height[i]
        # 測試用數據

        bolt_name = [0] * 9

        myArray = {1: "8-8", 2: "8-10", 3: "8-12", 4: "8-15", 5: "8-16", 6: "8-18", 7: "8-20", 8: "8-22", 9: "8-25",
                   10: "8-30", 11: "8-35", 12: "8-40", 13: "8-45", 14: "8-50", 15: "8-55", 16: "8-60", 17: "8-70",
                   18: "8-75", 19: "8-80", 20: "8-85", 21: "8-90", 22: "8-95", 23: "8-100", 24: "8-110",
                   25: "8-120", 26: "8-130", 27: "8-140", 28: "8-150", 29: "8-160", 30: "8-200"}

        while b != 0 and i != 0:
            for j in range(1, iSize):
                bolt_name[1] = str(a[i]) + "-" + str(b)
                bolt_name[2] = str(a[i]) + "-" + str(b)

                if myArray[j] == bolt_name[i]:
                    strParam1.Value = bolt_name[i]
                    part1.Update()
                    break
            if strParam1.Value != bolt_name[i]:
                b -= 1
            else:
                break
        # ===============================螺栓長度===============================

        partDocument1 = catapp.ActiveDocument
        product1 = partDocument1.getItem("CB_" + str(a[i]))  # 改part名稱
        product1.PartNumber = "Bolt_CB_" + str(bolt_name[i])  # 改樹狀圖名稱
        partDocument1.SaveAs(save_path + "Bolt_CB_" + str(bolt_name[i]) + ".CATPart")

        specsAndGeomWindow1 = catapp.ActiveWindow  # 關閉視窗
        specsAndGeomWindow1.Close()


def assemble(now_op_number, op_number, total_op_number):
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    for j in range(1, plate_line_allotype_cut_line_number[g][n] + 1):
        insert_interferance_no_delete = 0
        insert_interferance_decide_Excavation = 0
        for ii in range(1, total_op_number + 1):
            for qq in range(1, 10 + 1):
                insert_interferance_count[ii][qq][1] = 0
                insert_interferance_count[ii][qq][2] = 0
                if insert_interferance_count[ii][qq][1] == n and insert_interferance_count[ii][qq][2] == j:
                    insert_interferance_now = ii
                    insert_interferance_decide_Excavation = 1
                    if qq > 1:
                        break

        # ===============================進行拘束==============================
        constraints1 = product1.Connections("CATIAConstraints")
        reference1 = product1.CreateReferenceFromName(
            "GTA022\\op" + str(op_number) + "_allotype_QR_Stripper_insert_0" + str(j) + ".1\\!GTA022\\op" + str(
                op_number) + "_allotype_QR_Stripper_insert_0" + str(j) + ".1\\")
        constraint1 = constraints1.AddMonoEltCst(0, reference1)
        # ===============================進行拘束==============================

        ffff = 4
        gggg = 2
        if insert_interferance_decide_Excavation == 0:
            ffff = 2
            gggg = 1
        for k in range(1, ffff):
            # ===============================搜尋上夾板入子螺栓==============================
            selection2 = document.Selection
            selection2.Clear()
            selection2.Search("name='Bolt_CB_" + str(bolt_name[1]) + "'.*, All ")
            M = selection2.Count + 1
            selection2.Clear()
            # ===============================搜尋上夾板入子螺栓==============================

            # ===============================匯入檔案==============================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_CB_" + str(bolt_name[2]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ===============================匯入檔案==============================\

            # ===============================進行拘束==============================
            reference2 = product1.CreateReferenceFromName(
                "GTA022/op" + str(op_number) + "_allotype_QR_Splint_insert_0" + str(j) + ".1/!Start_Point" + str(
                    gggg) + "_" + str(k))
            reference3 = product1.CreateReferenceFromName(
                "GTA022/Bolt_CB_" + str(bolt_name[2]) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0

            reference2 = product1.CreateReferenceFromName(
                "GTA022/op" + str(op_number) + "_allotype_QR_Splint_insert_0" + str(j) + ".1/!End_Point" + str(
                    gggg) + "_" + str(k))
            reference3 = product1.CreateReferenceFromName(
                "GTA022/Bolt_CB_" + str(bolt_name[2]) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0
            # ===============================進行拘束==============================

            # ===============================脫料板入子螺栓==============================
            selection2 = document.Selection
            selection2.Clear()
            selection2.Search("name='Bolt_CB_" + str(M) + "'.*, All ")
            M = selection2.Count + 1
            selection2.Clear()
            # ===============================脫料板入子螺栓==============================

            # ===============================匯入檔案==============================
            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + "Bolt_CB_" + str(bolt_name[1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
            # ===============================匯入檔案==============================\

            # ===============================進行拘束==============================
            reference2 = product1.CreateReferenceFromName(
                "GTA022/op" + str(op_number) + "_allotype_QR_Stripper_insert_0" + str(j) + ".1/!Start_Point" + str(
                    gggg) + "_" + str(k))
            reference3 = product1.CreateReferenceFromName(
                "GTA022/Bolt_CB_" + str(bolt_name[1]) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0

            reference2 = product1.CreateReferenceFromName(
                "GTA022/op" + str(op_number) + "_allotype_QR_Stripper_insert_0" + str(j) + ".1/!End_Point" + str(
                    gggg) + "_" + str(k))
            reference3 = product1.CreateReferenceFromName(
                "GTA022/Bolt_CB_" + str(bolt_name[1]) + "." + str(M) + "/!End_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0
            # ===============================進行拘束==============================


allotype_punch_assemble()

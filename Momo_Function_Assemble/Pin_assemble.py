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

Pin_data = [[None] * 99 for i in range(99)]
Pin_data21 = "MSTH_10-50"
Pin_data[2][1] = Pin_data21
Pin_data22 = "LP_10-40"
Pin_data[2][2] = Pin_data22
Pin_data23 = "MSTH_10-60"
Pin_data[2][3] = Pin_data23


PinQuantity = [0] * 99
PinQuantity[1] = 4
PinQuantity[2] = 4
PinQuantity[3] = 4


def Pin_assemble():  # 和銷
    splint()
    stop_plate()
    lower_die()
    # guide_plate()
    # hide()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def splint():
    catapp = win32.Dispatch('CATIA.Application')

    # a = Form19.Combo2 & "_" & Pin_data(1, 1)  # pin直徑
    b = "Pin_" + str(Pin_data[2][1])  # pin長度

    plate_line_number = 1
    for g in range(1, plate_line_number + 1):
        M = 0

        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================

        for i in range(1, PinQuantity[1] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0

            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)

            WordCount_PinLength = len(Pin_data[2][1])
            for j in range(0, WordCount_PinLength):
                word = Pin_data21[j]  # 提取Pin_data[2][1]中的值
                if word == "1":
                    length2 = constraint3.dimension
                    length2.Value = int(Pin_data21[j + 3:10]) - float(strip_parameter_list[14]) * 1 / 2
            product1.Update()


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')

    # a = "LP_" + Pin_data[1, 2]  # pin直徑
    b = "Stripper_pin_" + str(Pin_data[2][2])  # 帶頭合銷_型號_直徑-長度

    plate_line_number = 1
    for g in range(1, plate_line_number + 1):

        M = 0

        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================

        for i in range(1, PinQuantity[2] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Product1/Stop_plate_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 5

            reference4 = product1.CreateReferenceFromName(
                "Product1/Stop_plate_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length2 = constraint3.dimension
            length2.Value = 0

            # WordCount_PinLength = len(Pin_data[2][2])
            # for j in range(0, WordCount_PinLength):
            #     word = Pin_data22[j]  # 提取Pin_data[2][1]中的值
            #     if word == "1":
            #         length2 = constraint3.dimension
            #         length2.Value = int(Pin_data22[j + 3:10]) - float(strip_parameter_list[17]) * 1 / 2
            product1.Update()


def lower_die():
    catapp = win32.Dispatch('CATIA.Application')

    plate_line_number = 1
    # a = Form19.Combo2 + Pin_data[1, 3]
    b = "Pin_" + str(Pin_data[2][3])

    for g in range(1, plate_line_number + 1):

        M = 0

        # =====================pin判斷(搜尋)===============================
        partdoc = catapp.ActiveDocument
        selection1 = partdoc.Selection
        selection1.Clear()
        selection1.Search("Name=" + b + "_*")
        M = selection1.Count
        selection1.Clear()
        # =====================pin判斷(搜尋)===============================

        for i in range(1, PinQuantity[3] + 1):
            M += 1

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1/lower_die_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Pin_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0

            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Pin_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + b + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)

            WordCount_PinLength = len(Pin_data[2][3])
            for j in range(0, WordCount_PinLength):
                word = Pin_data23[j]  # 提取Pin_data[2][1]中的值
                if word == "1":
                    lower_die_cavity_plate_height = 40
                    length2 = constraint3.dimension
                    length2.Value = int(Pin_data23[j + 3:10]) - lower_die_cavity_plate_height * 1 / 2
            product1.Update()


def guide_plate():
    catapp = win32.Dispatch('CATIA.Application')


def hide():
    catapp = win32.Dispatch('CATIA.Application')


Pin_assemble()

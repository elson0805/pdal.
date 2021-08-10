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

plate_line_number = 1

InnerGuidingQuantity = [0] * 9
InnerGuidingQuantity[1] = 4

Inner_Guiding_data = [[None] * 9 for i in range(9)]
Inner_Guiding_data[1][1] = 20
Inner_Guiding_data[2][1] = 100


def Inner_Guiding_post_assemble():  # 內導柱/套
    splint()
    Stripper()
    lower_die()
    # hide()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def splint():
    catapp = win32.Dispatch('CATIA.Application')

    M = 0

    for g in range(1, plate_line_number + 1):

        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1

            Inner_Guiding_Post_Material = "SGPH"
            Inner_Guiding_Post_Diameter = 20

            a = Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_Post_Diameter)
            b = str(Inner_Guiding_data[2][1])

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + a + "-" + b + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Product1/Splint_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Inner_Guiding_Post_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + a + "." + str(M) + "/!Inner_Guiding_Post_dir_point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0.3

            reference4 = product1.CreateReferenceFromName(
                "Product1/Splint_" + str(g) + ".1/!Inner_Guiding_Post_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + a + "." + str(M) + "/!Inner_Guiding_Post_point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint3.dimension
            length1.Value = 0


def Stripper():
    catapp = win32.Dispatch('CATIA.Application')

    M = 0

    for g in range(1, plate_line_number + 1):

        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1

            Under_Inner_Guiding_Post_Material = "SGFZ"
            a = Under_Inner_Guiding_Post_Material + "_" + str(Inner_Guiding_data[1][1])
            b = 20  # Inner_Guiding_Post_Bush_up_data[2][1]

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + str(a) + "-" + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Product1/Stripper_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Inner_Guiding_Post_Bush_up_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(a) + "." + str(M) + "/!Start_Point")
            constraint2 = constraints1.AddBiEltCst(1, reference2, reference3)
            length1 = constraint2.dimension
            length1.Value = 0.3

            reference4 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Inner_Guiding_Post_Bush_up_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(a) + "." + str(M) + "/!End_Point")
            constraint3 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint3.dimension
            length1.Value = 0


def lower_die():
    catapp = win32.Dispatch('CATIA.Application')

    M = InnerGuidingQuantity[1]

    for g in range(1, plate_line_number + 1):

        for i in range(1, InnerGuidingQuantity[1] + 1):
            M += 1
            a = Inner_Guiding_data[1][1]
            b = 20  # Inner_Guiding_Post_Bush_up_data[2, 2]
            Under_Inner_Guiding_Post_Material = "SGFZ"

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = \
                save_path + Under_Inner_Guiding_Post_Material + "_" + str(a) + "-" + str(b) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Product1//Stop_plate_" + str(g) + ".1//")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Inner_Guiding_Post_Bush_down_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + Under_Inner_Guiding_Post_Material + "_" + str(a) + "." + str(M) + "/!Start_Point")

            reference4 = product1.CreateReferenceFromName(
                "Product1/lower_die_" + str(g) + ".1/!Inner_Guiding_Post_Bush_down_dir_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + Under_Inner_Guiding_Post_Material + "_" + str(a) + "." + str(M) + "/!End_Point")

            constraint2 = constraints1.AddBiEltCst(1, reference4, reference5)
            length1 = constraint2.dimension
            length1.Value = 0.3
            constraint3 = constraints1.AddBiEltCst(1, reference2, reference3)
            length2 = constraint3.dimension
            length2.Value = 0


def hide():
    catapp = win32.Dispatch('CATIA.Application')


Inner_Guiding_post_assemble()

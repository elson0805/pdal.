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

SBT_data = [[None] * 99 for i in range(99)]
SBT_data[1][1] = 16
SBT_data[2][1] = 60
SBT_data[7][1] = "Shoulder_Screw_SBT_16-60"


def SBT_assemble():  # 螺栓組立
    # stop_plate()
    Stripper()
    # hide()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')


def Stripper():
    catapp = win32.Dispatch('CATIA.Application')

    M = 0

    for g in range(1, plate_line_number + 1):
        # =====================螺栓判斷(搜尋)===============================
        # partdoc = catapp.ActiveDocument
        # selection1 = partdoc.Selection
        # selection1.Clear()
        # selection1.Search("Name=*" + Pin_Hole_ + "_*")
        # M = selection1.Count
        # selection1.Clear()
        # =====================螺栓判斷(搜尋)===============================

        for i in range(1, 2 + 1):
            M += 1

            a = SBT_data[1][1]
            b = SBT_data[2][1]

            # document = catapp.Documents
            # partDocument1 = document.Open(
            #     "C:\\Users\\PDAL\\Desktop\\auto\\Standard_Assembly\\MSTP " + str(a) + "-" + str(b) + ".CATPart")
            # part1 = partDocument1.part
            # length1 = part1.Parameters.Item("T")
            # g = now_plate_line_number
            # length1.Value = Bolt_data[2][1]

            document = catapp.ActiveDocument
            product1 = document.Product
            products1 = product1.Products

            arrayOfVariantOfBSTR1 = [0]
            arrayOfVariantOfBSTR1[0] = save_path + str(SBT_data[7][1]) + ".CATPart"
            products1Variant = products1
            products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")

            constraints1 = product1.Connections("CATIAConstraints")
            reference1 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!Product1/Stripper_" + str(g) + ".1/")
            constraint1 = constraints1.AddMonoEltCst(0, reference1)

            reference2 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!SBT_dir1_point_" + str(i))
            reference3 = product1.CreateReferenceFromName(
                "Product1/" + str(SBT_data[7][1]) + "." + str(M) + "/!SBT_dir_point")
            constraint2 = constraints1.AddBiEltCst(2, reference2, reference3)

            reference4 = product1.CreateReferenceFromName(
                "Product1/Stripper_" + str(g) + ".1/!SBT_dir2_point_" + str(i))
            reference5 = product1.CreateReferenceFromName(
                "Product1/" + str(SBT_data[7][1]) + "." + str(M) + "/!SBT_point")
            constraint3 = constraints1.AddBiEltCst(2, reference4, reference5)


def hide():
    catapp = win32.Dispatch('CATIA.Application')


SBT_assemble()

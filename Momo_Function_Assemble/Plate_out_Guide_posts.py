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
out_Guide_Material = "MYJP"
outer_Guiding_data = [[None] * 99 for i in range(99)]
outer_Guiding_data[1][1] = "32"
outer_Guiding_data[2][1] = "90"
outer_Guiding_data[3][1] = "MYJP"

arrayOfVariantOfBSTR1 = [0] * 9

out_Guide_posts = [0] * 9

out_Guide_Diameter = 32
out_Guide_Length = 90


def Plate_out_Guide_posts():  # 外導柱/套組立
    if out_Guide_Material == "DANLY":
        Guide_assemble()
    else:
        Post_down()
        # hide()
        Post_UP()
        # hide1()

    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products
    product1.Update()


def Post_down():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    for i in range(1, 4 + 1):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1 = [0]
        arrayOfVariantOfBSTR1[0] = \
            save_path + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "-" + str(
                outer_Guiding_data[2][1]) + "_down.CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        constraints1 = product1.Connections("CATIAConstraints")

        # ================進行拘束================
        # reference1 = product1.CreateReferenceFromName("Product1\\MYKP_32_120_DOWN." + str(i) + "\\!MYKP_Pin_point_1")
        # reference1 = product1.CreateReferenceFromName(
        #     "Product1\\" + Form19.Combo10 + "_" + Form19.Combo12 + "-" + Form19.Combo13 + "_down." + str(
        #         i) + "\\!Start_Point")
        reference1 = product1.CreateReferenceFromName(
            "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/lower_die_set.1/!Pillar_center_" + str(i) + "_pin_point_1")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0

        # reference3 = product1.CreateReferenceFromName("Product1\\MYKP_32_120_DOWN." + str(i) + "\\!MYKP_Pin_point_2")
        # reference3 = product1.CreateReferenceFromName(
        #     "Product1\\" + Form19.Combo10 + "_" + Form19.Combo12 + "-" + Form19.Combo13 + "_down." + str(
        #         i) + "\\!End_Point")
        reference3 = product1.CreateReferenceFromName(
            "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                i) + "/!End_Point")
        reference4 = product1.CreateReferenceFromName(
            "Product1/lower_die_set.1/!Pillar_center_" + str(i) + "_pin_point_2")
        constraint2 = constraints1.AddBiEltCst(1, reference3, reference4)
        length2 = constraint2.dimension
        length2.Value = 0

        reference5 = product1.CreateReferenceFromName(
            "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_down." + str(
                i) + "/!down_plane")
        reference6 = product1.CreateReferenceFromName("Product1/lower_die_set.1/!up_plane")
        constraint3 = constraints1.AddBiEltCst(1, reference5, reference6)
        length3 = constraint3.dimension
        length3.Value = 0
        constraint3.Orientation = 0
        # ================進行拘束================
        product1.Update()


def Post_UP():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    for i in range(1, 4 + 1):
        # ================匯入檔案================
        # arrayOfVariantOfBSTR1[0] = standard_path + "MYKP_32_120_UP.CATPart"
        arrayOfVariantOfBSTR1 = [0]
        arrayOfVariantOfBSTR1[0] = \
            save_path + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "-" + str(
                outer_Guiding_data[2][1]) + "_up.CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        constraints1 = product1.Connections("CATIAConstraints")
        reference1 = product1.CreateReferenceFromName(
            "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/upper_die_set.1/!Pillar_center_" + str(i) + "_pin_point_1")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0

        reference3 = product1.CreateReferenceFromName(
            "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
                i) + "/!End_Point")
        reference4 = product1.CreateReferenceFromName(
            "Product1/upper_die_set.1/!Pillar_center_" + str(i) + "_pin_point_2")
        constraint2 = constraints1.AddBiEltCst(1, reference3, reference4)
        length2 = constraint2.dimension
        length2.Value = 0

        # reference5 = product1.CreateReferenceFromName(
        #     "Product1/" + str(Form19.Combo10) + "_" + str(Form19.Combo12) + "_down.")
        # reference6 = product1.CreateReferenceFromName("Product1/lower_die_set.1/!up_plane")
        # constraint3 = constraints1.AddBiEltCst(1, reference5, reference6)
        # length3 = constraint3.dimension
        # length3.Value = 0
        # constraint3.Orientation = catCstOrientSame

        # reference7 = product1.CreateReferenceFromName(
        #     "Product1/" + str(outer_Guiding_data[3][1]) + "_" + str(outer_Guiding_data[1][1]) + "_up." + str(
        #         i) + "/!Point")
        # reference8 = product1.CreateReferenceFromName(
        #     "Product1/upper_die_set.1/!Pillar_center_" + str(i))
        # constraint4 = constraints1.AddBiEltCst(1, reference7, reference8)
        # length4 = constraint4.dimension
        # length4.Value = 0
        # ================進行拘束================
        # product1.Update()


def Guide_assemble():
    SAVE_DANLY()  # 改變導柱長度
    DANLY()  # 組裝導柱
    SAVE_Ball_Bushing()  # 改變鋼珠導套
    Ball_Bushing()  # 組裝鋼珠導套
    SAVE_outer_bush()  # 改變下模板導套
    outer_bush()  # 組裝下模板導套
    save_as_CB()  # 改變+組裝螺栓


def SAVE_DANLY():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Document
    partDocument1 = document.Open(
        standard_path + "\\DANLT\\DANLT-" + str(out_Guide_Diameter) + ".CATPart")
    part1 = partDocument1.Part
    part1.Updeta()

    # ================改變尺寸 L================
    parameters1 = part1.Parameters
    strParam1 = parameters1.Item("D-L")
    strParam1.Value = out_Guide_Diameter + int("-") + out_Guide_Length

    part1.Updeta()

    product1 = partDocument1.getItem("DANLT-" + str(out_Guide_Diameter))
    product1.PartNumber = "DANLT-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length)
    # ================改變尺寸 L================

    # ================另存檔案================
    document1 = catapp.ActiveDocument
    document1.SaveAs(
        save_path + "DANLT-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + ".CATPart")
    # ================另存檔案================

    # ================關閉視窗================
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    document1.Close()
    # ================關閉視窗================


def DANLY():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Document
    product1 = document.Product
    products1 = product1.Products

    for i in range(1, 4 + 1):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1 = [0]
        arrayOfVariantOfBSTR1[0] = \
            save_path + "DANLT-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + ".CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        product1 = product1.ReferenceProduct
        constraints1 = product1.Connections("CATIAConstraints")

        # ================進行拘束================
        reference1 = product1.CreateReferenceFromName(
            "Product1/DANLT-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + "." + str(
                i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName("Product1/upper_die_set.1/!Pillar_center_" + str(i))
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0
        # ================進行拘束================
        # product1.Update()


def SAVE_Ball_Bushing():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Document
    partDocument1 = document.Open(
        standard_path + "\\DANLT\\Ball Bushing-" + str(out_Guide_Diameter) + ".CATPart")

    part1 = partDocument1.Part

    # ================改變尺寸 L================
    parameters1 = part1.Parameters
    strParam1 = parameters1.Item("D-L")
    strParam1.Value = out_Guide_Diameter + int("-120")

    part1.Updeta()

    product1 = partDocument1.getItem("Ball Bushing-" + str(out_Guide_Diameter))
    product1.PartNumber = "Ball Bushing-" + str(out_Guide_Diameter) + "-120"
    # ================改變尺寸 L================

    # ================另存檔案================
    document1 = catapp.ActiveDocument
    document1.SaveAs(
        save_path + "DANLT-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + ".CATPart")
    # ================另存檔案================

    # ================關閉視窗================
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    document1.Close()
    # ================關閉視窗================


def Ball_Bushing():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    for i in range(1, 4 + 1):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1[0] = \
            save_path + "Ball Bushing-" + str(out_Guide_Diameter) + "-120.CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        product1 = product1.ReferenceProduct
        constraints1 = product1.Connections("CATIAConstraints")

        # ================進行拘束================
        reference1 = product1.CreateReferenceFromName(
            "Product1/Ball Bushing-" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + "." + str(
                i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/DANLT" + str(out_Guide_Diameter) + "-" + str(out_Guide_Length) + "." + str(i) + "/!CB_Point")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0
        # ================進行拘束================
        # product1.Update()


def SAVE_outer_bush():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Document
    partDocument1 = document.Open(
        standard_path + "\\DANLT\\outer bush-" + str(out_Guide_Diameter) + ".CATPart")

    part1 = partDocument1.Part
    part1.Update()

    # ================改變尺寸 L================
    parameters1 = part1.Parameters
    strParam1 = parameters1.Item("D-L")
    strParam1.Value = out_Guide_Diameter + int("-55.1")

    part1.Updeta()

    product1 = partDocument1.getItem("outer bush-" + str(out_Guide_Diameter))
    product1.PartNumber = "outer bush-" + str(out_Guide_Diameter) + "-" + str(out_Guide_posts[2])
    # ================改變尺寸 L================

    # ================另存檔案================
    document1 = catapp.ActiveDocument
    document1.SaveAs(
        save_path + "outer bush--" + str(out_Guide_Diameter) + "-" + str(out_Guide_posts[2]) + ".CATPart")
    # ================另存檔案================

    # ================關閉視窗================
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    document1.Close()
    # ================關閉視窗================


def outer_bush():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.ActiveDocument
    product1 = document.Product
    products1 = product1.Products

    for i in range(1, 4):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1[0] = \
            save_path + "outer bush-" + str(out_Guide_Diameter) + str(out_Guide_posts[2]) + "-120.CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        product1 = product1.ReferenceProduct
        constraints1 = product1.Connections("CATIAConstraints")

        # ================進行拘束================
        reference1 = product1.CreateReferenceFromName(
            "Product1/outer bush-" + str(out_Guide_Diameter) + "-" + str(out_Guide_posts[2]) + "." + str(
                i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/lower_die_set.1/!Pillar_center_" + str(i))
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0
        # ================進行拘束================
        # product1.Update()


def save_as_CB():
    catapp = win32.Dispatch('CATIA.Application')

    document = catapp.Document
    partDocument1 = document.Open(open_path + "CB_8.CATPart")

    part1 = partDocument1.Part
    part1.Update()

    # ================改變尺寸 L================
    parameters1 = part1.Parameters
    strParam1 = parameters1.Item("CB_M_L")
    strParam1.Value = out_Guide_Diameter + int("8-25")

    part1.Updeta()

    product1 = partDocument1.getItem("CB_" + qq)
    product1.PartNumber = "CB_8-25"
    # ================改變尺寸 L================

    # ================另存檔案================
    document1 = catapp.ActiveDocument
    document1.SaveAs(save_path + "CB_8-25.CATPart")
    # ================另存檔案================

    # ================關閉視窗================
    specsAndGeomWindow1 = catapp.ActiveWindow
    specsAndGeomWindow1.Close()
    document1.Close()
    # ================關閉視窗================

    # ================螺栓組裝================
    product1 = document1.Product
    products1 = product1.Products
    for i in range(1, 4 + 1):
        # ================匯入檔案================
        arrayOfVariantOfBSTR1[0] = save_path + "CB_8-25.CATPart"
        products1Variant = products1
        products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
        # ================匯入檔案================

        product1 = product1.ReferenceProduct
        constraints1 = product1.Connections("CATIAConstraints")

        # ================進行拘束================
        reference1 = product1.CreateReferenceFromName(
            "Product1/CB_8-25" + "." + str(i) + "/!Start_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/Ball Bushing-" + str(out_Guide_Diameter) + "-" + str(out_Guide_posts[1]) + "." + str(
                i) + "/!End_Point")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0

        reference1 = product1.CreateReferenceFromName("Product1/CB_8-25." + str(i) + "/!End_Point")
        reference2 = product1.CreateReferenceFromName(
            "Product1/Ball Bushing-" + str(out_Guide_Diameter) + "-" + str(out_Guide_posts[1]) + "." + str(
                i) + "/!CB_Point")
        constraint1 = constraints1.AddBiEltCst(1, reference1, reference2)
        length1 = constraint1.dimension
        length1.Value = 0
        # ================進行拘束================
        # product1.Update()
    # ================螺栓組裝================


def Bend_shaping_form_bolt_assemble():
    catapp = win32.Dispatch('CATIA.Application')


def hide():
    catapp = win32.Dispatch('CATIA.Application')


def hide1():
    catapp = win32.Dispatch('CATIA.Application')


Plate_out_Guide_posts()

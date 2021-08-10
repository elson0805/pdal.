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


def PunchMaking():
    now_plate_line_number = 1
    g = now_plate_line_number
    total_op_number = 7
    plate_line_bending_punch_surface = [[0] * 99 for i in range(99)]
    plate_line_bending_punch_surface[1][2] = 0
    plate_line_bending_punch_surface[1][3] = 0
    plate_line_bending_punch_surface[1][4] = 0
    plate_line_bending_punch_surface[1][6] = 0
    plate_line_bending_punch_surface[1][7] = 1

    for now_op_number in range(1, total_op_number + 1):
        n = now_op_number
        op_number = 10 * n
        if plate_line_bending_punch_surface[g][n] > 0:
            for now_data_number in range(1, plate_line_bending_punch_surface[g][n] + 1):  #
                part_open('Data1.CATPart')
                interferance_pad_name = ""
                interferance_line_name = "_bending_punch_surface_"
                open_file_name = "bending_punch"
                punch(open_file_name)
                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number, op_number,
                             now_data_number, total_op_number)


def punch(open_file_name):
    part_name = open_file_name + '.CATPart'
    part_open(part_name)
    window_change('Data1.CATPart', part_name)


def punch_change(InterferancePadName, InterferanceLineName, NowPlateLineNumber, OpNumber, NowDataNumber,
                 TotalOpNumber):  # 模具名稱,沖頭線段,線段編號,op號碼,第幾條
    F_function_environment("bending_punch", "die")  # 環境設定
    all_part_number = 0
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = element_Document1.Part

    # ======XY平台宣告======
    originElements1 = part1.OriginElements
    hybridShapePlaneExplicit1 = originElements1.PlaneXY
    # ======XY平台宣告======

    # ======參數宣告======
    parametersParent = part1.Parameters
    parameters1 = parametersParent.RootParameterSet.ParameterSets.getItem("bending_punch_parameter")
    parameter1 = parameters1.DirectParameters.Item("down_plane")
    parameter2 = parameters1.DirectParameters.Item("high_plane")
    parameter3 = parameters1.DirectParameters.Item("length")
    parameter4 = parameters1.DirectParameters.Item("wide")
    parameter5 = parameters1.DirectParameters.Item("gap")
    parameter6 = parameters1.DirectParameters.Item("pad_high")
    length1 = parameter1
    length2 = parameter2
    length3 = parameter3
    length4 = parameter4
    length5 = parameter5
    length6 = parameter6
    # ======參數宣告======

    # ========設定沖頭高度==========
    if Mode_status == '開模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]) + float(strip_parameter_list[14]) + 28)
    elif Mode_status == '閉模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]) + float(strip_parameter_list[14]))
    # ========設定沖頭高度==========
    length2.value = -30
    # ========草圖置換==========
    bending_surface_formula_1_name = (str(
        "`die\\plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + str(InterferanceLineName) + str(
            NowDataNumber) + '`'))
    relaitons = part1.Relations
    formula1 = relaitons.Item("bending_surface_formula_1")
    formula1.Modify(bending_surface_formula_1_name)
    formula1.Rename(bending_surface_formula_1_name)
    part1.Update()
    # ========草圖置換==========

    hybridShape1 = element_hybridBody1.HybridShapes.Item(str(
        "plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + str(InterferanceLineName) + str(
            NowDataNumber)))  # 宣告元素
    hybridShape2 = element_hybridBody1.HybridShapes.Item("lower_die_seat_line")  # 下模座的線段
    hybridShape5 = element_hybridBody1.HybridShapes.Item("plate_line_1_op10_A_punch_1")  # 宣告元素
    hybridShape6 = element_hybridBody1.HybridShapes.Item("plate_line_1_op10_A_punch_2")  # 宣告元素
    hybridShape3 = element_body1.HybridShapes.Item("up_plane")  # 宣告平面
    hybridShape9 = element_body1.HybridShapes.Item("down_plane")  # 宣告平面
    hybridShape4 = element_body1.HybridShapes.Item("X_Y_min")
    hybridShape7 = element_body1.HybridShapes.Item("center_location")
    hybridShape10 = element_body1.HybridShapes.Item("Z_min")
    F_function_original_point(hybridShape2, NowDataNumber)  # element_Reference(1) 為原點陳述式
    element_point1 = element_Reference1
    basic_element1 = element_Reference1
    product1 = document.GetItem("Strip_Data-2")
    F_function_build_sketch("position_sketch", hybridShape3)  # (sketch名稱,產生sketch的平面)  element_sketch(1)為產生出來的草圖
    part1.InWorkObject = element_sketch1
    part1.Update()
    reference2 = part1.CreateReferenceFromObject(element_point1)
    factory2D1 = element_sketch1.OpenEdition()
    element_point2 = element_point1
    element_point3 = hybridShape4
    element_sketch1.CloseEdition()
    F_function_sketch_build_callout(element_sketch1, "Horizontal", "Callout", H_dir, 2)
    #  element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    F_function_sketch_build_callout(element_sketch1, "Vertical", "Callout", V_dir, 2)
    #  element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"

    #  ======取得模座參數======
    F_function_Extremum_point("X_max", "Y_max", 0, 2, hybridShape2)
    element_point3 = element_Reference1
    part1.Update()
    sketch_build_callout(element_sketch(1), "Horizontal", "Callout", lower_die_wide, 2)
    # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    sketch_build_callout(element_sketch(1), "Vertical", "Callout", lower_die_length, 2)
    # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    part1.Update()
    #  ======取得模座參數======

    element_point2 = hybridShape7
    sketch_build_callout(element_sketch(1), "Horizontal", "Callout", part_center_X, 1)
    # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    sketch_build_callout(element_sketch(1), "Vertical", "Callout", part_center_Y, 1)
    # element_point(now_point)[點之陳述句]和element_point(now_point+1進行標註)direction="Horizontal"or"Vertical"or"free"
    part1.Update()
    hybridShape8 = element_body1.HybridShapes.Item("center_location")  # 宣告平面
    if part_center_Y > lower_die_length / 2:
        hole_location_Y = 10
    elif part_center_Y < lower_die_length / 2:
        hole_location_Y = length(3).Value - 10
    else:
        hole_location_Y = None
    hole_location_X = length(4).Value / 2

    element_Reference10 = hybridShape4
    F_function_build_point()  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out
    element_point5.X.Value = hole_location_X
    element_point5.Y.Value = hole_location_Y
    part1.Update()
    element_point5.Name = "hole_point"
    element_Reference11 = element_point5
    element_Reference12 = hybridShape9
    F_function_hole_simple_D(11, 20, hole, 0)
    # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point element_Reference12 = plane direction = > 0 = 下 1 = 上

    limit1 = hole.BottomLimit
    limit1.LimitMode = catUpToPlaneLimit
    reference3 = part1.CreateReferenceFromObject(hybridShapePlaneExplicit1)
    limit1.LimitingElement = reference3
    part1.Update()
    element_Reference11 = element_point5
    element_Reference12 = hybridShape10
    F_function_hole_simple_D(14, 30, hole, 1)
    # 直孔  (M,深度,out孔陳述式,方向) element_Reference(11)=point element_Reference(12) = plane direction = > 0 = 下 1 = 上

    element_Reference10 = element_point5
    F_function_build_point()  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out

    element_point5.Z.Value = -(length(6).Value - 30)
    element_point5.Name = "Bolt_Start_point"

    element_Reference10 = element_point5
    F_function.build_point()  # 建立點型態(點-點)  element_Reference(10)依據點(全域變數)  element_point(5) 為out

    element_point5.Z.Value = -10
    element_point5.Name = "Bolt_End_point"

    # ========數字二位化+part改名==========
    x = 0
    if NowDataNumber >= 10:
        x = None
    PartName = (str("op" + str(OpNumber) + str(InterferancePadName) + str(x) + str(NowDataNumber)))  # 樹枝圖名稱
    product1.PartNumber = (PartName)
    # ========數字二位化+part改名==========

    # ========設定性質==========
    partDocument1 = catapp.ActiveDocument
    product1 = partDocument1.GetItem(str("op" + str(OpNumber) + str(InterferancePadName) + str(x) + str(NowDataNumber)))
    product1 = product1.ReferenceProduct
    parameters1 = product1.UserRefProperties
    strParam1 = parameters1.CreateString("NO.", "")
    strParam1.ValuateFromString("")
    parameters2 = product1.UserRefProperties
    strParam2 = parameters2.CreateString("Part Name", "")
    strParam2.ValuateFromString("")
    parameters3 = product1.UserRefProperties
    strParam3 = parameters3.CreateString("Size", "")
    strParam3.ValuateFromString("")
    parameters4 = product1.UserRefProperties
    strParam4 = parameters4.CreateString("Material_Data", "")
    strParam4.ValuateFromString(strip_parameter_list[37])
    parameters5 = product1.UserRefProperties
    strParam5 = parameters5.CreateString("Heat Treatment", "")
    strParam5.ValuateFromString(strip_parameter_list[38])
    parameters6 = product1.UserRefProperties
    strParam6 = parameters6.CreateString("Quantity", "")
    strParam6.ValuateFromString("")
    parameters7 = product1.UserRefProperties
    strParam7 = parameters7.CreateString("Page", "")
    strParam7.ValuateFromString("")
    parameters8 = product1.UserRefProperties
    strParam8 = parameters8.CreateString("L1", "")
    strParam8.ValuateFromString("")
    parameters9 = product1.UserRefProperties
    strParam9 = parameters9.CreateString("A", "")
    strParam9.ValuateFromString("")
    parameters10 = product1.UserRefProperties
    strParam10 = parameters10.CreateString("HP", "")
    strParam10.ValuateFromString("")
    parameters11 = product1.UserRefProperties
    strParam11 = parameters11.CreateString("B", "")
    strParam11.ValuateFromString("")
    parameters12 = product1.UserRefProperties
    strParam12 = parameters12.CreateString("BP", "")
    strParam12.ValuateFromString("")
    parameters13 = product1.UserRefProperties
    strParam13 = parameters13.CreateString("TS", "")
    strParam13.ValuateFromString("")
    parameters14 = product1.UserRefProperties
    strParam14 = parameters14.CreateString("IG", "")
    strParam14.ValuateFromString("")
    parameters15 = product1.UserRefProperties
    strParam15 = parameters15.CreateString("F", "")
    strParam15.ValuateFromString("")
    parameters16 = product1.UserRefProperties
    strParam16 = parameters16.CreateString("CS", "")
    strParam16.ValuateFromString("")
    parameters17 = product1.UserRefProperties
    strParam17 = parameters17.CreateString("AP", "")
    strParam17.ValuateFromString("")
    # ========設定性質==========

    Delete_Data(OP)  # 將(now_op_number)以外的工站DATA皆刪除

    part1.Update()
    partDocument1.SaveAs(
        save_path + 'op' + str(OpNumber) + InterferancePadName + str(x) + str(NowDataNumber) + '.CATPart')
    all_part_number = all_part_number + 1
    all_part_name.insert(all_part_number, str(product1.PartNumber))
    partDocument1.Close()


def window_change(DataWindow, CloseWindow):
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    part1 = partdoc.Part
    relations1 = part1.Relations
    formulal_Count = part1.Relations.Count
    for form_number in range(1, formulal_Count + 1):
        formula1 = relations1.Item(form_number)
        print(formula1)
        selection1.Add(formula1)

    parameters2 = part1.Parameters
    parameter_count1 = parameters2.RootParameterSet.DirectParameters.Count
    for parame_number in range(1, parameter_count1 + 1):
        paramet1 = parameters2.RootParameterSet.DirectParameters.Item(parame_number)
        selection1.Add(paramet1)

    parameter_count2 = parameters2.RootParameterSet.ParameterSets.Count
    for parame_number in range(1, parameter_count2 + 1):
        paramet1 = parameters2.RootParameterSet.ParameterSets.Item(parame_number)
        selection1.Add(paramet1)

    bodies1 = part1.Bodies
    bodies_Count = part1.Bodies.Count
    for bodies_number in range(1, bodies_Count + 1):
        bodie1 = bodies1.Item(bodies_number)
        selection1.Add(bodie1)

    AxisSystems1 = part1.AxisSystems
    AxisSystems_Count = part1.AxisSystems.Count
    for Axis_number in range(1, AxisSystems_Count + 1):
        Axis1 = AxisSystems1.Item(Axis_number)
        selection1.Add(Axis1)

    hybridBody_Count = part1.HybridBodies.Count
    for hybridBody_number in range(1, hybridBody_Count + 1):
        hybridBody1 = part1.HybridBodies.Item(hybridBody_number)
        selection1.Add(hybridBody1)
    selection1.Copy()
    window = catapp.Windows
    PasteWindow = window.Item(DataWindow)
    PasteWindow.Activate()
    catapp = win32.Dispatch("CATIA.Application")
    partdoc = catapp.ActiveDocument
    selection1 = partdoc.Selection
    part1 = partdoc.part
    selection1.Add(part1)
    selection1.Paste()
    selection1.Clear()
    CloseWin = window.Item(CloseWindow)
    CloseWin.Activate()
    partdoc = catapp.ActiveDocument
    partdoc.Close()


def part_open(dir):
    # 連結CATIA
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents
    # 將路徑設為目錄的文字宣告
    # folderdir = directory
    # 定義零件檔檔名
    part_dir = input_root + dir
    print(part_dir)
    # partdoc = document.Open("%s%s.%s" % (directory,target,"CATPart"))
    # 開啟該零件檔
    partdoc = document.Open(part_dir)
    # return target+'.CATPart'


def DieRuleSerch(die_rule_file_name, excel_Sheet_name):
    DieRuleRoot = str(die_rule_path + '其他規則\\' + die_rule_file_name)
    workbook = openpyxl.load_workbook(DieRuleRoot)
    sheet = workbook[excel_Sheet_name]
    serch_result = sheet.cell(row=3, column=2).value


PunchMaking()

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
    plate_line_down_quickly_remove_cut_line_number = [[0] * 99 for i in range(99)]
    plate_line_down_quickly_remove_cut_line_number[1][2] = 1
    plate_line_down_quickly_remove_cut_line_number[1][3] = 0
    plate_line_down_quickly_remove_cut_line_number[1][4] = 0
    plate_line_down_quickly_remove_cut_line_number[1][6] = 0
    plate_line_down_quickly_remove_cut_line_number[1][7] = 0

    for now_op_number in range(1, total_op_number + 1):
        n = now_op_number
        op_number = 10 * n
        if plate_line_down_quickly_remove_cut_line_number[g][n] > 0:
            for now_data_number in range(1, plate_line_down_quickly_remove_cut_line_number[g][n] + 1):  #
                part_open('Data1.CATPart')
                interferance_pad_name = "down_quickly_remove_cut_punch"
                interferance_line_name = "_down_quickly_remove_cut_line_"
                open_file_name = "quickly_remove_punch"
                punch(open_file_name)
                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number, op_number,
                             now_data_number, total_op_number)


def punch(open_file_name):
    part_name = open_file_name + '.CATPart'
    part_open(part_name)
    window_change('Data1.CATPart', part_name)


def punch_change(InterferancePadName, InterferanceLineName, NowPlateLineNumber, OpNumber, NowDataNumber,
                 TotalOpNumber):  # 模具名稱,沖頭線段,線段編號,op號碼,第幾條
    all_part_number = 0
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = document.part
    parameter1 = part1.Parameters
    g = now_plate_line_number
    n = now_op_number
    data_type = "line"
    part_type = "cut_punch_down"
    part_name = part_type
    product_name = "down_quickly_remove_cut_punch"
    modify_name = plate_line_ + g + _op + OpNumber + InterferanceLineName
    creat_point_name1 = "down_quickly_remove_cut_line_Ymax_point"
    creat_point_name2 = "down_quickly_remove_cut_line_Ymin_point"
    X_direction1 = 0
    Y_direction1 = 1
    Z_direction1 = 0
    first_direction1 = 1
    first_direction2 = 0
    X_direction2 = 1
    Y_direction2 = 0
    Z_direction2 = 0
    second_direction1 = 1
    second_direction2 = 1

    strParam1 = part1.Parameters.Item("Type")
    strParam1.Value = part_type
    # ========草圖置換==========
    if data_type == "line":
        cut_line_formula_1_name = (str("`die\\" + str(modify_name) + str(NowDataNumber)))
        relaitons = part1.Relations
        formula1 = relaitons.Item("line_formula_1")
        formula1.Modify(cut_line_formula_1_name)
        formula1.Rename(cut_line_formula_1_name)
    else:
        surface_formula_1_name = (str("`die\\" + str(modify_name) + str(NowDataNumber)))
        relaitons = part1.Relations
        formula1 = relaitons.Item("surface_formula_1")
        formula1.Modify(surface_formula_1_name)
        formula1.Rename(surface_formula_1_name)
        line_formula_1_name = (str("`die\\" + str(modify_name) + str(NowDataNumber)))
        relaitons = part1.Relations
        formula1 = relaitons.Item("line_formula_1_1")
        formula1.Modify(line_formula_1_name)
        formula1.Rename(line_formula_1_name)
    part1.Update()
    # ========草圖置換==========

    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    part1.InWorkObject = body1  # 設定工作物件

    # ========建點參數==========
    if data_type == "line":
        item_belong = "die"
    if data_type == "surface":
        item_belong = "Body.2"
        creat_point_number = 2
    Extremum_data_two_condition()  # 建點模組
    # ========建點參數==========

    # ========量測參數==========
    MeasureDistance_number = 1
    Measure_distance_item1 = creat_point_name1
    Measure_distance_item2 = creat_point_name2
    parameter_name1 = modify_name & "max_length_"
    Measure_Distance.CATMain()
    # ========量測參數==========

    length1 = parameter1.Item(parameter_name1 & NowDataNumber)
    length2 = parameter1.Item("QR_width")
    if -int(-length1.value) < 24:
        length2.value = 20
        parttype = part_type + "_narrow"
    if -int(-length1.value) >= 24 and length1.value < 33:
        length2.value = 20
    if -int(-length1.value) >= 33:
        length2.value = -int(-length1.value / 1.5)
    strParam1 = part1.Parameters.Item("Type")
    strParam1.Value = part_type
    length3 = parameter1.Item("bolt_distance")
    if length2.value > 40:
        length3.value = length2.value - (10 + 15)

    # ========設定沖頭高度==========
    length1 = parameter1.Item("punch_up_plane_height")
    if Mode_status == '開模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]) + float(strip_parameter_list[14]) + 28)
    elif Mode_status == '閉模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]) + float(strip_parameter_list[14]))
    length2 = parameter1.Item('QR_height')
    length2.value = (float(strip_parameter_list[14]) + float(strip_parameter_list[17]) * 0.5)

    if data_type != "surface":
        if QR_punch_height != None:
            length5.value = QR_punch_height
        else:
            die_rule_file_name = "沖頭切入深度.xlsx"
            Row_string_serch = die_level  # x
            Column_string_serch = cut_punch_material_kind  # y
            Thickness = float(strip_parameter_list[1])
            excel_Sheet_name = str()
            if Thickness >= 0 and Thickness < 0.1:
                excel_Sheet_name = str("0.1以下")
            elif Thickness >= 0.1 and Thickness < 0.25:
                excel_Sheet_name = str("0.1~0.25")
            elif Thickness >= 0.25 and Thickness < 0.5:
                excel_Sheet_name = str("0.25~0.5")
            elif Thickness >= 0.5 and Thickness < 0.8:
                excel_Sheet_name = str("0.5~0.8")
            elif Thickness >= 0.8 and Thickness < 1.2:
                excel_Sheet_name = str("0.8~1.2")
            elif Thickness >= 1.2 and Thickness < 1.6:
                excel_Sheet_name = str("1.2~1.6")
            elif Thickness >= 1.6 and Thickness < 2.5:
                excel_Sheet_name = str("1.6~2.5")
            elif Thickness >= 2.5 and Thickness < 3.5:
                excel_Sheet_name = str("2.5~3.5")
            elif Thickness >= 3.5:
                excel_Sheet_name = str("3.5以上")
            DieRuleSerch(die_rule_file_name, excel_Sheet_name)
            length2.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
                strip_parameter_list[17]) + float(strip_parameter_list[14]) + serch_result)
    # ========設定沖頭高度==========

    part1.UpdateObject(part1.Bodies.Item("Body.2"))
    product1 = document.GetItem("Part1")
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

    # ========刪除不需要的Data==========
    selection1 = partDocument1.Selection
    NowOpNumber = int(OpNumber / 10)
    if NowPlateLineNumber == 1:
        selection1.Clear()
        selection1.Search('Name=plate_line_2*')
        if selection1.count != 0:
            selection1.Delete()
        selection1.Clear()
    if NowPlateLineNumber == 2:
        selection1.Clear()
        selection1.Search('Name=plate_line_1*')
        if selection1.count != 0:
            selection1.Delete()
        selection1.Clear()
    for o in range(1, NowOpNumber):
        selection1.Clear()
        selection1.Search("Name=*_op" + str(o) + "0_*")
        if selection1.count != 0:
            selection1.Delete()
        selection1.Clear()
    for o in range(NowOpNumber + 1, TotalOpNumber + 1):
        selection1.Clear()
        selection1.Search("Name=*_op" + str(o) + "0_*")
        if selection1.count != 0:
            selection1.Delete()
        selection1.Clear()
    # ========刪除不需要的Data==========

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

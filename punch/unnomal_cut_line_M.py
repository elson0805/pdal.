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
    plate_line_unnomal_cut_line_M_number = [[0] * 99 for i in range(99)]
    plate_line_unnomal_cut_line_M_number[1][2] = 1
    plate_line_unnomal_cut_line_M_number[1][3] = 0
    plate_line_unnomal_cut_line_M_number[1][4] = 0
    plate_line_unnomal_cut_line_M_number[1][6] = 0
    plate_line_unnomal_cut_line_M_number[1][7] = 0

    for now_op_number in range(1, total_op_number + 1):
        n = now_op_number
        op_number = 10 * n
        if plate_line_unnomal_cut_line_M_number[g][n] > 0:
            for now_data_number in range(1, plate_line_unnomal_cut_line_M_number[g][n] + 1):  #
                part_open('Data1.CATPart')
                interferance_pad_name = "_unnomal_cut_punch_M_"
                interferance_line_name = "_unnomal_cut_line_M_"
                open_file_name = "unnomal_cut_punch_M"
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
    # ========設定沖頭==========
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = document.part
    parameter1 = part1.Parameters
    line2_name = str(
        ("die\\plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + "_unnomal_cut_line_M_symmetric_" + str(
            NowDataNumber)))
    line3_name = str(("die\\plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + str(
        InterferanceLineName) + str(NowDataNumber)))
    selection1 = document.Selection
    selection1.clear()
    selection1.Search("Name=" + line2.name)
    if selection1.Count > 0:
        line1.name = line2.name
    else:
        line1.name = line3.name
    cut_line_st = line1.name
    # ========設定沖頭==========
    splint_height = 20
    stop_plate_height = 16
    stripper_plate_height = 16
    length1 = parameter1.Item("M_punch_length")
    length1.Value = (
                float(strip_parameter_list[14]) + float(strip_parameter_list[17]) + float(strip_parameter_list[20]) + 5)
    length2 = parameter1.Item("M_down_thickness")
    length2.Value = -5
    part1.Update()
    # ========改變參數==========
    # ========改變參數==========

    # ========草圖置換==========
    parameters1 = parameter1.Item("1_MMM_1")
    # relaitons = part1.Relations
    formula1 = parameters1.OptionalRelation
    formula1.Modify(cut_line_st)
    formula1.Rename(cut_line_st)
    part1.Update()
    # ========草圖置換==========

    product1 = document.GetItem("Strip_Data-2")
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

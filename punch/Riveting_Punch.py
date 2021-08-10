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
    plate_line_cut_line_number = [[0] * 99 for i in range(99)]
    plate_line_cut_line_number[1][2] = 0
    plate_line_cut_line_number[1][3] = 0
    plate_line_cut_line_number[1][4] = 0
    plate_line_cut_line_number[1][6] = 4
    plate_line_cut_line_number[1][7] = 0

    for now_op_number in range(1, total_op_number + 1):
        n = now_op_number
        op_number = 10 * n
        if plate_line_cut_line_number[g][n] > 0:
            for now_data_number in range(1, plate_line_cut_line_number[g][n] + 1):
                if plate_line_cut_line_number[g][n] == 4:
                    if n == 6:
                        part_open('Data1.CATPart', 0)
                        interferance_pad_name = " "
                        interferance_line_name = "_cut_line_"
                        for i in range(1, 4):
                            type = 1
                            if i == 1:
                                part_name = "op60_cut_punch_01" + '.CATPart'
                                part_open(part_name, 1)
                                open_file_name = "Riveting_Punch"
                                punch(open_file_name, type, i)
                                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number,
                                             op_number, now_data_number, total_op_number, now_op_number)
                            elif i == 2:
                                part_name = "op60_cut_punch_02" + '.CATPart'
                                part_open(part_name, 1)
                                open_file_name = "Riveting_Punch"
                                punch(open_file_name, type, i)
                                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number,
                                             op_number, now_data_number, total_op_number, now_op_number)
                            elif i == 3:
                                part_name = "op60_cut_punch_03" + '.CATPart'
                                part_open(part_name, 1)
                                open_file_name = "Riveting_Punch"
                                punch(open_file_name, type, i)
                                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number,
                                             op_number, now_data_number, total_op_number, now_op_number)
                            elif i == 4:
                                part_name = "op60_cut_punch_04" + '.CATPart'
                                part_open(part_name, 1)
                                open_file_name = "Riveting_Punch"
                                punch(open_file_name, type, i)
                                punch_change(interferance_pad_name, interferance_line_name, now_plate_line_number,
                                             op_number, now_data_number, total_op_number, now_op_number)


def punch(open_file_name, type, i):
    part_name = open_file_name + '.CATPart'
    part_open(part_name, 0)
    if type == 0:
        window_change('Data1.CATPart', part_name)
    elif type == 1:
        if i == 1:
            window_change("op60_cut_punch_01.CATPart", part_name)
            return
        elif i == 2:
            window_change("op60_cut_punch_02.CATPart", part_name)
            return
        elif i == 3:
            window_change("op60_cut_punch_03.CATPart", part_name)
            return
        elif i == 4:
            window_change("op60_cut_punch_04.CATPart", part_name)
            return


def punch_change(InterferancePadName, InterferanceLineName, NowPlateLineNumber, OpNumber, NowDataNumber,
                 TotalOpNumber, now_op_number):  # 模具名稱,沖頭線段,線段編號,op號碼,第幾條
    all_part_number = 0
    # ========設定沖頭==========
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = document.part
    parameter1 = part1.Parameters
    create_reference_60()  # 新增參考點
    relations1 = part1.Relations
    length1 = parameter1.CreateDimension("", "LENGTH", 0)
    length2 = parameter1.CreateDimension("", "LENGTH", 0)
    length1.rename("Data_D")
    length2.rename("Data_height")
    formula1 = relations1.CreateFormula("Data_D", "", length1, "length()")
    formula2 = relations1.CreateFormula("Data_height", "", length2, "length()")
    S_point_distance_parameter = "die\\Extremum.7"
    E_point_distance_parameter = "die\\Extremum.8"
    formula1.Modify('distance(' + S_point_distance_parameter + "," + E_point_distance_parameter + ")")
    part1.Update()
    Data_D = parameter1.Item("Data_D")
    Data_D = Data_D.Value
    Data_height = parameter1.Item("Data_height")
    Data_height = Data_height.Value
    length3 = parameter1.Item("D")
    if Data_D >= 0.3 and Data_D <= 1.56:
        length3.value = 1.6
    elif Data_D >= 0.5 and Data_D <= 1.96:
        length3.value = 2
    elif Data_D >= 0.8 and Data_D <= 2.46:
        length3.value = 2.5
    length4 = parameter1.Item("L")
    length4.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
        strip_parameter_list[17]) + float(strip_parameter_list[14]) + 1)  # 沖頭高度
    length5 = parameter1.Item("H")
    length6 = parameter1.Item("height")
    length6.value = -1.0
    length7 = parameter1.Item("B")
    length8 = parameter1.Item("V")
    length8.value = Data_D + 0.01
    length9 = parameter1.Item("F")
    if length8.value >= 0.31 and length8.value <= 0.49:
        length9.value = 6
    elif length8.value >= 0.5 and length8.value <= 0.79:
        length9.value = 8
    elif length8.value >= 0.8 and length8.value <= 0.99:
        length9.value = 10
    elif length8.value >= 1 and length8.value <= 1.99:
        length9.value = 20
    elif length8.value >= 2:
        length9.value = 35
    # ========設定沖頭==========
    # ========草圖置換==========
    plate_line_A_punch_number = [[0] * 99 for i in range(99)]
    plate_line_A_punch_number[1][2] = 2
    plate_line_A_punch_number[1][3] = 0
    plate_line_A_punch_number[1][4] = 0
    plate_line_A_punch_number[1][6] = 0
    plate_line_A_punch_number[1][7] = 0
    g = 1
    n = now_op_number
    for i in range(1, plate_line_A_punch_number[g][n] + 1):
        cut_line_formula_1_name = (str(
            "`die\\plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + str(InterferanceLineName) + str(
                NowDataNumber) + '`'))
        relaitons = part1.Relations
        formula1 = relaitons.Item("cut_line_formula_1")
        formula1.Modify(cut_line_formula_1_name)
        formula1.Rename(cut_line_formula_1_name)
        part1.Update()

    # ========草圖置換==========
    Translate_punch()
    product1 = document.GetItem("Strip_Data-2")
    partDocument1 = catapp.ActiveDocument
    partDocument1.SaveAs(save_path + "Riveting_Punch_Temporary_" + str(NowDataNumber) + '.CATPart')
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


def part_open(dir, type):
    # 連結CATIA
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents
    if type == 0:
        part_dir = input_root + dir
        print(part_dir)
        partdoc = document.Open(part_dir)
    elif type == 1:
        part_dir = save_path + dir
        print(part_dir)
        partdoc = document.Open(part_dir)


def DieRuleSerch(die_rule_file_name, excel_Sheet_name):
    DieRuleRoot = str(die_rule_path + '其他規則\\' + die_rule_file_name)
    workbook = openpyxl.load_workbook(DieRuleRoot)
    sheet = workbook[excel_Sheet_name]
    serch_result = sheet.cell(row=3, column=2).value


def create_reference_60():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = document.part
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
    parameters1 = part1.Parameters
    hybridShapeCircleExplicit1 = parameters1.Item("cut_line_1")
    reference1 = part1.CreateReferenceFromObject(hybridShapeCircleExplicit1)
    hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)
    hybridBodies1 = part1.HybridBodies
    hybridBody1 = hybridBodies1.Item("die")
    hybridBody1.AppendHybridShape(hybridShapeExtremum1)
    part1.InWorkObject = hybridShapeExtremum1
    part1.Update()
    hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
    reference2 = part1.CreateReferenceFromObject(hybridShapeCircleExplicit1)
    hybridShapeExtremum2 = hybridShapeFactory1.AddNewExtremum(reference2, hybridShapeDirection2, 0)
    hybridBody1.AppendHybridShape(hybridShapeExtremum2)
    part1.InWorkObject = hybridShapeExtremum2
    part1.Update()


def Translate_punch():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    part1 = document.part
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    part1.InWorkObject = body1
    hybridShapeFactory1 = part1.HybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
    shapeFactory1 = part1.ShapeFactory
    pitch = 40  # 暫用
    translate1 = shapeFactory1.AddNewTranslate2(5 * pitch)
    hybridShapeTranslate1 = translate1.HybridShape
    hybridShapeTranslate1.VectorType = 0
    hybridShapeTranslate1.direction = hybridShapeDirection1
    part1.InWorkObject = hybridShapeTranslate1
    part1.Update()


PunchMaking()

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
    plate_line_forming_punch_surface_number = [[0] * 99 for i in range(99)]
    plate_line_forming_punch_surface_number[1][2] = 1
    plate_line_forming_punch_surface_number[1][3] = 0
    plate_line_forming_punch_surface_number[1][4] = 0
    plate_line_forming_punch_surface_number[1][6] = 0
    plate_line_forming_punch_surface_number[1][7] = 0

    for now_op_number in range(1, total_op_number + 1):
        n = now_op_number
        op_number = 10 * n
        if plate_line_forming_punch_surface_number[g][n] > 0:
            for now_data_number in range(1, plate_line_forming_punch_surface_number[g][n] + 1):
                part_open('Data1.CATPart')
                interferance_pad_name = "_forming_punch_"
                interferance_line_name = "forming_punch_surface_project_"
                open_file_name = "forming_cavity"
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
    selection1 = document.Selection
    length1 = parameter1.Item("forming_cavity_center_point_distance_Z")
    if Mode_status == '開模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]) + 28)
        # Thickness + stripper_plate_height + stop_plate_height + upper_die_open_height
    elif Mode_status == '閉模':
        length1.value = (float(strip_parameter_list[1]) + float(strip_parameter_list[20]) + float(
            strip_parameter_list[17]))
    length2 = parameter1.Item('forming_cavity_height')
    length2.value = float(strip_parameter_list[14])
    # ========設定沖頭==========

    # ========將型面往上OFFSET========
    hybridShapeFactory1 = part1.hybridShapeFactory
    hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(0, 0, 1)
    hybridShapeTranslate1 = hybridShapeFactory1.AddNewEmptyTranslate()
    hybridShapeSurfaceExplicit1 = parameter1.Item(str(
        "die\\plate_line_" + str(NowPlateLineNumber) + "_op" + str(OpNumber) + "_forming_punch_surface_" + str(NowDataNumber)))
    reference1 = part1.CreateReferenceFromObject(hybridShapeSurfaceExplicit1)
    hybridShapeTranslate1.ElemToTranslate = reference1
    hybridShapeTranslate1.VectorType = 0
    hybridShapeTranslate1.direction = hybridShapeDirection1
    hybridShapeTranslate1.DistanceValue = (float(strip_parameter_list[20]) + float(
        strip_parameter_list[17]) + float(strip_parameter_list[14]) + 28)
    hybridShapeTranslate1.VolumeResult = False
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    body1.InsertHybridShape(hybridShapeTranslate1)
    part1.InWorkObject = hybridShapeTranslate1
    part1.Update()

    reference2 = part1.CreateReferenceFromObject(hybridShapeTranslate1)
    hybridShapeSurfaceExplicit2 = hybridShapeFactory1.AddNewSurfaceDatum(reference2)
    body1.InsertHybridShape(hybridShapeSurfaceExplicit2)
    part1.InWorkObject = hybridShapeSurfaceExplicit2
    part1.InWorkObject.Name = "forming_punch_surface_offset_" + str(NowDataNumber)
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference2)

    selection1.Clear()
    selection1.Search("Name=*_offset*, All ")
    selection1.VisProperties.SetShow("1")  # 1為隱藏,0為顯示
    selection1.Clear()
    # ========將型面往上OFFSET========

    # ========Boundary取得型面外形線========
    reference3 = part1.CreateReferenceFromObject(hybridShapeSurfaceExplicit2)
    hybridShapeBoundary1 = hybridShapeFactory1.AddNewBoundaryOfSurface(reference3)
    body1.InsertHybridShape(hybridShapeBoundary1)
    part1.InWorkObject = hybridShapeBoundary1
    part1.Update()

    reference4 = part1.CreateReferenceFromObject(hybridShapeBoundary1)
    hybridShapeCurveExplicit1 = hybridShapeFactory1.AddNewCurveDatum(reference4)
    body1.InsertHybridShape(hybridShapeCurveExplicit1)
    part1.InWorkObject = hybridShapeCurveExplicit1
    part1.InWorkObject.Name = "forming_punch_surface_boundary_" + str(NowDataNumber)
    hybridShapeFactory1.DeleteObjectForDatum(reference4)
    part1.Update()

    selection1.Clear()
    selection1.Search("Name=*_Boundary*, All ")
    selection1.VisProperties.SetShow("1")  # 1為隱藏,0為顯示
    selection1.Clear()

    # visPropertySet2 = selection1.VisProperties
    # parameters3 = hybridShapeCurveExplicit1.Parent
    # bSTR4 = hybridShapeCurveExplicit1.Name
    # selection2.Add(hybridShapeCurveExplicit1)
    # visPropertySet2 = visPropertySet2.Parent
    # bSTR5 = visPropertySet2.Name
    # bSTR6 = visPropertySet2.Name
    # visPropertySet2.SetShow("1")
    # selection1.Clear()
    # ========Boundary取得型面外形線========

    # ========投影外型線========
    reference5 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)
    hybridShapes1 = body1.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item("forming_cavity_down_plane")
    reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
    hybridShapeProject1 = hybridShapeFactory1.AddNewProject(reference5, reference6)
    hybridShapeProject1.SolutionType = 0
    hybridShapeProject1.Normal = True
    hybridShapeProject1.SmoothingType = 0
    body1.InsertHybridShape(hybridShapeProject1)
    part1.InWorkObject = hybridShapeProject1
    part1.Update()

    reference7 = part1.CreateReferenceFromObject(hybridShapeProject1)
    hybridShapeCurveExplicit2 = hybridShapeFactory1.AddNewCurveDatum(reference7)
    body1.InsertHybridShape(hybridShapeCurveExplicit2)
    part1.InWorkObject = hybridShapeCurveExplicit2
    part1.InWorkObject.Name = "forming_punch_surface_project_" + NowDataNumber
    part1.Update()
    hybridShapeFactory1.DeleteObjectForDatum(reference7)

    # ========草圖置換==========
    cut_line_formula_1_name = str("Body.2\\" + str(InterferanceLineName) + str(NowDataNumber))
    relaitons = part1.Relations
    formula1 = relaitons.Item("cut_line_formula_1")
    formula1.Modify(cut_line_formula_1_name)
    formula1.Rename(cut_line_formula_1_name)
    part1.Update()
    # ========草圖置換==========

    # =========================
    # length34 = part1.Parameters.Item("x_to_x")
    # s = length34.Value
    # length35 = part1.Parameters.Item("y_to_y")
    # r = length35.Value
    # S = -int(-s)
    # R = -int(-r)
    # length36 = part1.Parameters.Item("int_x")
    # length36.Value = S
    # length37 = part1.Parameters.Item("int_y")
    # length37.Value = T
    # =========================

    # part1.Update()
    # selection4 = part1.Selection
    # visPropertySet3 = selection4.VisProperties
    # parameters4 = hybridShapeCurveExplicit2.Parent
    # bSTR7 = hybridShapeCurveExplicit2.Name
    # selection4.Add(hybridShapeCurveExplicit2)
    # visPropertySet3 = visPropertySet3.Parent
    # bSTR8 = visPropertySet3.Name
    # bSTR9 = visPropertySet3.Name
    # visPropertySet3.SetShow("1")
    # selection4.Clear()
    # ========投影外型線========

    # ========長出實體==========
    shapeFactory1 = part1.ShapeFactory
    reference8 = part1.CreateReferenceFromName("")
    pad1 = shapeFactory1.AddNewPadFromRef(reference8, 20)
    reference9 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
    pad1.SetProfileElement(reference9)
    reference10 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit2)
    pad1.SetProfileElement(reference10)
    limit1 = pad1.FirstLimit
    limit1.LimitMode = catUpToSurfaceLimit
    reference11 = part1.CreateReferenceFromObject(hybridShapeSurfaceExplicit1)
    limit1.LimitingElement = reference11
    part1.Update()
    # ========長出實體==========

    # ========刪除前一個零件pad==========
    if i > 1:
        # part1.InWorkObject.Name = "Pad.2"
        bodynumber = part1.bodynumber.Value
        for bodynumber_number in range(20, 2):
            selection1.Clear()
            selection1.Search("Name=*pad." + bodynumber + ",all")
            if selection1.count > 0:
                selection1.Search("Name=*pad." + bodynumber - str(1) + ",all")
                selection1.Delete()
    # ========刪除前一個零件pad==========
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
    product1 = partDocument1.GetItem(
        str("op" + str(OpNumber) + str(InterferancePadName) + str(x) + str(NowDataNumber)))
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
        save_path + 'op' + str(OpNumber) + str(InterferancePadName) + str(x) + str(NowDataNumber) + '.CATPart')
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

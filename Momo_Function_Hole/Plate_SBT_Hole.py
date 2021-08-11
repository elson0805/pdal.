import csv
import win32com.client as win32
import openpyxl
import time

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
stop_plate_height = 16
stripper_plate_height = 16
splint_height = 20
upper_pad_height = 20
upper_die_set_height = 50
lower_die_set_height = 50
lower_die_cavity_plate_height = 40
lower_pad_height = 16
slower_die_cavity_plate_height = 0
Inner_Guiding_Post_Diameter = 20
Inner_Guiding_Post_Length = 100
Up_Inner_Guiding_Post_Length = 16
Under_Inner_Guiding_Post_Length = 40
Under_Inner_Guiding_Post_Material = "SGFZ"
lower_die_set_plate_length = 309.4
lower_die_set_plate_width = 450
SBT_Diameter = 16
SBT_Length = 60
upper_die_set_plate_length = 450

SBT_data = [[0] * 9 for i in range(9)]
SBT_CB_data = [[0] * 9 for j in range(9)]
SBTQuantity = [0] * 9
SBT_Hole_point_X = [0] * 9
SBT_Hole_point_Y = [0] * 9


def Plate_SBT_Hole():  # 等高螺栓挖孔
    (SBT_data11, SBT_data41, SBT_data61, SBT_CB_data51, SBT_CB_data61, SBT_CB_data71, SBT_CB_data81,
     SBTQuantity1) = upper_die_set()  # 上模座
    up_plate(SBT_data41, SBT_CB_data81, SBTQuantity1)  # 上墊板
    splint(SBT_data11, SBT_CB_data81, SBTQuantity1)  # 上夾板
    stop_plate(SBT_data11, SBT_data61, SBT_CB_data81, SBTQuantity1)  # 止擋板
    Stripper(SBT_CB_data51, SBT_CB_data61, SBT_CB_data71, SBT_CB_data81, SBTQuantity1)  # 脫料板


def upper_die_set():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    partDocument1 = document.Open(save_path + "upper_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("plate_length")  # 159mm
    plate_length = length1.Value
    length2 = parameters1.Item("plate_width")  # 390mm
    plate_width = length2.Value
    length3 = parameters1.Item("plate_height")  # 20mm
    if length3.Value > 0:
        plate_height = length3.Value
    elif length3.Value < 0:
        plate_height = -length3.Value
    part1.Update()

    for i in range(1, plate_line_number + 1):
        n = i - 1
        if n > 0:
            length4 = parameters1.Item("Spacing_" + str(n))
            Spacing = length4.Value
        elif n == 0:
            Spacing = 0

        # ==========================判斷等高螺栓==========================
        if SBT_Diameter == 10:
            SBT_data[1][1] = 10                  # 直徑_D
            SBT_data[2][1] = SBT_Length          # 長度_L
            SBT_data[3][1] = 6                   # 沉頭厚_E
            SBT_data[4][1] = 16                  # 沉頭直徑_A
            SBT_data[5][1] = 8                   # 前端平面寬_S
            SBT_data[6][1] = 8                   # 前端平面高_F
            SBT_CB_data[1][1] = 5                # 螺栓大小_M
            SBT_CB_data[2][1] = 25               # 螺栓長_BL
            SBT_CB_data[3][1] = 5                # 螺栓沉頭厚
            SBT_CB_data[4][1] = 8                # 螺栓沉頭直徑
            SBT_CB_data[5][1] = 9                # 螺栓沉頭孔直徑
            SBT_CB_data[6][1] = 8                # 螺栓沉頭孔深度
            SBT_CB_data[7][1] = 5.5              # 頸部孔直徑
            SBT_CB_data[8][1] = 200              # 等高螺栓與等高螺栓間的間距
        elif SBT_Diameter == 13:
            SBT_data[1][1] = 13
            SBT_data[2][1] = SBT_Length
            SBT_data[3][1] = 7
            SBT_data[4][1] = 18
            SBT_data[5][1] = 11
            SBT_data[6][1] = 10
            SBT_CB_data[1][1] = 6
            SBT_CB_data[2][1] = 25
            SBT_CB_data[3][1] = 6
            SBT_CB_data[4][1] = 10
            SBT_CB_data[5][1] = 11
            SBT_CB_data[6][1] = 9
            SBT_CB_data[7][1] = 7
            SBT_CB_data[8][1] = 200
        elif SBT_Diameter == 16:
            SBT_data[1][1] = 16
            SBT_data[2][1] = SBT_Length
            SBT_data[3][1] = 9
            SBT_data[4][1] = 24
            SBT_data[5][1] = 14
            SBT_data[6][1] = 12
            SBT_CB_data[1][1] = 8
            SBT_CB_data[2][1] = 30
            SBT_CB_data[3][1] = 8
            SBT_CB_data[4][1] = 13
            SBT_CB_data[5][1] = 15
            SBT_CB_data[6][1] = 11
            SBT_CB_data[7][1] = 9
            SBT_CB_data[8][1] = 245
        elif SBT_Diameter == 20:
            SBT_data[1][1] = 10
            SBT_data[2][1] = SBT_Length
            SBT_data[3][1] = 11
            SBT_data[4][1] = 27
            SBT_data[5][1] = 17
            SBT_data[6][1] = 14
            SBT_CB_data[1][1] = 10
            SBT_CB_data[2][1] = 30
            SBT_CB_data[3][1] = 10
            SBT_CB_data[4][1] = 16
            SBT_CB_data[5][1] = 17
            SBT_CB_data[6][1] = 13
            SBT_CB_data[7][1] = 11
            SBT_CB_data[8][1] = 245
        # ==========================判斷等高螺栓==========================

        document = catapp.ActiveDocument
        part1 = document.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("PartBody")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes

        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeProject1 = hybridShapes1.Item("Project.1")
        reference1 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 0
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.Update()
        # ==========================建基準點==========================

        # ==========================隱藏基準點==========================
        selection1 = partDocument1.Selection
        visPropertySet1 = selection1.VisProperties
        hybridShapes1 = hybridShapeExtremum1.Parent
        bSTR1 = hybridShapeExtremum1.Name
        selection1.Add(hybridShapeExtremum1)
        visPropertySet1 = visPropertySet1.Parent
        bSTR2 = visPropertySet1.Name
        bSTR3 = visPropertySet1.Name
        visPropertySet1.SetShow(1)
        selection1.Clear()
        # ==========================隱藏基準點==========================

        b = Spacing + 0
        if i == 1:
            b = 0
        length6 = parameters1.Item("lower_die_set_lower_die_X")
        lower_die_set_lower_die_X = length6.Value
        length6 = parameters1.Item("lower_up_die_set_X")
        lower_up_die_set_X = length6.Value
        length6 = parameters1.Item("lower_die_set_lower_die_Y")
        lower_die_set_lower_die_Y = length6.Value
        length6 = parameters1.Item("lower_up_die_set_X")
        lower_up_die_set_Y = length6.Value

        upper_plate_stripper_bolt_number = 0
        SBTQuantity[1] = (-int(
            -(upper_die_set_plate_length - (75 + lower_die_set_lower_die_X + lower_up_die_set_X + b) * 2) /
            SBT_CB_data[8][1]) + 1) * 2
        # SBTQuantity[1] = 2
        for j in range(1, SBTQuantity[1] + 1):
            C1 = 75 + lower_die_set_lower_die_X + lower_up_die_set_X + b  # 140 + b
            C2 = 46 + lower_die_set_lower_die_Y + lower_up_die_set_Y  # 186
            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (SBTQuantity[1] / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = plate_width - C2
            elif j > (SBTQuantity[1] / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data[8][1] * (j - SBTQuantity[1] / 2 - 1)
            elif j != 1 and j != (SBTQuantity[1] / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data[8][1] * (j - 1)

            SBT_Hole_point_X[j] = X_Coordinate
            SBT_Hole_point_Y[j] = Y_Coordinate

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "SBT_point_" + str(j)
            part1.Update()

            # ==========================建點==========================
            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference3 = hybridShapes1.Item("up_plane")
            sketch1 = sketches1.Add(reference3)
            arrayOfVariantOfDouble1 = [0, 0, 70, 1, 0, 0, 0, 1, 0]
            sketch1Variant = sketch1
            sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
            part1.InWorkObject = sketch1

            factory2D1 = sketch1.OpenEdition()
            geometricElements1 = sketch1.GeometricElements
            axis2D1 = geometricElements1.Item("AbsoluteAxis")
            line2D1 = axis2D1.getItem("HDirection")
            line2D1.ReportName = 1
            line2D2 = axis2D1.getItem("VDirection")
            line2D2.ReportName = 2
            point2D1 = factory2D1.CreatePoint(22.51, -90)
            point2D1.ReportName = 3
            circle2D1 = factory2D1.CreateClosedCircle(22.51, -90, 6)
            circle2D1.CenterPoint = point2D1
            circle2D1.ReportName = 4

            constraints1 = sketch1.Constraints
            reference4 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference4)
            constraint1.mode = 0

            length5 = constraint1.dimension
            length5.Value = SBT_data[4][1] / 2 + 2
            reference5 = hybridShapes1.Item("SBT_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference5)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference6 = part1.CreateReferenceFromObject(point2D1)
            reference7 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference6, reference7)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "SBT_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            sketches2 = body1.Sketches
            M = sketches2.Item("SBT_Hole_Sketch_" + str(j))
            part1.Update()

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            pocket1 = shapeFactory1.AddNewPocket(M, 20)
            limit1 = pocket1.FirstLimit
            limit1.LimitMode = 2

            upper_plate_stripper_bolt_number += 1
            pocket1.Name = "Upper-die-set-Stripper-bolt-spacer-" + str(upper_plate_stripper_bolt_number)
            part1.Update()
            # ==========================挖孔==========================

    strParam1 = parameters1.Item("Properties\\CS")
    strParam1.Value = "IG: " + str(upper_plate_stripper_bolt_number) + "-%%C" + str(
        length5.Value * 2) + "鑽穿(等高套筒)"

    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=*_point_*, All")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()
    selection1.Search("Name=*Sketch.*, All")
    selection1.VisProperties.SetShow(1)

    time.sleep(1)
    partDocument1.save()
    partDocument1.Close()

    SBTQuantity1 = SBTQuantity[1]
    SBT_data11 = SBT_data[1][1]
    SBT_data41 = SBT_data[4][1]
    SBT_data61 = SBT_data[6][1]
    SBT_CB_data51 = SBT_CB_data[5][1]
    SBT_CB_data61 = SBT_CB_data[6][1]
    SBT_CB_data71 = SBT_CB_data[7][1]
    SBT_CB_data81 = SBT_CB_data[8][1]

    return SBT_data11, SBT_data41, SBT_data61, SBT_CB_data51, SBT_CB_data61, SBT_CB_data71, SBT_CB_data81, SBTQuantity1


def up_plate(SBT_data41, SBT_CB_data81, SBTQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "up_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("up_plate_" + str(i) + "\\plate_length")
        plate_length = length1.Value

        length2 = parameters1.Item("up_plate_" + str(i) + "\\plate_width")
        plate_width = length2.Value

        length3 = parameters1.Item("up_plate_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        Quantity = 4

        document = catapp.ActiveDocument
        part1 = document.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes

        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeProject1 = hybridShapes1.Item("Project.1")
        reference1 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 0
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.Update()
        # ==========================建基準點==========================

        # ==========================隱藏基準點==========================
        selection1 = partDocument1.Selection
        visPropertySet1 = selection1.VisProperties
        hybridShapes1 = hybridShapeExtremum1.Parent
        bSTR1 = hybridShapeExtremum1.Name
        selection1.Add(hybridShapeExtremum1)
        visPropertySet1 = visPropertySet1.Parent
        bSTR2 = visPropertySet1.Name
        bSTR3 = visPropertySet1.Name
        visPropertySet1.SetShow(1)
        selection1.Clear()
        # ==========================隱藏基準點==========================

        up_plate_machining_instructions_SBT_Hole = 0
        upper_plate_stripper_bolt_number = 0
        # ==========================建點==========================
        for j in range(1, SBTQuantity1 + 1):
            up_plate_machining_instructions_SBT_Hole += 1
            C1 = 75
            C2 = 46
            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (SBTQuantity[1] / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = plate_length - C2
            elif j > (SBTQuantity[1] / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - SBTQuantity1 / 2 - 1)
            elif j != 1 and j != (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - 1)

            SBT_Hole_point_X[j] = X_Coordinate
            SBT_Hole_point_Y[j] = Y_Coordinate

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "SBT_point_" + str(j)
            part1.Update()

            # ==========================建點==========================
            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference3 = hybridShapes1.Item("down_die_plate_up_plane")
            sketch1 = sketches1.Add(reference3)
            arrayOfVariantOfDouble1 = [0, 0, 70, 1, 0, 0, 0, 1, 0]
            sketch1Variant = sketch1
            sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
            part1.InWorkObject = sketch1

            factory2D1 = sketch1.OpenEdition()
            geometricElements1 = sketch1.GeometricElements
            axis2D1 = geometricElements1.Item("AbsoluteAxis")
            line2D1 = axis2D1.getItem("HDirection")
            line2D1.ReportName = 1
            line2D2 = axis2D1.getItem("VDirection")
            line2D2.ReportName = 2
            point2D1 = factory2D1.CreatePoint(22.51, -90)
            point2D1.ReportName = 3
            circle2D1 = factory2D1.CreateClosedCircle(22.51, -90, 6)
            circle2D1.CenterPoint = point2D1
            circle2D1.ReportName = 4

            constraints1 = sketch1.Constraints
            reference4 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference4)
            constraint1.mode = 0

            length5 = constraint1.dimension
            length5.Value = SBT_data41 / 2 + 2
            reference5 = hybridShapes1.Item("SBT_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference5)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference6 = part1.CreateReferenceFromObject(point2D1)
            reference7 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference6, reference7)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "SBT_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            sketches2 = body1.Sketches
            M = sketches2.Item("SBT_Hole_Sketch_" + str(j))
            part1.Update()

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            pocket1 = shapeFactory1.AddNewPocket(M, 20)
            limit1 = pocket1.FirstLimit
            limit1.LimitMode = 2

            upper_plate_stripper_bolt_number += 1
            pocket1.Name = "Upper-plate-Stripper-bolt-spacer-" + str(upper_plate_stripper_bolt_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\CS")
        strParam1.Value = "IG: " + str(up_plate_machining_instructions_SBT_Hole) + "-%%C" + str(
            length5.Value * 2) + "鑽穿(等高套筒)"

        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All")
        selection1.VisProperties.SetShow(1)

        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


def splint(SBT_data11, SBT_CB_data81, SBTQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Splint_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Splint_" + str(i) + "\\plate_length")
        plate_length = length1.Value

        length2 = parameters1.Item("Splint_" + str(i) + "\\plate_width")
        plate_width = length2.Value

        length3 = parameters1.Item("Splint_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        Quantity = 4

        document = catapp.ActiveDocument
        part1 = document.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes

        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeProject1 = hybridShapes1.Item("Project.1")
        reference1 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 0
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.Update()
        # ==========================建基準點==========================

        # ==========================隱藏基準點==========================
        selection1 = partDocument1.Selection
        visPropertySet1 = selection1.VisProperties
        hybridShapes1 = hybridShapeExtremum1.Parent
        bSTR1 = hybridShapeExtremum1.Name
        selection1.Add(hybridShapeExtremum1)
        visPropertySet1 = visPropertySet1.Parent
        bSTR2 = visPropertySet1.Name
        bSTR3 = visPropertySet1.Name
        visPropertySet1.SetShow(1)
        selection1.Clear()
        # ==========================隱藏基準點==========================

        splint_machining_instructions_SBT_Hole = 0
        splint_stripper_bolt_number = 0
        for j in range(1, SBTQuantity1 + 1):
            splint_machining_instructions_SBT_Hole += 1

            C1 = 75
            C2 = 46
            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = plate_length - C2
            elif j > (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - SBTQuantity1 / 2 - 1)
            elif j != 1 and j != (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - 1)

            SBT_Hole_point_X[j] = X_Coordinate
            SBT_Hole_point_Y[j] = Y_Coordinate

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "SBT_point_" + str(j)
            part1.Update()
            # ==========================建點==========================

            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference3 = hybridShapes1.Item("down_die_plate_up_plane")
            sketch1 = sketches1.Add(reference3)
            arrayOfVariantOfDouble1 = [0, 0, 70, 1, 0, 0, 0, 1, 0]
            sketch1Variant = sketch1
            sketch1Variant.SetAbsoluteAxisData(arrayOfVariantOfDouble1)
            part1.InWorkObject = sketch1

            factory2D1 = sketch1.OpenEdition()
            geometricElements1 = sketch1.GeometricElements
            axis2D1 = geometricElements1.Item("AbsoluteAxis")
            line2D1 = axis2D1.getItem("HDirection")
            line2D1.ReportName = 1
            line2D2 = axis2D1.getItem("VDirection")
            line2D2.ReportName = 2
            point2D1 = factory2D1.CreatePoint(22.51, -90)
            point2D1.ReportName = 3
            circle2D1 = factory2D1.CreateClosedCircle(22.51, -90, 6)
            circle2D1.CenterPoint = point2D1
            circle2D1.ReportName = 4

            constraints1 = sketch1.Constraints
            reference4 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference4)
            constraint1.mode = 0

            length5 = constraint1.dimension
            length5.Value = SBT_data11 / 2 + 1
            reference5 = hybridShapes1.Item("SBT_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference5)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference6 = part1.CreateReferenceFromObject(point2D1)
            reference7 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference6, reference7)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "SBT_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            sketches2 = body1.Sketches
            M = sketches2.Item("SBT_Hole_Sketch_" + str(j))
            part1.Update()

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            pocket1 = shapeFactory1.AddNewPocket(M, 20)
            limit1 = pocket1.FirstLimit
            limit1.LimitMode = 2

            splint_stripper_bolt_number += 1
            pocket1.Name = "Splint-Stripper-bolt-spacer-" + str(splint_stripper_bolt_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(splint_machining_instructions_SBT_Hole) + "-%%C" + str(
            length5.Value) + "鑽穿(等高套筒)"

        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All ")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All ")
        selection1.VisProperties.SetShow(1)

        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


def stop_plate(SBT_data11, SBT_data61, SBT_CB_data81, SBTQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Stop_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_length")
        plate_length = length1.Value

        length2 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_width")
        plate_width = length2.Value

        length3 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        Quantity = 4

        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes

        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeProject1 = hybridShapes1.Item("Project.1")
        reference1 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 0
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.Update()
        # ==========================建基準點==========================

        # ==========================隱藏基準點==========================
        selection1 = partDocument1.Selection
        visPropertySet1 = selection1.VisProperties
        hybridShapes1 = hybridShapeExtremum1.Parent
        bSTR1 = hybridShapeExtremum1.Name
        selection1.Add(hybridShapeExtremum1)
        visPropertySet1 = visPropertySet1.Parent
        bSTR2 = visPropertySet1.Name
        bSTR3 = visPropertySet1.Name
        visPropertySet1.SetShow(1)
        selection1.Clear()
        # ==========================隱藏基準點==========================

        stop_plate_machining_instructions_SBT_Hole = 0
        stop_plate_stripper_bolt_number = 0
        for j in range(1, SBTQuantity1 + 1):
            # ==========================建點==========================
            stop_plate_machining_instructions_SBT_Hole += 1

            document = catapp.Documents
            partDocument1 = document.Open(open_path + "shoulder_screw_line.CATPart")
            Part_name = "Stop_plate_" + str(i) + ".CATPart"
            part_name = "shoulder_screw_line.CATPart"
            window_change(Part_name, part_name)

            C1 = 75
            C2 = 46
            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = plate_length - C2
            elif j > (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - SBTQuantity1 / 2 - 1)
            elif j != 1 and j != (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - 1)

            SBT_Hole_point_X[j] = X_Coordinate
            SBT_Hole_point_Y[j] = Y_Coordinate

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "SBT_point_" + str(j)
            part1.Update()

            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(0, 0, stop_plate_height)
            reference5 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "SBT_dir_point_" + str(j)
            part1.Update()
            # ==========================建點==========================

            # ==========================挖孔==========================
            shapeFactory2 = part1.ShapeFactory
            hybridShapes1 = body1.HybridShapes
            reference8 = part1.CreateReferenceFromObject(hybridShapePointCoord2)
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference9 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            hole1 = shapeFactory2.AddNewHoleFromRefPoint(reference8, reference9, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            hole1.ThreadingMode = 1
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length6 = hole1.Diameter
            length6.Value = SBT_data11 + 2
            length7 = limit1.dimension
            length7.Value = stop_plate_height - SBT_data61 + 2

            stop_plate_stripper_bolt_number += 1
            hole1.Name = "Stop-plate-Stripper-bolt-" + str(stop_plate_stripper_bolt_number)
            part1.Update()
            # ==========================建點==========================

            time.sleep(1)
            # ========草圖置換==========
            Point_formula_1_name = (str("Body.2\\SBT_point_" + str(j)))
            parameters1 = part1.Parameters
            parameter1 = parameters1.Item("Point_formula_1")
            formula1 = parameter1.OptionalRelation
            formula1.Modify(Point_formula_1_name)
            formula1.Rename(Point_formula_1_name)
            part1.Update()
            # ========草圖置換==========

            # ==========================判斷Body數量==========================
            document = catapp.ActiveDocument
            selection1 = document.Selection
            for B_n in range(50, 1, -1):
                selection1.Clear()
                selection1.Search("Name=Body." + str(B_n))
                Body_n = selection1.Count
                selection1.Clear()
                if Body_n > 0:
                    body_number = B_n
                    break
            # Body2(j, stop_plate_stripper_bolt_number, body_number)
            catapp = win32.Dispatch("CATIA.Application")
            document = catapp.ActiveDocument
            part1 = document.Part
            parameters1 = part1.Parameters
            bodies1 = part1.Bodies
            body1 = bodies1.Item("Body.2")
            body2 = bodies1.Item("Body." + str(body_number))
            part1.InWorkObject = body2

            shapeFactory1 = part1.ShapeFactory
            reference1 = part1.CreateReferenceFromName("")
            pad1 = shapeFactory1.AddNewPadFromRef(reference1, 20)
            sketches1 = body2.Sketches
            sketch1 = sketches1.Item("Sketch_1")
            reference2 = part1.CreateReferenceFromObject(sketch1)
            pad1.SetProfileElement(reference2)
            limit1 = pad1.FirstLimit
            limit1.LimitMode = 3

            selection1 = document.Selection
            visPropertySet1 = selection1.VisProperties

            hybridShapes1 = body1.HybridShapes
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            hybridShapes1 = hybridShapePlaneOffset1.Parent
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            limit1.LimitingElement = reference3

            part1.InWorkObject = body1
            remove1 = shapeFactory1.AddNewRemove(body2)
            remove1.Name = "Stop-plate-Stripper-bolt-" + str(stop_plate_stripper_bolt_number)
            part1.Update()

            # selection1 = document.Selection
            # selection1.Clear()
            # selection1.Search("Name=shoulder_screw_D,All")
            # selection1.Delete()
            # selection1.Clear()
            # selection1.Search("Name=shoulder_screw_S,All")
            # selection1.Delete()
            # selection1.Clear()

            if j == 1:
                hybridShapePointExplicit1 = parameters1.Item("Point_formula_1")
                part1.InWorkObject = hybridShapePointExplicit1
                part1.InWorkObject.Name = "Point_formula_1" + str(j)
                part1.Update()
            # ==========================判斷Body數量==========================

        strParam1 = parameters1.Item("Properties\\CS")
        strParam1.Value = "CS: " + str(stop_plate_machining_instructions_SBT_Hole) + "-%%C" + str() + "鑽穿(等高套筒)"

        document = catapp.ActiveDocument
        selection1 = document.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All")
        selection1.VisProperties.SetShow(1)

        time.sleep(1)
        document.save()
        document.Close()


def Stripper(SBT_CB_data51, SBT_CB_data61, SBT_CB_data71, SBT_CB_data81, SBTQuantity1):
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Stripper_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Stripper_" + str(i) + "\\plate_length")
        plate_length = length1.Value

        length2 = parameters1.Item("Stripper_" + str(i) + "\\plate_width")
        plate_width = length2.Value

        length3 = parameters1.Item("Stripper_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        document = catapp.ActiveDocument
        part1 = document.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body1.HybridShapes

        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapeProject1 = hybridShapes1.Item("Project.1")
        reference1 = part1.CreateReferenceFromObject(hybridShapeProject1)
        hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 0)
        hybridShapeDirection2 = hybridShapeFactory1.AddNewDirectionByCoord(0, 1, 0)
        hybridShapeExtremum1.Direction2 = hybridShapeDirection2
        hybridShapeExtremum1.ExtremumType2 = 0
        body1.InsertHybridShape(hybridShapeExtremum1)
        part1.InWorkObject = hybridShapeExtremum1
        part1.Update()
        # ==========================建基準點==========================

        # ==========================隱藏基準點==========================
        selection1 = partDocument1.Selection
        visPropertySet1 = selection1.VisProperties
        hybridShapes1 = hybridShapeExtremum1.Parent
        bSTR1 = hybridShapeExtremum1.Name
        selection1.Add(hybridShapeExtremum1)
        visPropertySet1 = visPropertySet1.Parent
        bSTR2 = visPropertySet1.Name
        bSTR3 = visPropertySet1.Name
        visPropertySet1.SetShow(1)
        selection1.Clear()
        # ==========================隱藏基準點==========================

        Stripper_machining_instructions_SBT_Hole = 0
        stripper_plate_stripper_bolt_number = 0
        for j in range(1, SBTQuantity1 + 1):
            Stripper_machining_instructions_SBT_Hole += 1

            # ==========================建點==========================
            C1 = 75
            C2 = 46
            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = plate_length - C2
            elif j > (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - SBTQuantity1 / 2 - 1)
            elif j != 1 and j != (SBTQuantity1 / 2 + 1):
                X_Coordinate = C1 + SBT_CB_data81 * (j - 1)

            SBT_Hole_point_X[j] = X_Coordinate
            SBT_Hole_point_Y[j] = Y_Coordinate

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "SBT_point_" + str(j)
            part1.Update()

            hybridShapes1 = body1.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("SBT_point_" + str(j))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              SBT_CB_data51 / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.Update()

            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              SBT_CB_data71 / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.Update()
            # ==========================建點==========================

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            reference5 = part1.CreateReferenceFromName("")
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference5, 20)
            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference6)
            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference7)

            limit1 = pocket1.FirstLimit
            length1 = limit1.dimension
            length1.Value = SBT_CB_data61
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pocket2 = shapeFactory1.AddNewPocketFromRef(reference8, 15)
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference9)
            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference10)

            limit2 = pocket2.FirstLimit
            limit2.LimitMode = 3

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference11 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit2.LimitingElement = reference11
            stripper_plate_stripper_bolt_number += 1
            pocket1.Name = "Stripper-plate-Stripper-bolt-" + str(stripper_plate_stripper_bolt_number)
            pocket2.Name = "Stripper-plate-Stripper-bolt-" + str(stripper_plate_stripper_bolt_number)
            part1.Update()
            # ==========================挖孔==========================

            # ==========================等高螺栓二組裝點==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, SBT_CB_data61)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "SBT_dir1_point_" + str(j)
            part1.Update()

            hybridShapePointCoord3 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          stripper_plate_height)
            reference6 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord3.PtRef = reference6
            body1.InsertHybridShape(hybridShapePointCoord3)
            part1.InWorkObject = hybridShapePointCoord3
            part1.InWorkObject.Name = "SBT_dir2_point_" + str(j)
            part1.Update()
            # ==========================等高螺栓二組裝點==========================
    strParam1 = parameters1.Item("Properties\\CS")
    strParam1.Value = "IG: " + str(Stripper_machining_instructions_SBT_Hole) + "-%%C" + str(
        SBT_data[2][1]) + "背面沉頭 %%C" + str(SBT_CB_data51 / 2) + "深(等高套筒)"

    document = catapp.ActiveDocument
    selection1 = document.Selection
    selection1.Clear()
    selection1.Search("Name=*_point_*, All")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()
    selection1.Search("Name=*Sketch.*, All")
    selection1.VisProperties.SetShow(1)

    time.sleep(1)
    document.save()
    document.Close()


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
    parameter_count = parameters2.RootParameterSet.DirectParameters.Count
    for parame_number in range(1, parameter_count + 1):
        paramet1 = parameters2.RootParameterSet.DirectParameters.Item(parame_number)
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
    time.sleep(1)
    selection1.Add(part1)
    time.sleep(1)
    selection1.Paste()
    selection1.Clear()
    CloseWin = window.Item(CloseWindow)
    CloseWin.Activate()
    partdoc = catapp.ActiveDocument
    partdoc.Close()


def Body2(j, stop_plate_stripper_bolt_number, body_number):
    catapp = win32.Dispatch("CATIA.Application")
    document = catapp.ActiveDocument
    part1 = document.Part
    parameters1 = part1.Parameters
    bodies1 = part1.Bodies
    body1 = bodies1.Item("Body.2")
    body2 = bodies1.Item("Body." + str(body_number))
    part1.InWorkObject = body2

    shapeFactory1 = part1.ShapeFactory
    reference1 = part1.CreateReferenceFromName("")
    pad1 = shapeFactory1.AddNewPadFromRef(reference1, 20)
    sketches1 = body2.Sketches
    sketch1 = sketches1.Item("Sketch_1")
    reference2 = part1.CreateReferenceFromObject(sketch1)
    pad1.SetProfileElement(reference2)
    limit1 = pad1.FirstLimit
    limit1.LimitMode = 3

    selection1 = document.Selection
    visPropertySet1 = selection1.VisProperties

    hybridShapes1 = body1.HybridShapes
    hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
    hybridShapes1 = hybridShapePlaneOffset1.Parent
    reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
    limit1.LimitingElement = reference3

    part1.InWorkObject = body1
    remove1 = shapeFactory1.AddNewRemove(body2)
    remove1.Name = "Stop-plate-Stripper-bolt-" + str(stop_plate_stripper_bolt_number)
    part1.Update()

    selection1 = document.Selection
    selection1.Clear()
    selection1.Search("Name=shoulder_screw_D,All")
    selection1.Delete()
    selection1.Clear()
    selection1.Search("Name=shoulder_screw_S,All")
    selection1.Delete()
    selection1.Clear()

    hybridShapePointExplicit1 = parameters1.Item("Point_formula_1")
    part1.InWorkObject = hybridShapePointExplicit1
    part1.InWorkObject.Name = "Point_formula_1" + str(j)
    part1.Update()


Plate_SBT_Hole()

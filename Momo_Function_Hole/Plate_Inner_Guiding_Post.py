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

Inner_Guiding_data = [[0] * 5 for i in range(5)]
Inner_Guiding_Post_Bush_up_data = [[0] * 5 for j in range(5)]
InnerGuidingQuantity = [0] * 9
Inner_Guiding_Post_point_X = [0] * 9
Inner_Guiding_Post_point_Y = [0] * 9


def Plate_Inner_Guiding_Post():  # 內導柱/套挖孔
    (Inner_Guiding_data11, InnerGuidingQuantity1) = splint()  # 上夾板
    stop_plate(Inner_Guiding_data11, InnerGuidingQuantity1)  # 止擋板
    (c, d) = Stripper(InnerGuidingQuantity1)  # 脫料板
    lower_die(InnerGuidingQuantity1, c, d)  # 下模板
    lower_pad(InnerGuidingQuantity1)  # 下墊板
    lower_die_set(InnerGuidingQuantity1)  # 下模座


def splint():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Splint_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Splint_" + str(i) + "\\plate_length")  # 159mm
        plate_length = length1.Value
        length2 = parameters1.Item("Splint_" + str(i) + "\\plate_width")  # 390mm
        plate_width = length2.Value
        length3 = parameters1.Item("Splint_" + str(i) + "\\plate_height")  # 20mm
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        # document = catapp.ActiveDocument
        # part1 = document.Part
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

        splint_machining_instructions_Inner_Guiding = 0
        splint_inner_guide_number = 0
        if 200 <= plate_width < 1000:
            InnerGuidingQuantity[1] = 4
        else:
            InnerGuidingQuantity[1] = 6
        for j in range(1, InnerGuidingQuantity[1] + 1):
            splint_machining_instructions_Inner_Guiding += 1  # 加工說明
            # ==========================建內導柱點==========================
            if InnerGuidingQuantity[1] == 4:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = plate_length - C2
                    n = 5
                elif j == 4:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = plate_length - C2
                    n = 7
            if InnerGuidingQuantity[1] == 6:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = plate_width / 2
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = C2
                    n = 5
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = plate_length - C2
                    n = 7
                elif j == 5:
                    X_Coordinate = plate_width / 2
                    Y_Coordinate = plate_length - C2
                    n = 9
                elif j == 6:
                    X_Coordinate = plate_width - 40
                    Y_Coordinate = plate_length - C2
                    n = 11

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, splint_height)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Inner_Guiding_Post_point_" + str(j)
            part1.Update()
            # ==========================建內導柱點==========================

            # ==========================挖導柱孔==========================
            shapeFactory1 = part1.ShapeFactory
            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference3, reference4, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length4 = hole1.Diameter
            length4.Value = 10
            length4.Value = Inner_Guiding_Post_Diameter  # 20mm
            length5 = limit1.dimension
            length5.Value = Inner_Guiding_Post_Length  # 100mm
            # hole1.Reverse()
            part1.InWorkObject.Name = "Splint-Inner-guide-" + str(j)
            part1.Update()

            Inner_Guiding_data[1][1] = length4.Value  # D 20
            Inner_Guiding_data[2][1] = length5.Value  # L 100
            # ==========================挖導柱孔==========================

            # ==========================建內導套點==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          splint_height - 5.3)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Inner_Guiding_Post_dir_point_" + str(j)
            part1.Update()

            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference6 = hybridShapes1.Item("down_die_plate_down_plane")
            sketch1 = sketches1.Add(reference6)
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
            reference7 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference7)
            constraint1.mode = 0

            length6 = constraint1.dimension
            length6.Value = (Inner_Guiding_Post_Diameter + 5) / 2  # 12.5
            reference8 = hybridShapes1.Item("Inner_Guiding_Post_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference8)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference9 = part1.CreateReferenceFromObject(point2D1)
            reference10 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference9, reference10)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "Inner_Guiding_Post_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            # ==========================挖導套孔==========================
            sketches2 = body1.Sketches
            M = sketches2.Item("Inner_Guiding_Post_Hole_Sketch_" + str(j))
            part1.Update()
            pocket1 = shapeFactory1.AddNewPocket(M, 5.3)
            limit2 = pocket1.FirstLimit
            length7 = limit2.dimension
            length7.Value = 5.3
            splint_inner_guide_number += 1
            part1.Update()
            # ==========================挖導套孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(splint_machining_instructions_Inner_Guiding) + "-%%C" + str(
            Inner_Guiding_data[1][1]) + "割, 單+0.005, 正面沉頭 %%C" + str(length6.Value) + "深 (內導柱)"

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

    Inner_Guiding_data11 = Inner_Guiding_data[1][1]

    InnerGuidingQuantity1 = InnerGuidingQuantity[1]
    return Inner_Guiding_data11, InnerGuidingQuantity1


def stop_plate(Inner_Guiding_data11, InnerGuidingQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Stop_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_length")
        plate_length = length1.Value
        length2 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_width")  # 390
        plate_width = length2.Value
        length3 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        # ==========================建基準點==========================
        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")

        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapes1 = body1.HybridShapes
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

        stop_plate_machining_instructions_Inner_Guidin = 0
        stop_plate_inner_guide_number = 0
        for j in range(1, InnerGuidingQuantity1 + 1):
            stop_plate_machining_instructions_Inner_Guidin += 1  # 加工說明
            # ==========================建內導柱點==========================
            if InnerGuidingQuantity1 == 4:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = plate_length - C2
                    n = 5
                elif j == 4:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = plate_length - C2
                    n = 7
            if InnerGuidingQuantity1 == 6:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = plate_width / 2
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = C2
                    n = 5
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = plate_length - C2
                    n = 7
                elif j == 5:
                    X_Coordinate = plate_width / 2
                    Y_Coordinate = plate_length - C2
                    n = 9
                elif j == 6:
                    X_Coordinate = plate_width - C1
                    Y_Coordinate = plate_length - C2
                    n = 11

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "in_Guide_posts_point_" + str(j)
            part1.Update()
            # ==========================建內導柱點==========================

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference3, reference4, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            limit1.LimitMode = 2
            length4 = hole1.Diameter
            length4.Value = Inner_Guiding_Post_Diameter  # 20mm
            stop_plate_inner_guide_number += 1
            hole1.Name = "Stop-plate-Inner-guide-" + str(stop_plate_inner_guide_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(stop_plate_machining_instructions_Inner_Guidin) + "-%%C" + str(
            length4.Value) + "鑽穿(內導柱)"

        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*, All")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*, All ")
        selection1.VisProperties.SetShow(1)

        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


def Stripper(InnerGuidingQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    partDocument1 = document.Open(
        standard_path + "\\Inner_Guiding_Post\\" + Under_Inner_Guiding_Post_Material + "_" + str(
            Inner_Guiding_Post_Diameter) + ".CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    # ==========================抓取導套數值==========================
    length8 = parameters1.Item("D+2")
    length9 = parameters1.Item("Dh14")
    c = length8.Value
    d = length9.Value
    partDocument1.Close()
    # ==========================抓取導套數值==========================

    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Stripper_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("Stripper_" + str(i) + "\\plate_length")
        length2 = parameters1.Item("Stripper_" + str(i) + "\\plate_width")
        length3 = parameters1.Item("Stripper_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        partDocument1 = catapp.ActiveDocument
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("Body.2")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapes1 = body1.HybridShapes
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

        Stripper_machining_instructions_Inner_Guiding = 0
        stripper_plate_inner_guide_number = 0
        for j in range(1, InnerGuidingQuantity1 + 1):
            Stripper_machining_instructions_Inner_Guiding += 1  # 加工說明

            # ==========================建內導柱點==========================
            if InnerGuidingQuantity1 == 4:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 5
                elif j == 4:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
            if InnerGuidingQuantity1 == 6:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 5
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
                elif j == 5:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = length1.Value - C2
                    n = 9
                elif j == 6:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 11

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          stripper_plate_height)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Inner_Guiding_Post_Bush_up_point_" + str(j)
            part1.Update()
            # ==========================建內導柱點==========================

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference3, reference4, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            hole1.ThreadingMode = 1
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length4 = hole1.Diameter
            # length4.Value = 10
            length4.Value = d
            length5 = limit1.dimension
            length5.Value = Up_Inner_Guiding_Post_Length
            part1.InWorkObject.Name = "Stripper-plate-Inner-guide-bush-up-hole-" + str(j)
            part1.Update()

            Inner_Guiding_Post_Bush_up_data[1][1] = length4.Value  # D
            Inner_Guiding_Post_Bush_up_data[2][1] = length5.Value  # L
            # ==========================挖孔==========================

            # ==========================建內導套點==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          stripper_plate_height - 3.3)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Inner_Guiding_Post_dir_point_" + str(j)
            part1.Update()

            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference6 = hybridShapes1.Item("down_die_plate_down_plane")
            sketch1 = sketches1.Add(reference6)
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
            reference7 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference7)
            constraint1.mode = 0

            length6 = constraint1.dimension
            length6.Value = c
            reference8 = hybridShapes1.Item("Inner_Guiding_Post_Bush_up_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference8)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference9 = part1.CreateReferenceFromObject(point2D1)
            reference10 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference9, reference10)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "Inner_Guiding_Post_Bush_up_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            # ==========================挖孔==========================
            sketches2 = body1.Sketches
            M = sketches2.Item("Inner_Guiding_Post_Bush_up_Hole_Sketch_" + str(j))
            part1.Update()
            pocket1 = shapeFactory1.AddNewPocket(M, 3.3)
            limit2 = pocket1.FirstLimit
            length7 = limit2.dimension
            length7.Value = 3.3
            part1.Update()
            stripper_plate_inner_guide_number += 1
            hole1.Name = "Stop-plate-Inner-guide-" + str(stripper_plate_inner_guide_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(Stripper_machining_instructions_Inner_Guiding) + "-%%C" + str(
            length4.Value) + "割, 單+0.005,　正面沉頭 %%C" + str(length6.Value * 2) + "深(內導柱)"

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

    return c, d


def lower_die(InnerGuidingQuantity1, c, d):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "lower_die_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("lower_die_" + str(i) + "\\plate_length")
        length2 = parameters1.Item("lower_die_" + str(i) + "\\plate_width")
        length3 = parameters1.Item("lower_die_" + str(i) + "\\plate_height")
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
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapes1 = body1.HybridShapes
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

        lower_die_machining_instructions_Inner_Guiding = 0
        lower_die_inner_guide_number = 0
        for j in range(1, InnerGuidingQuantity1 + 1):
            lower_die_machining_instructions_Inner_Guiding += 1
            # ==========================建內導柱點==========================
            if InnerGuidingQuantity[1] == 4:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 5
                elif j == 4:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
            if InnerGuidingQuantity[1] == 6:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 5
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
                elif j == 5:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = length1.Value - C2
                    n = 9
                elif j == 6:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 11

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          -lower_die_cavity_plate_height)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Inner_Guiding_Post_Bush_down_point_" + str(j)
            part1.Update()
            # ==========================建內導柱點==========================

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference3, reference4, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            hole1.ThreadingMode = 1
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length4 = hole1.Diameter
            # length4.Value = 10
            length4.Value = d
            length5 = limit1.dimension
            length5.Value = Under_Inner_Guiding_Post_Length
            hole1.Reverse()
            part1.InWorkObject.Name = "Lower-die-Inner-guide-bush-down-hole-" + str(j)
            part1.Update()

            Inner_Guiding_data[1][1] = length4.Value  # D
            Inner_Guiding_data[2][1] = length5.Value  # L
            # ==========================挖孔==========================

            # ==========================建內導套點==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          -lower_die_cavity_plate_height + 3.3)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Inner_Guiding_Post_Bush_down_dir_point_" + str(j)
            part1.Update()

            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            reference6 = hybridShapes1.Item("down_die_plate_down_plane")
            sketch1 = sketches1.Add(reference6)
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
            reference7 = part1.CreateReferenceFromObject(circle2D1)
            constraint1 = constraints1.AddMonoEltCst(14, reference7)
            constraint1.mode = 0

            length6 = constraint1.dimension
            length6.Value = c
            reference8 = hybridShapes1.Item("Inner_Guiding_Post_Bush_down_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference8)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference9 = part1.CreateReferenceFromObject(point2D1)
            reference10 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference9, reference10)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "Inner_Guiding_Post_Bush_down_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            sketches2 = body1.Sketches
            M = sketches2.Item("Inner_Guiding_Post_Bush_down_Hole_Sketch_" + str(j))
            part1.Update()
            pocket1 = shapeFactory1.AddNewPocket(M, 3.3)
            limit2 = pocket1.FirstLimit
            length7 = limit2.dimension
            length7.Value = -3.3
            lower_die_inner_guide_number += 1
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(lower_die_machining_instructions_Inner_Guiding) + "-%%C" + str(
            Inner_Guiding_data[1][1]) + "割, 單+0.005, 背面沉頭 %%C" + str(length6.Value * 2) + "深 (內導柱)"

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


def lower_pad(InnerGuidingQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "lower_pad_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("lower_pad_" + str(i) + "\\plate_length")
        length2 = parameters1.Item("lower_pad_" + str(i) + "\\plate_width")
        length3 = parameters1.Item("lower_pad_" + str(i) + "\\plate_height")
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
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapes1 = body1.HybridShapes
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

        lower_pad_machining_instructions_Inner_Guiding = 0
        lower_pad_inner_guide_number = 0
        for j in range(1, InnerGuidingQuantity1 + 1):
            lower_pad_machining_instructions_Inner_Guiding += 1
            # ==========================建內導柱點==========================
            if InnerGuidingQuantity[1] == 4:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 5
                elif j == 4:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
            if InnerGuidingQuantity[1] == 6:
                C1 = 25
                C2 = 25
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                    n = 1
                elif j == 2:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = C2
                    n = 3
                elif j == 3:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = C2
                    n = 5
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = length1.Value - C2
                    n = 7
                elif j == 5:
                    X_Coordinate = length2.Value / 2
                    Y_Coordinate = length1.Value - C2
                    n = 9
                elif j == 6:
                    X_Coordinate = length2.Value - C1
                    Y_Coordinate = length1.Value - C2
                    n = 11

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "in_Guide_posts_point_" + str(j)
            part1.Update()

            # ==========================建點==========================
            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            hybridShapes1 = body1.HybridShapes
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

            length4 = constraint1.dimension
            length4.Value = 11
            reference5 = hybridShapes1.Item("in_Guide_posts_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference5)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True

            reference6 = part1.CreateReferenceFromObject(point2D1)
            reference7 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference6, reference7)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "in_Guide_posts_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            sketches2 = body1.Sketches
            M = sketches2.Item("in_Guide_posts_Hole_Sketch_" + str(j))
            part1.Update()

            # ==========================挖孔==========================
            shapeFactory1 = part1.ShapeFactory
            pocket1 = shapeFactory1.AddNewPocket(M, 20)
            limit1 = pocket1.FirstLimit
            limit1.LimitMode = 2

            lower_pad_inner_guide_number += 1
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(lower_pad_machining_instructions_Inner_Guiding) + "-%%C" + str(
            length4.Value * 2) + "鑽穿(內導柱)"

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


def lower_die_set(InnerGuidingQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    partDocument1 = document.Open(save_path + "lower_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters
    length1 = parameters1.Item("plate_length")  # 309.4
    plate_length = length1.Value
    length2 = parameters1.Item("plate_width")  # 450
    plate_width = length2.Value
    length3 = parameters1.Item("plate_height")  # -50
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

        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Item("PartBody")
        part1.InWorkObject = body1

        # ==========================建基準點==========================
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1, 0, 0)
        hybridShapes1 = body1.HybridShapes
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

        lower_die_set_machining_instructions_Inner_Guiding = 0
        lower_die_set_inner_guide_number = 0
        for j in range(1, InnerGuidingQuantity1 + 1):
            lower_die_set_machining_instructions_Inner_Guiding += 1

            b = Spacing
            if i == 1:
                b = 0

            # ==========================建內導柱點==========================
            length6 = parameters1.Item("lower_die_set_lower_die_X")
            lower_die_set_lower_die_X = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_X = length6.Value
            length6 = parameters1.Item("lower_die_set_lower_die_Y")
            lower_die_set_lower_die_Y = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_Y = length6.Value
            C1 = 25 + lower_die_set_lower_die_X + lower_up_die_set_X + b
            C2 = 25 + lower_die_set_lower_die_Y

            if InnerGuidingQuantity1 == 4:
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                elif j == 2:
                    X_Coordinate = C1 + 390 - 50
                    Y_Coordinate = C2
                elif j == 3:
                    X_Coordinate = C1
                    Y_Coordinate = plate_length - C2
                elif j == 4:
                    X_Coordinate = C1 + 390 - 50
                    Y_Coordinate = plate_length - C2
            if InnerGuidingQuantity1 == 6:
                if j == 1:
                    X_Coordinate = C1
                    Y_Coordinate = C2
                elif j == 2:
                    X_Coordinate = 390 / 2
                    Y_Coordinate = C2
                elif j == 3:
                    X_Coordinate = C1 + 390 - 50
                    Y_Coordinate = C2
                elif j == 4:
                    X_Coordinate = C1
                    Y_Coordinate = plate_width - C2
                elif j == 5:
                    X_Coordinate = 390 / 2
                    Y_Coordinate = plate_width - C2
                elif j == 6:
                    X_Coordinate = C1 + 390 - 65
                    Y_Coordinate = plate_width - C2

            Inner_Guiding_Post_point_X[j] = X_Coordinate  # 孔位置X座標
            Inner_Guiding_Post_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "in_Guide_posts_point_" + str(j)
            part1.Update()
            # ==========================建內導柱點==========================

            # ==========================繪製草圖==========================
            sketches1 = body1.Sketches
            hybridShapes1 = body1.HybridShapes
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
            length5.Value = 12.5

            reference5 = hybridShapes1.Item("in_Guide_posts_point_" + str(j))
            geometricElements2 = factory2D1.CreateProjections(reference5)
            geometry2D1 = geometricElements2.Item("Mark.1")
            geometry2D1.Construction = True
            reference6 = part1.CreateReferenceFromObject(point2D1)
            reference7 = part1.CreateReferenceFromObject(geometry2D1)
            constraint2 = constraints1.AddBiEltCst(2, reference6, reference7)
            constraint2.mode = 0
            sketch1.CloseEdition()

            part1.InWorkObject = sketch1
            part1.InWorkObject.Name = "in_Guide_posts_Hole_Sketch_" + str(j)
            # ==========================繪製草圖==========================

            # ==========================挖孔==========================
            sketches2 = body1.Sketches
            M = sketches2.Item("in_Guide_posts_Hole_Sketch_" + str(j))
            part1.Update()
            shapeFactory1 = part1.ShapeFactory
            pocket1 = shapeFactory1.AddNewPocket(M, 20)
            limit1 = pocket1.FirstLimit
            limit1.LimitMode = 2
            lower_die_set_inner_guide_number += 1
            pocket1.Name = "Lower-die-set-Inner-guide-" + str(lower_die_set_inner_guide_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\IG")
        strParam1.Value = "IG: " + str(lower_die_set_machining_instructions_Inner_Guiding) + "-%%C" + str(
            length5.Value * 2) + "鑽穿(內導柱)"

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


Plate_Inner_Guiding_Post()

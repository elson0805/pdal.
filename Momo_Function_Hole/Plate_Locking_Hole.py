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
upper_pad_height = 20
upper_die_set_height = 50
lower_die_set_height = 50
lower_die_cavity_plate_height = 40
lower_pad_height = 16
slower_die_cavity_plate_height = 0

CB_data = [[0] * 13 for i in range(13)]
Bolt_data = [[0] * 6 for j in range(6)]
BoltQuantity = [0] * 9
bolt_point_X = [0] * 9
bolt_point_Y = [0] * 9


def Plate_Locking_Hole():  # 模板螺栓挖孔
    (MM1, Bolt_data11, Bolt_data21, Bolt_data31, Bolt_data41, Bolt_data51, BoltQuantity1, CB_data_MM11, CB_data_MM12,
     CB_data_MM13) = upper_die_set()  # 上模座 1
    up_plate(MM1, Bolt_data51, BoltQuantity1)  # 上墊板 1
    splint(MM1, Bolt_data11, Bolt_data21, Bolt_data41, Bolt_data51, BoltQuantity1, CB_data_MM12)  # 上夾板 1
    (Bolt_data12, Bolt_data22, Bolt_data42, Bolt_data52, BoltQuantity2) = stop_plate(MM1)  # 止擋板 2
    Stripper(MM1, Bolt_data12, Bolt_data22, Bolt_data42, Bolt_data52, BoltQuantity2, CB_data_MM12)  # 脫料板 2
    (MM2, Bolt_data13, Bolt_data23, Bolt_data33, Bolt_data43, Bolt_data53, BoltQuantity3, CB_data_MM21,
     CB_data_MM22) = lower_die()  # 下模板 3
    lower_pad(Bolt_data13, Bolt_data53, BoltQuantity3)  # 下墊板 3
    lower_die_set(MM1, Bolt_data53, BoltQuantity3)  # 下模座 3


def upper_die_set():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    partDocument1 = document.Open(save_path + "upper_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters

    length1 = parameters1.Item("plate_length")
    upper_die_set_plate_length = length1.Value

    length2 = parameters1.Item("plate_width")
    upper_die_set_plate_width = length2.Value

    length3 = parameters1.Item("plate_height")  # 模板厚度  判斷螺栓數量
    if length3.Value > 0:
        plate_height = length3.Value
    elif length3.Value < 0:
        plate_height = -length3.Value
    part1.Update()

    q = 0
    R = 0

    for g in range(1, plate_line_number + 1):
        n = g - 1
        if n > 0:
            length4 = parameters1.Item("Spacing_" + str(n))
            Spacing = length4.Value
        elif n == 0:
            Spacing = 0

        # basic_point("PartBody")  # 建基準點

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

        # ==========================依模板厚度決定螺栓直徑大小==========================
        if length3.Value < 15:
            MM1 = 5  # 螺栓直徑
        elif 15 < length3.Value < 22:
            MM1 = 6
        elif 22 < length3.Value < 27:
            MM1 = 8
        elif 27 < length3.Value < 35:
            MM1 = 10
        elif length3.Value > 35:
            MM1 = 12
        # ==========================依模板厚度決定螺栓直徑大小==========================

        # ==========================判斷螺栓沉頭直徑&厚度==========================
        if MM1 == 5:
            CB_data[5][1] = 8  # 沉頭直徑
            CB_data[5][2] = 5  # 沉頭厚度
            CB_data[5][3] = 5.5  # 頸部直徑
        elif MM1 == 6:
            CB_data[6][1] = 10
            CB_data[6][2] = 6
            CB_data[6][3] = 7
        elif MM1 == 8:
            CB_data[8][1] = 13
            CB_data[8][2] = 8
            CB_data[8][3] = 9
        elif MM1 == 10:
            CB_data[10][1] = 16
            CB_data[10][2] = 10
            CB_data[10][3] = 11
        elif MM1 == 12:
            CB_data[12][1] = 18
            CB_data[12][2] = 12
            CB_data[12][3] = 13
        # ==========================判斷螺栓沉頭直徑&厚度==========================

        # ==========================判斷螺栓挖孔尺寸==========================
        if MM1 == 5:
            Bolt_data[1][1] = 6.5  # 頸部直徑
            Bolt_data[2][1] = 8  # 沉頭深
            Bolt_data[3][1] = \
                float(strip_parameter_list[14]) + upper_pad_height + upper_die_set_height  # 沉頭深+螺紋總長
            Bolt_data[4][1] = 9  # 沉頭孔直徑
            Bolt_data[5][1] = 80  # 螺栓與螺栓間的間距
        elif MM1 == 6:
            Bolt_data[1][1] = 8
            Bolt_data[2][1] = 9
            Bolt_data[3][1] = float(strip_parameter_list[14]) + upper_pad_height + upper_die_set_height
            Bolt_data[4][1] = 11
            Bolt_data[5][1] = 80
        elif MM1 == 8:
            Bolt_data[1][1] = 10
            Bolt_data[2][1] = 11
            Bolt_data[3][1] = float(strip_parameter_list[14]) + upper_pad_height + upper_die_set_height
            Bolt_data[4][1] = 15
            Bolt_data[5][1] = 100
        elif MM1 == 10:
            Bolt_data[1][1] = 12
            Bolt_data[2][1] = 13
            Bolt_data[3][1] = float(strip_parameter_list[14]) + upper_pad_height + upper_die_set_height
            Bolt_data[4][1] = 17
            Bolt_data[5][1] = 125
        elif MM1 == 12:
            Bolt_data[1][1] = 13
            Bolt_data[2][1] = 15
            Bolt_data[3][1] = float(strip_parameter_list[14]) + upper_pad_height + upper_die_set_height
            Bolt_data[4][1] = 19
            Bolt_data[5][1] = 150
        # ==========================判斷螺栓挖孔尺寸==========================

        # ==========================建螺栓點==========================
        BoltQuantity[1] = (-int(-(int(length1.Value) - 55 * 2) / Bolt_data[5][1]) + 1) * 2  # 螺栓孔數量(55為防呆數值)
        upper_die_set_machining_instructions_bolt_hole = 0  # 加工說明
        for j in range(1, BoltQuantity[1] + 1):
            upper_die_set_machining_instructions_bolt_hole += 1  # 加工說明
            q += 1
            b = Spacing + length1.Value  # 模板與模板X方向間距
            if g == 1:
                b = 0

            length6 = parameters1.Item("lower_die_set_lower_die_X")
            lower_die_set_lower_die_X = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_X = length6.Value
            length6 = parameters1.Item("lower_die_set_lower_die_Y")
            lower_die_set_lower_die_Y = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_Y = length6.Value

            C1 = int(55 + lower_die_set_lower_die_X + lower_up_die_set_X + b) + 1
            # 基礎值(固定)+下模座與模板X方向間距+下模座與上模座X方向間距+模板與模板X方向間距(複數模板)
            C2 = int(20 + lower_die_set_lower_die_Y + lower_up_die_set_Y) + 1
            # 基礎值(固定)+下模座與模板Y方向間距+下模座與上模座Y方向間距

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (BoltQuantity[1] / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = upper_die_set_plate_width - C2
            elif j > (BoltQuantity[1] / 2 + 1):
                X_Coordinate = C1 + int(Bolt_data[5][1]) * (j - BoltQuantity[1] / 2 - 1)
            elif j != 1 and j != (BoltQuantity[1] / 2 + 1):
                X_Coordinate = C1 + int(Bolt_data[5][1]) * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[g] / 2 - 55
            #     Y_Coordinate = C2
            # elif j == 3:
            #     X_Coordinate = C1 + 2 * (plate_length[g] / 2 - 55)
            #     Y_Coordinate = C2
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = upper_die_set_plate_width - C2
            # elif j == 5:
            #     X_Coordinate = C1 + plate_length[g] / 2 - 55
            #     Y_Coordinate = upper_die_set_plate_width - C2
            # elif j == 6:
            #     X_Coordinate = C1 + 2 * (plate_length[g] / 2 - 55) + 40
            #     Y_Coordinate = upper_die_set_plate_width - C2
            # elif j == 7:
            #     X_Coordinate = C1 + 2 * (plate_length[g] / 2 - 55) + 40
            #     Y_Coordinate = upper_die_set_plate_width - C2

            bolt_point_X[j] = X_Coordinate  # 孔位置X座標  85
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標  95

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(q)
            part1.Update()
        # ==========================建螺栓點==========================

        # ==========================挖孔==========================
        upper_die_set_bolt_number = 0
        for k in range(1, BoltQuantity[1] + 1):
            R += 1
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            # part1.Update()
            shapeFactory1 = part1.ShapeFactory
            body2 = bodies1.Item("PartBody")

            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(R))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, MM1 * 2)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.ThreadingMode = 0
            hole1.CreateStandardThreadDesignTable(1)

            strParam1 = hole1.HoleThreadDescription
            hole1.BottomType = 2
            strParam1.Value = str("M" + str(MM1))  # 螺栓大小
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            upper_die_set_bolt_number += 1
            remove1.Name = "Upper-die-set-Bolt-" + str(upper_die_set_bolt_number)
            part1.Update()
        # ==========================挖孔==========================
    strParam1 = parameters1.Item("Properties\\A")
    strParam1.Value = "A: " + str(upper_die_set_machining_instructions_bolt_hole) + str(16) + "-背面攻M" + str(
        MM1) + "深(下模螺絲)"

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

    Bolt_data11 = Bolt_data[1][1]
    Bolt_data21 = Bolt_data[2][1]
    Bolt_data31 = Bolt_data[3][1]
    Bolt_data41 = Bolt_data[4][1]
    Bolt_data51 = Bolt_data[5][1]
    BoltQuantity1 = BoltQuantity[1]
    CB_data_MM11 = CB_data[MM1][1]
    CB_data_MM12 = CB_data[MM1][2]
    CB_data_MM13 = CB_data[MM1][3]

    return MM1, Bolt_data11, Bolt_data21, Bolt_data31, Bolt_data41, Bolt_data51, BoltQuantity1, CB_data_MM11, CB_data_MM12, CB_data_MM13


def up_plate(MM1, Bolt_data51, BoltQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "up_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("up_plate_" + str(i) + "\\plate_length")
        upper_die_set_plate_length = length1.Value

        length2 = parameters1.Item("up_plate_" + str(i) + "\\plate_width")
        upper_die_set_plate_width = length2.Value

        length3 = parameters1.Item("up_plate_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

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

        # ==========================建螺栓點==========================
        up_plate_machining_instructions_bolt_hole = 0

        for j in range(1, BoltQuantity1 + 1):
            up_plate_machining_instructions_bolt_hole += 1  # 加工說明

            C1 = 55
            C2 = 20

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - BoltQuantity1 / 2 - 1)
            elif j != 1 and j != (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[i] / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length[i] - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width[i] - C2
            #     n =7
            # elif j == 5:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length[i] - 15
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 11
            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()
        # ==========================建螺栓點==========================

        # ==========================挖孔==========================
        shapeFactory1 = part1.ShapeFactory
        body2 = bodies1.Item("Body.2")
        up_plate_bolt_number = 0
        for k in range(1, BoltQuantity1 + 1):
            R += 1
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            # part1.Update()

            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(R))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.Reverse()
            hole1.BottomType = 2
            limit1.LimitMode = 3
            length1 = hole1.Diameter
            length1.Value = Bolt_data[1][1]  # 通孔直徑
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference3
            part1.Update()
            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            up_plate_bolt_number += 1
            remove1.Name = "Upper-plate-Bolt-spacer-" + str(up_plate_bolt_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "A: " + str(up_plate_machining_instructions_bolt_hole) + "-M" + str(MM1) + "鑽穿(上模螺絲)"

        selection1 = partDocument1.Selection
        selection1.Clear()
        selection1.Search("Name=*_point_*")
        selection1.VisProperties.SetShow(1)
        selection1.Clear()
        selection1.Search("Name=*Sketch.*,")
        selection1.VisProperties.SetShow(1)

        time.sleep(1)
        partDocument1.save()
        partDocument1.Close()


def splint(MM1, Bolt_data11, Bolt_data21, Bolt_data41, Bolt_data51, BoltQuantity1, CB_data_MM12):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Splint_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Splint_" + str(i) + "\\plate_length")
        upper_die_set_plate_length = length1.Value
        length2 = parameters1.Item("Splint_" + str(i) + "\\plate_width")
        upper_die_set_plate_width = length2.Value
        length3 = parameters1.Item("Splint_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        if 200 <= length1.Value < 1000:
            Quantity = 6

        document = catapp.ActiveDocument
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

        # ==========================建螺栓點==========================
        splint_machining_instructions_bolt_hole = 0
        for j in range(1, BoltQuantity1 + 1):
            splint_machining_instructions_bolt_hole += 1  # 加工說明

            C1 = 55
            C2 = 20

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - BoltQuantity1 / 2 - 1)
            elif j != 1 and j != (BoltQuantity1 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length(i) / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length(i) - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width(i) - C2
            #     n =7
            # elif j == 5:
            #     X_Coordinate = plate_length(i) / 2
            #     Y_Coordinate = plate_width(i) - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length(i) - 15
            #     Y_Coordinate = plate_width(i) - C2
            #     n = 11
            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          Bolt_data21 - CB_data_MM12)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()

            # ==========================螺栓第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          Bolt_data21)  # 方向點距離(沉頭深度)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Locking_dir_point_" + str(j)
            part1.Update()
            # ==========================螺栓第二點(決定方向點)==========================
        # ==========================建螺栓點==========================

        # ==========================挖孔==========================
        body2 = bodies1.Item("Body.2")
        shapeFactory1 = part1.ShapeFactory
        hybridShapeFactory1 = part1.HybridShapeFactory
        hybridShapes1 = body2.HybridShapes
        splint_bolt_number = 0
        for k in range(1, BoltQuantity1 + 1):
            body1 = bodies1.Add()
            body1.Name = "body_remove_Body"
            part1.Update()

            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(k))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              Bolt_data11 / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference3, reference4, True,
                                                                              Bolt_data41 / 2)
            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference5, float(strip_parameter_list[14]))
            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference6)
            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference7)
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pocket2 = shapeFactory1.AddNewPocketFromRef(reference8, Bolt_data21)
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference9)
            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference10)
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            splint_bolt_number += 1
            remove1.Name = "Splint-Bolt-" + str(splint_bolt_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "A: " + str(splint_machining_instructions_bolt_hole) + "-M" + str(
            Bolt_data11) + "鑽穿, 背面沉頭 %%C" + str(Bolt_data41) + "深" + str(Bolt_data21) + "mm(上模螺絲)"

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


def stop_plate(MM1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        R = 0
        partDocument1 = document.Open(save_path + "Stop_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_length")
        upper_die_set_plate_length = length1.Value
        length2 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_width")
        upper_die_set_plate_width = length2.Value
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

        # ==========================建螺栓點==========================

        if float(strip_parameter_list[17]) > float(strip_parameter_list[20]):
            #  stop_plate_height > stripper_plate_height

            # ==========================依模板厚度決定螺栓直徑大小==========================
            if length3.Value < 15:
                MM1 = 5  # 螺栓直徑
            elif 15 < length3.Value < 22:
                MM1 = 6
            elif 22 < length3.Value < 27:
                MM1 = 8
            elif 27 < length3.Value < 35:
                MM1 = 10
            elif length3.Value > 35:
                MM1 = 12
            # ==========================依模板厚度決定螺栓直徑大小==========================

            # ==========================判斷螺栓沉頭直徑&厚度==========================
            if MM1 == 5:
                CB_data[MM1][1] = 8  # 沉頭直徑
                CB_data[MM1][2] = 5  # 沉頭厚度
                CB_data[MM1][3] = 5.5  # 頸部直徑
            elif MM1 == 6:
                CB_data[MM1][1] = 10
                CB_data[MM1][2] = 6
                CB_data[MM1][3] = 7
            elif MM1 == 8:
                CB_data[MM1][1] = 13
                CB_data[MM1][2] = 8
                CB_data[MM1][3] = 9
            elif MM1 == 10:
                CB_data[MM1][1] = 16
                CB_data[MM1][2] = 10
                CB_data[MM1][3] = 11
            elif MM1 == 12:
                CB_data[MM1][1] = 18
                CB_data[MM1][2] = 12
                CB_data[MM1][3] = 13
            # ==========================判斷螺栓沉頭直徑&厚度==========================

        elif float(strip_parameter_list[17]) < float(strip_parameter_list[20]):
            #  stop_plate_height < stripper_plate_height

            # ==========================依模板厚度決定螺栓直徑大小==========================
            if float(strip_parameter_list[20]) < 15:
                MM1 = 5  # 螺栓直徑
            elif 15 < float(strip_parameter_list[20]) < 22:
                MM1 = 6
            elif 22 < float(strip_parameter_list[20]) < 27:
                MM1 = 8
            elif 27 < float(strip_parameter_list[20]) < 35:
                MM1 = 10
            elif float(strip_parameter_list[20]) > 35:
                MM1 = 12
            # ==========================依模板厚度決定螺栓直徑大小==========================

            # ==========================判斷螺栓沉頭直徑&厚度==========================
            if MM1 == 5:
                CB_data[5][1] = 8  # 沉頭直徑
                CB_data[5][2] = 5  # 沉頭厚度
                CB_data[5][3] = 5.5  # 頸部直徑
            elif MM1 == 6:
                CB_data[6][1] = 10
                CB_data[6][2] = 6
                CB_data[6][3] = 7
            elif MM1 == 8:
                CB_data[8][1] = 13
                CB_data[8][2] = 8
                CB_data[8][3] = 9
            elif MM1 == 10:
                CB_data[10][1] = 16
                CB_data[10][2] = 10
                CB_data[10][3] = 11
            elif MM1 == 12:
                CB_data[12][1] = 18
                CB_data[12][2] = 12
                CB_data[12][3] = 13
            # ==========================判斷螺栓沉頭直徑&厚度==========================

        # ==========================判斷螺栓挖孔尺寸==========================
        if MM1 == 5:
            Bolt_data[1][2] = 5.5  # 頸部直徑
            Bolt_data[2][2] = 8  # 沉頭深
            Bolt_data[3][2] = float(strip_parameter_list[17]) + float(strip_parameter_list[20])
            Bolt_data[4][2] = 9  # 沉頭孔直徑
            Bolt_data[5][2] = 80  # 螺栓與螺栓間的間距
        elif MM1 == 6:
            Bolt_data[1][2] = 7
            Bolt_data[2][2] = 9
            Bolt_data[3][2] = float(strip_parameter_list[17]) + float(strip_parameter_list[20])
            Bolt_data[4][2] = 11
            Bolt_data[5][2] = 80
        elif MM1 == 8:
            Bolt_data[1][2] = 9
            Bolt_data[2][2] = 11
            Bolt_data[3][2] = float(strip_parameter_list[17]) + float(strip_parameter_list[20])
            Bolt_data[4][2] = 15
            Bolt_data[5][2] = 100
        elif MM1 == 10:
            Bolt_data[1][2] = 11
            Bolt_data[2][2] = 13
            Bolt_data[3][2] = float(strip_parameter_list[17]) + float(strip_parameter_list[20])
            Bolt_data[4][2] = 17
            Bolt_data[5][2] = 125
        elif MM1 == 12:
            Bolt_data[1][2] = 13
            Bolt_data[2][2] = 15
            Bolt_data[3][2] = float(strip_parameter_list[17]) + float(strip_parameter_list[20])
            Bolt_data[4][2] = 19
            Bolt_data[5][2] = 150
        # ==========================判斷螺栓挖孔尺寸==========================

        BoltQuantity[2] = (-int(-int(length2.Value - 55 * 2) / Bolt_data[5][2]) + 1) * 2  # 螺栓孔數量

        stop_plate_machining_instructions_bolt_hole = 0
        for j in range(1, BoltQuantity[2] + 1):
            stop_plate_machining_instructions_bolt_hole += 1  # 加工說明
            C1 = 55
            C2 = 20

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (BoltQuantity[2] / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity[2] / 2 + 1):
                X_Coordinate = C1 + Bolt_data[5][2] * (j - BoltQuantity[2] / 2 - 1)
            elif j != 1 and j != (BoltQuantity[2] / 2 + 1):
                X_Coordinate = C1 + Bolt_data[5][2] * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[i] / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length[i] - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width[i] - C2
            #     n =7
            # elif j == 5:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length[i] - 15
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 11
            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, float(
                strip_parameter_list[17]) - (Bolt_data[2][2] - CB_data[MM1][2]))
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()

            # ==========================螺栓第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, float(
                strip_parameter_list[17]) - Bolt_data[2][2])  # 方向點距離
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Locking_dir_point_" + str(j)
            part1.Update()
            # ==========================螺栓第二點(決定方向點)==========================
        # ==========================建螺栓點==========================

        # ==========================挖孔==========================
        body1 = bodies1.Add()
        body1.Name = "bolt_remove_Body"
        part1.Update()

        stop_plate_bolt_number = 0

        for k in range(1, BoltQuantity[2] + 1):
            hybridShapeFactory1 = part1.HybridShapeFactory
            shapeFactory1 = part1.ShapeFactory
            body2 = bodies1.Item("Body.2")
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(k))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              Bolt_data[1][2] / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference3, reference4, True,
                                                                              Bolt_data[4][2] / 2)
            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference5, length3.Value)
            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference6)
            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference7)
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pocket2 = shapeFactory1.AddNewPocketFromRef(reference8, -Bolt_data[2][2])
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference9)
            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference10)
            part1.Update()

        part1.InWorkObject = body2
        remove1 = shapeFactory1.AddNewRemove(body1)
        stop_plate_bolt_number += 1
        remove1.Name = "Stop-plate-Bolt-" + str(stop_plate_bolt_number)
        part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "A: " + str(stop_plate_machining_instructions_bolt_hole) + "-M" + str(
            Bolt_data[1][2]) + "鑽穿, 正面沉頭 %%C" + str(Bolt_data[2][2]) + "深" + str(Bolt_data[3][2]) + "mm(上模螺絲)"

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

    Bolt_data12 = Bolt_data[1][2]
    Bolt_data22 = Bolt_data[2][2]
    Bolt_data42 = Bolt_data[4][2]
    Bolt_data52 = Bolt_data[5][2]
    BoltQuantity2 = BoltQuantity[2]

    return Bolt_data12, Bolt_data22, Bolt_data42, Bolt_data52, BoltQuantity2


def Stripper(MM1, Bolt_data11, Bolt_data21, Bolt_data41, Bolt_data51, BoltQuantity2, CB_data_MM12):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

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

        # ==========================建螺栓點==========================
        Stripper_machining_instructions_bolt_hole = 0
        for j in range(1, BoltQuantity2 + 1):
            Stripper_machining_instructions_bolt_hole += 1  # 加工說明

            C1 = 55
            C2 = 20

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (BoltQuantity2 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity2 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - BoltQuantity2 / 2 - 1)
            elif j != 1 and j != (BoltQuantity2 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[i] / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length[i] - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width[i] - C2
            #     n =7
            # elif j == 5:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length[i] - 15
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 11
            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          Bolt_data21 - CB_data_MM12)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()
        # ==========================建螺栓點==========================

        body2 = bodies1.Item("Body.2")
        shapeFactory1 = part1.ShapeFactory

        stripper_plate_bolt_number = 0

        # ==========================挖孔==========================
        for k in range(1, BoltQuantity2 + 1):
            body1 = bodies1.Add()
            body1.Name = "bolt_remove_Body"
            part1.Update()

            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(k))

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.ThreadingMode = 0
            hole1.CreateStandardThreadDesignTable(1)
            strParam1 = hole1.HoleThreadDescription
            hole1.BottomType = 2
            hole1.Reverse()
            strParam1.Value = "M" + str(MM1)  # 螺栓大小
            length1 = hole1.ThreadDepth
            length1.Value = float(strip_parameter_list[20])

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference6 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitMode = 3
            limit1.LimitingElement = reference6
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            stripper_plate_bolt_number += 1
            remove1.Name = "Stripper-plate-Bolt-" + str(stripper_plate_bolt_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "A: " + str(
            Stripper_machining_instructions_bolt_hole) + "-M" + str(Bolt_data11) + str(MM1) + "攻穿"

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


def lower_die():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "lower_die_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("lower_die_" + str(i) + "\\plate_length")
        plate_length = length1.Value  # 159.4
        length2 = parameters1.Item("lower_die_" + str(i) + "\\plate_width")
        plate_width = length2.Value   # 390
        length3 = parameters1.Item("lower_die_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        if 200 <= length1.Value < 1000:
            Quantity = 6

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

        # ==========================依模板厚度決定螺栓直徑大小==========================
        if lower_die_set_height < 15:
            MM2 = 5  # 螺栓直徑
        elif 15 < lower_die_set_height < 22:
            MM2 = 6
        elif 22 < lower_die_set_height < 27:
            MM2 = 8
        elif 27 < lower_die_set_height < 35:
            MM2 = 10
        elif lower_die_set_height > 35:
            MM2 = 12
        # ==========================依模板厚度決定螺栓直徑大小==========================

        # ==========================判斷螺栓沉頭直徑&厚度==========================
        if MM2 == 5:
            CB_data[MM2][1] = 8  # 沉頭直徑
            CB_data[MM2][2] = 5  # 沉頭厚度
        elif MM2 == 6:
            CB_data[MM2][1] = 10
            CB_data[MM2][2] = 6
        elif MM2 == 8:
            CB_data[MM2][1] = 13
            CB_data[MM2][2] = 8
        elif MM2 == 10:
            CB_data[MM2][1] = 16
            CB_data[MM2][2] = 10
        elif MM2 == 12:
            CB_data[MM2][1] = 18
            CB_data[MM2][2] = 12
        # ==========================判斷螺栓沉頭直徑&厚度==========================
        if MM2 == 5:
            Bolt_data[1][3] = 5.5  # 頸部直徑
            Bolt_data[2][3] = 8  # 沉頭深
            Bolt_data[3][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height  # 沉頭深+螺紋總長
            Bolt_data[4][3] = 9  # 沉頭孔直徑
            Bolt_data[5][3] = 80  # 螺栓與螺栓間的間距
        elif MM2 == 6:
            Bolt_data[1][3] = 7
            Bolt_data[2][3] = 9
            Bolt_data[3][3] = slower_die_cavity_plate_height + lower_pad_height + lower_die_set_height
            Bolt_data[4][3] = 11
            Bolt_data[5][3] = 80
        elif MM2 == 8:
            Bolt_data[1][3] = 9
            Bolt_data[2][3] = 11
            Bolt_data[3][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height
            Bolt_data[4][3] = 15
            Bolt_data[5][3] = 100
        elif MM2 == 10:
            Bolt_data[1][3] = 11
            Bolt_data[2][3] = 13
            Bolt_data[3][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height
            Bolt_data[4][3] = 17
            Bolt_data[5][3] = 125
        elif MM2 == 12:
            Bolt_data[1][3] = 13
            Bolt_data[2][3] = 15
            Bolt_data[3][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height
            Bolt_data[4][3] = 19
            Bolt_data[5][3] = 150
        # ==========================判斷螺栓挖孔尺寸==========================

        # ==========================建螺栓點==========================
        lower_die_machining_instructions_bolt_hole = 0
        BoltQuantity[3] = ((-int(-int(plate_width - 55 * 2) / Bolt_data[5][3])) + 1) * 2  # 螺栓孔數量
        for j in range(1, BoltQuantity[3] + 1):
            lower_die_machining_instructions_bolt_hole += 1
            C1 = 55
            C2 = 10

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (BoltQuantity[3] / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity[3] / 2 + 1):
                X_Coordinate = C1 + Bolt_data[5][3] * (j - BoltQuantity[3] / 2 - 1)
            elif j != 1 and j != (BoltQuantity[3] / 2 + 1):
                X_Coordinate = C1 + Bolt_data[5][3] * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[i] / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length[i] - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 7
            # elif j == 5:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length[i] - 70
            #     Y_Coordinate = plate_width[i] - C2

            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()

            # ==========================螺栓第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(
                X_Coordinate, Y_Coordinate, Bolt_data[2][3])  # 方向點距離(沉頭深度)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Locking_dir_point_" + str(j)
            part1.Update()
            # ==========================螺栓第二點(決定方向點)==========================
        # ==========================建螺栓點==========================

        # ==========================挖孔==========================
        R = 0
        lower_die_bolt_number = 0
        for k in range(1, BoltQuantity[3] + 1):
            R += 1
            document = catapp.ActiveDocument
            part1 = document.Part
            bodies1 = part1.Bodies
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            shapeFactory1 = part1.ShapeFactory
            body2 = bodies1.Item("Body.2")
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(
                reference1, reference2, True, Bolt_data[4][3] / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.InWorkObject.Name = "bolt_D_circle" + str(R)
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(
                reference3, reference4, True, Bolt_data[1][3] / 2)
            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.InWorkObject.Name = "bolt_d_circle" + str(R)
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pad1 = shapeFactory1.AddNewPadFromRef(reference5, Bolt_data[2][3])
            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pad1.SetProfileElement(reference6)
            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pad1.SetProfileElement(reference7)
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pad2 = shapeFactory1.AddNewPadFromRef(reference8, 11)
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pad2.SetProfileElement(reference9)
            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pad2.SetProfileElement(reference10)

            limit1 = pad2.FirstLimit
            limit1.LimitMode = 3
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference11 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference11
            part1.Update()
            part1.InWorkObject = body2

            remove1 = shapeFactory1.AddNewRemove(body1)

            lower_die_bolt_number += 1
            remove1.Name = "Lower-die-Bolt-spacer-" + str(lower_die_bolt_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "A: " + str(lower_die_machining_instructions_bolt_hole) + "-M" + str(
            Bolt_data[1][3]) + "鑽穿, 正面沉頭 %%C" + str(Bolt_data[4][3] / 2) + "深" + str(
            Bolt_data[2][3]) + "mm(下模螺絲)"  # 加工說明

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

    Bolt_data13 = Bolt_data[1][3]
    Bolt_data23 = Bolt_data[2][3]
    Bolt_data33 = Bolt_data[3][3]
    Bolt_data43 = Bolt_data[4][3]
    Bolt_data51 = Bolt_data[5][3]
    BoltQuantity3 = BoltQuantity[3]
    CB_data_MM21 = CB_data[MM2][1]
    CB_data_MM22 = CB_data[MM2][2]
    return MM2, Bolt_data13, Bolt_data23, Bolt_data33, Bolt_data43, Bolt_data51, BoltQuantity3, CB_data_MM21, CB_data_MM22


def lower_pad(Bolt_data13, Bolt_data51, BoltQuantity3):
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

        # ==========================建螺栓點==========================
        lower_pad_machining_instructions_bolt_hole = 0
        for j in range(1, BoltQuantity3 + 1):
            lower_pad_machining_instructions_bolt_hole += 1
            C1 = 55
            C2 = 20

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - BoltQuantity3 / 2 - 1)
            elif j != 1 and j != (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = plate_length[i] - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 7
            # elif j == 5:
            #     X_Coordinate = plate_length[i] / 2
            #     Y_Coordinate = plate_width[i] - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = plate_length[i] - 70
            #     Y_Coordinate = plate_width[i] - C2

            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標
            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()

        # ==========================挖孔==========================
        R = 0
        lower_pad_bolt_number = 0
        for k in range(1, BoltQuantity3 + 1):
            R += 1
            document = catapp.ActiveDocument
            part1 = document.Part
            bodies1 = part1.Bodies
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            shapeFactory1 = part1.ShapeFactory
            body2 = bodies1.Item("Body.2")
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0

            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length1 = hole1.Diameter
            length1.Value = Bolt_data13  # 通孔直徑
            hole1.BottomType = 2
            limit1.LimitMode = 3

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_plane")
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference3
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            lower_pad_bolt_number += 1
            remove1.Name = "Lower-pad-Bolt-spacer-" + str(lower_pad_bolt_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")

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


def lower_die_set(MM1, Bolt_data51, BoltQuantity3):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    partDocument1 = document.Open(save_path + "lower_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters

    length1 = parameters1.Item("plate_length")
    plate_length = length1.Value
    length2 = parameters1.Item("plate_width")
    plate_width = length2.Value
    length3 = parameters1.Item("plate_height")
    if length3.Value > 0:
        plate_height = length3.Value
        plate_height = 80
    elif length3.Value < 0:
        plate_height = -length3.Value
        plate_height = -80
    part1.Update()

    q = 0
    for g in range(1, plate_line_number + 1):
        n = g - 1
        if n > 0:
            length4 = parameters1.Item("Spacing_" + str(n))
            Spacing = length4.Value
        elif n == 0:
            Spacing = 0
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

        # ==========================建螺栓點==========================
        lower_die_set_machining_instructions_bolt_hole = 0
        for j in range(1, BoltQuantity3 + 1):
            lower_die_set_machining_instructions_bolt_hole += 1
            q += 1
            b = Spacing + length1.Value
            if g == 1:
                b = 0
            length6 = parameters1.Item("lower_die_set_lower_die_X")
            lower_die_set_lower_die_X = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_X = length6.Value
            length6 = parameters1.Item("lower_die_set_lower_die_Y")
            lower_die_set_lower_die_Y = length6.Value
            length6 = parameters1.Item("lower_up_die_set_X")
            lower_up_die_set_Y = length6.Value

            C1 = 55 + lower_die_set_lower_die_X + lower_up_die_set_X + b
            C2 = 20 + lower_die_set_lower_die_Y  # + lower_up_die_set_Y

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1
                Y_Coordinate = length1.Value - C2
            elif j > (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - BoltQuantity3 / 2 - 1)
            elif j != 1 and j != (BoltQuantity3 / 2 + 1):
                X_Coordinate = C1 + Bolt_data51 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + plate_length[g] / 2 - 55
            #     Y_Coordinate = C2
            # elif j == 3:
            #     X_Coordinate = C1 + 2 * (plate_length[g] / 2 - 55)
            #     Y_Coordinate = C2
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = lower_die_set_plate_width - C2
            # elif j == 5:
            #     X_Coordinate = C1 + plate_length[g] / 2 - 55
            #     Y_Coordinate = lower_die_set_plate_width - C2
            # elif j == 6:
            #     X_Coordinate = C1 + 2 * (plate_length[g] / 2 - 55) - 15
            #     Y_Coordinate = lower_die_set_plate_width - C2

            bolt_point_X[j] = X_Coordinate  # 孔位置X座標
            bolt_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Locking_point_" + str(j)
            part1.Update()

        # ==========================挖孔==========================
        R = 0
        part1 = partDocument1.Part
        bodies1 = part1.Bodies

        body1 = bodies1.Add()
        body1.Name = "body_remove"
        part1.Update()

        shapeFactory1 = part1.ShapeFactory
        body2 = bodies1.Item("PartBody")
        hybridShapes1 = body2.HybridShapes

        lower_die_set_bolt_number = 0
        for k in range(1, BoltQuantity3 + 1):
            R += 1
            hybridShapePointCoord1 = hybridShapes1.Item("Locking_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, lower_die_set_height)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0

            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0

            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.ThreadingMode = 0
            hole1.CreateStandardThreadDesignTable(1)

            strParam1 = hole1.HoleThreadDescription
            hole1.BottomType = 2
            strParam1.Value = "M" + str(MM1)  # 螺栓大小

            length1 = limit1.dimension
            length2 = hole1.ThreadDepth
            length2.Value = lower_die_set_height - 2
            part1.Update()

        part1.InWorkObject = body2
        remove1 = shapeFactory1.AddNewRemove(body1)

        lower_die_set_bolt_number += 1
        remove1.Name = "Lower-die-set-Bolt-" + str(lower_die_set_bolt_number)
        part1.Update()
        # ==========================挖孔==========================

    strParam1 = parameters1.Item("Properties\\A")

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


Plate_Locking_Hole()

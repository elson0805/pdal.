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

Pin_data = [[0] * 5 for i in range(5)]
Pinhole_data = [[0] * 5 for j in range(5)]
PinQuantity = [0] * 9
pin_point_X = [0] * 9
pin_point_Y = [0] * 9


def Plate_Pin_Hole():  # 模板螺栓挖孔
    (Pin_data11, Pin_data21, Pinhole_data11, Pinhole_data21, Pinhole_data31, PinQuantity1) = upper_die_set()  # 上模座 1
    up_plate(Pinhole_data31, PinQuantity1)  # 上墊板 1
    splint(Pin_data11, Pinhole_data31, PinQuantity1)  # 上夾板 1
    (Pin_data12, Pin_data22, Pinhole_data12, Pinhole_data22, Pinhole_data32, PinQuantity2) = stop_plate()  # 止擋板 2
    Stripper(Pin_data12, Pinhole_data32, PinQuantity2)  # 脫料板 2
    (Pin_data13, Pin_data23, Pinhole_data13, Pinhole_data23, Pinhole_data33, PinQuantity3) = lower_die()  # 下模板 3
    lower_pad(Pin_data13, Pinhole_data33, PinQuantity3)  # 下墊板 3
    lower_die_set(Pin_data13, Pin_data23, Pinhole_data13, Pinhole_data23, Pinhole_data33, PinQuantity3)  # 下模座 3


def upper_die_set():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    partDocument1 = document.Open(save_path + "upper_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters

    length1 = parameters1.Item("plate_length")
    plate_length = length1.Value  # 450

    length2 = parameters1.Item("plate_width")
    plate_width = length2.Value   # 309.4

    length3 = parameters1.Item("plate_height")
    if length3.Value > 0:         # 50
        plate_height = length3.Value
    elif length3.Value < 0:
        plate_height = -length3.Value
    part1.Update()

    q = 0
    R = 0

    document = catapp.ActiveDocument
    part1 = document.Part
    bodies1 = part1.Bodies

    for g in range(1, plate_line_number + 1):
        n = g - 1
        if n > 0:
            length4 = parameters1.Item("Spacing_" + str(n))
            Spacing = length4.Value
        elif n == 0:
            Spacing = 0

        # ==========================依模板尺寸決定Pin直徑大小==========================
        if plate_length < 500:
            Pinsize = 10  # Pin直徑
        elif 500 < plate_length < 1000:
            Pinsize = 10
        elif 1000 < plate_length < 1500:
            Pinsize = 12
        elif 1500 < plate_length < 2000:
            Pinsize = 12
        elif plate_length > 2000:
            Pinsize = 12
        # ==========================依模板尺寸決定Pin直徑大小==========================

        # ==========================判斷Pin長度、Pin孔尺寸==========================
        if Pinsize == 10:
            Pin_data[1][1] = 10  # pin直徑
            Pin_data[2][1] = (upper_die_set_height + upper_pad_height + splint_height) * 0.66  # pin長度
            Pinhole_data[1][1] = 12  # pin沉孔直徑
            Pinhole_data[2][1] = upper_die_set_height + upper_pad_height + splint_height  # pin孔總深度
            Pinhole_data[3][1] = 250  # pin與pin間的間距
        elif Pinsize == 12:
            Pin_data[1][1] = 12
            Pin_data[2][1] = (upper_die_set_height + upper_pad_height + splint_height) * 0.66
            Pinhole_data[1][1] = 14
            Pinhole_data[2][1] = upper_die_set_height + upper_pad_height + splint_height
            Pinhole_data[3][1] = 300
        # ==========================判斷螺栓沉頭直徑&厚度==========================

        # ==========================建基準點==========================
        body1 = bodies1.Item("PartBody")
        part1.InWorkObject = body1

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

        # ==========================建Pin點==========================
        upper_die_set_machining_instructions_pin_hole = 0
        PinQuantity[1] = (-int(-(int(length1.Value) - 15 * 2) / Pinhole_data[3][1]) + 1) * 2  # Pin孔數量
        for j in range(1, PinQuantity[1] + 1):
            upper_die_set_machining_instructions_pin_hole += 1  # 加工說明
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

            C1 = int(15 + lower_die_set_lower_die_X + lower_up_die_set_X + b) + 1
            # 基礎值(固定)+下模座與模板X方向間距+下模座與上模座X方向間距+模板與模板X方向間距(複數模板)
            C2 = int(55 + lower_die_set_lower_die_Y + lower_up_die_set_Y) + 1
            # 基礎值(固定)+下模座與模板Y方向間距+下模座與上模座Y方向間距

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (PinQuantity[1] / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length2.Value - C2
            elif j > (PinQuantity[1] / 2 + 1):
                X_Coordinate = C1 + int(Pinhole_data[3][1]) * (j - PinQuantity[1] / 2 - 1)
            elif j != 1 and j != (PinQuantity[1] / 2 + 1):
                X_Coordinate = C1 + int(Pinhole_data[3][1]) * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + length1.Value / 2 - 30
            #     Y_Coordinate = C2
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = length2.Value - C2
            # elif j == 4:
            #     X_Coordinate = C1 + length1.Value - 30
            #     Y_Coordinate = length2.Value - C2

            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(q)
            part1.Update()
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        body1 = bodies1.Add()
        body1.Name = "pin_remove_Body"

        upper_die_set_bolt_number = 0
        for k in range(1, PinQuantity[1] + 1):
            body2 = bodies1.Item("PartBody")
            shapeFactory1 = part1.ShapeFactory
            hybridShapes1 = body2.HybridShapes
            hybridShapeFactory1 = part1.HybridShapeFactory
            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(k))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(
                reference1, reference2, True, Pinhole_data[1][1] / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(
                reference3, reference4, True, Pin_data[1][1] / 2)
            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference5, upper_die_set_height * 0.5)

            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference6)

            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference7)
            pocket1.DirectionOrientation = 0
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pocket2 = shapeFactory1.AddNewPocketFromRef(reference8, upper_die_set_height)
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference9)

            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference10)
            pocket2.DirectionOrientation = 0
            part1.Update()

        part1.InWorkObject = body2
        remove1 = shapeFactory1.AddNewRemove(body1)
        upper_die_set_bolt_number += 1
        remove1.Name = "Upper-die-set-Pin-" + str(upper_die_set_bolt_number)
        part1.Update()
        # ==========================挖孔==========================
    strParam1 = parameters1.Item("Properties\\HP")
    # strParam1.Value = "HP: " + str(upper_die_set_machining_instructions_pin_hole) + "-%%C" + str(
    #     length1.Value * 2) + "鉸 ,正面沉頭 %%C 12.0 深(合銷CNC鉸鑽)"  # 加工說明

    selection1 = partDocument1.Selection
    selection1.Clear()
    selection1.Search("Name=*_point_*, All")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()
    selection1.Search("Name=*Sketch.*, All")
    selection1.VisProperties.SetShow(1)
    selection1.Clear()

    time.sleep(1)
    partDocument1.save()
    partDocument1.Close()

    Pin_data11 = Pin_data[1][1]
    Pin_data21 = Pin_data[2][1]
    Pinhole_data11 = Pinhole_data[1][1]
    Pinhole_data21 = Pinhole_data[2][1]
    Pinhole_data31 = Pinhole_data[3][1]
    PinQuantity1 = PinQuantity[1]

    return Pin_data11, Pin_data21, Pinhole_data11, Pinhole_data21, Pinhole_data31, PinQuantity1


def up_plate(Pinhole_data31, PinQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "up_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("up_plate_" + str(i) + "\\plate_length")  # 159.4
        plate_length = length1.Value

        length2 = parameters1.Item("up_plate_" + str(i) + "\\plate_width")   # 390
        plate_width = length2.Value

        length3 = parameters1.Item("up_plate_" + str(i) + "\\plate_height")
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

        # ==========================建Pin點==========================
        up_plate_machining_instructions_pin_hole = 0

        for j in range(1, PinQuantity1 + 1):
            up_plate_machining_instructions_pin_hole += 1  # 加工說明

            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = plate_length - C2
            elif j > (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data31 * (j - PinQuantity1 / 2 - 1)
            elif j != 1 and j != (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data31 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + Pinhole_data31 / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = Pinhole_data31 - C1
            #     Y_Coordinate = C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = C1
            #     Y_Coordinate = Pinhole_data31 - C2
            #     n =7
            # elif j == 5:
            #     X_Coordinate = Pinhole_data31 / 2
            #     Y_Coordinate = Pinhole_data31 - C2
            #     n = 9
            # elif j == 6:
            #     X_Coordinate = Pinhole_data31 - 15
            #     Y_Coordinate = Pinhole_data31 - C2
            #     n = 11
            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        shapeFactory1 = part1.ShapeFactory
        body2 = bodies1.Item("Body.2")
        up_plate_bolt_number = 0
        for k in range(1, PinQuantity1 + 1):
            R += 1
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            # part1.Update()

            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length1 = hole1.Diameter
            length1.Value = Pin_data[1][1] + 0.5  # 通孔直徑
            hole1.Reverse()
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            limit1.LimitMode = 3

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference3
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            up_plate_bolt_number += 1
            remove1.Name = "Upper-plate-Pin-" + str(up_plate_bolt_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\HP")
        strParam1.Value = "HP: " + str(up_plate_machining_instructions_pin_hole) + "-%%C" + str(
            length1.Value) + "鑽穿(合銷)"

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


def splint(Pin_data11, Pinhole_data31, PinQuantity1):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    R = 0
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

        # ==========================建Pin點==========================
        splint_machining_instructions_pin_hole = 0
        for j in range(1, PinQuantity1 + 1):
            splint_machining_instructions_pin_hole += 1  # 加工說明

            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length1.Value - 55
            elif j > (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data31 * (j - PinQuantity1 / 2 - 1)
            elif j != 1 and j != (PinQuantity1 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data31 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = Pinhole_data31 / 2
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = Pinhole_data31 - C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = Pinhole_data31 - C1
            #     Y_Coordinate = Pinhole_data31 - C2
            #     n =7

            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()

            # ==========================螺栓第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          splint_height)  # 方向點距離(沉頭深度)
            reference3 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference3
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Pin_dir_point_" + str(j)
            part1.Update()
            # ==========================螺栓第二點(決定方向點)==========================
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        splint_pin_number = 0
        for k in range(1, PinQuantity1 + 1):
            R += 1

            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            body2 = bodies1.Item("Body.2")
            hybridShapeFactory1 = part1.HybridShapeFactory
            shapeFactory1 = part1.ShapeFactory
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 15)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0

            length4 = hole1.Diameter
            length4.Value = Pin_data[1][1]  # 直徑
            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            hole1.Reverse()

            length5 = limit1.dimension
            length5.Value = splint_height  # 深度
            part1.Update()

            part1.InWorkObject = body2
            body2 = bodies1.Item("body_remove_" + str(R))
            remove1 = shapeFactory1.AddNewRemove(body2)
            splint_pin_number += 1
            remove1.Name = "Splint-Pin-" + str(splint_pin_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\HP")
        strParam1.Value = "HP: " + str(splint_machining_instructions_pin_hole) + "-%%C" + str(
            Pin_data11) + "割, 單+0.005(合銷)"

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


def stop_plate():
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    R = 0
    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "Stop_plate_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters

        length1 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_length")
        plate_length = length1.Value   # 159.4
        length2 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_width")
        plate_width = length2.Value   # 390
        length3 = parameters1.Item("Stop_plate_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        # ==========================依模板尺寸決定Pin直徑大小==========================
        if plate_length < 500:
            Pinsize = 10  # Pin直徑
        elif 500 < plate_length < 1000:
            Pinsize = 10
        elif 1000 < plate_length < 1500:
            Pinsize = 13
        elif 1500 < plate_length < 2000:
            Pinsize = 13
        elif plate_length > 2000:
            Pinsize = 13
        # ==========================依模板尺寸決定Pin直徑大小==========================

        # ==========================判斷Pin長度、Pin孔尺寸==========================
        if Pinsize == 10:
            Pin_data[1][2] = 10  # pin直徑
            Pin_data[2][2] = 10 * 2 * 2  # pin長度
            Pinhole_data[1][2] = 10 + 2  # pin沉孔直徑
            Pinhole_data[2][2] = stop_plate_height + stripper_plate_height  # pin孔總深度
            Pinhole_data[3][2] = 250  # pin與pin間的間距
        elif Pinsize == 13:
            Pin_data[1][2] = 13
            Pin_data[2][2] = 13 * 2 * 2
            Pinhole_data[1][2] = 13 + 2
            Pinhole_data[2][2] = stop_plate_height + stripper_plate_height
            Pinhole_data[3][2] = 300
        # ==========================判斷Pin長度、Pin孔尺寸==========================

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

        # ==========================建Pin點==========================
        # PinQuantity[2] = (-int(-(int(length1.Value) - 15 * 2) / Pinhole_data[3][1]) + 1) * 2
        PinQuantity[2] = 4

        stop_plate_machining_instructions_pin_hole = 0
        for j in range(1, PinQuantity[2] + 1):
            stop_plate_machining_instructions_pin_hole += 1  # 加工說明
            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity[2] / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length1.Value - C2
            elif j > (PinQuantity[2] / 2 + 1):
                X_Coordinate = C1 + Pinhole_data[3][2] * (j - PinQuantity[2] / 2 - 1)
            elif j != 1 and j != (PinQuantity[2] / 2 + 1):
                X_Coordinate = C1 + Pinhole_data[3][2] * (j - 1)
            # elif j == 2:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = length1.Value - C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = length2.Value - C2
            #     n =7
            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, stop_plate_height)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()

            # ==========================Pin第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          stop_plate_height - 8)  # 方向點距離
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Pin_dir_point_" + str(j)
            part1.Update()
            # ==========================Pin第二點(決定方向點)==========================
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        stop_plate_pin_number = 0
        R = 0
        for k in range(1, PinQuantity[2] + 1):  # 1~4
            R += 1
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            body2 = bodies1.Item("Body.2")
            hybridShapeFactory1 = part1.HybridShapeFactory
            shapeFactory1 = part1.ShapeFactory
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_down_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              Pinhole_data[1][2] / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.InWorkObject.Name = "pin_D_circle" + str(R)
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference3, reference4, True,
                                                                              Pin_data[1][2] / 2)
            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject.Name = "pin_d_circle" + str(R)
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pad1 = shapeFactory1.AddNewPadFromRef(reference5, 8)
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
            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_up_plane")
            reference11 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference11
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            stop_plate_pin_number += 1
            remove1.Name = "Stop-plate-Pin-" + str(stop_plate_pin_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\HP")
        strParam1.Value = "HP: " + str(stop_plate_machining_instructions_pin_hole) + "-%%C" + str(
            Pin_data[1][2]) + "割, 單+0.01 ,正面鑽孔 %%C 14.0深(合銷)"

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

    Pin_data12 = Pin_data[1][2]
    Pin_data22 = Pin_data[2][2]
    Pinhole_data12 = Pinhole_data[1][2]
    Pinhole_data22 = Pinhole_data[2][2]
    Pinhole_data32 = Pinhole_data[3][2]
    PinQuantity2 = PinQuantity[2]

    return Pin_data12, Pin_data22, Pinhole_data12, Pinhole_data22, Pinhole_data32, PinQuantity2


def Stripper(Pin_data12, Pinhole_data32, PinQuantity2):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
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

        # ==========================建Pin點==========================
        Stripper_machining_instructions_pin_hole = 0
        for j in range(1, PinQuantity2 + 1):
            Stripper_machining_instructions_pin_hole += 1  # 加工說明

            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity2 / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length1.Value - C2
            elif j > (PinQuantity2 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data32 * (j - PinQuantity2 / 2 - 1)
            elif j != 1 and j != (PinQuantity2 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data32 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = length1.Value - C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = length2.Value - C2
            #     n =7
            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        stripper_plate_pin_number = 0
        for k in range(1, PinQuantity2 + 1):
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            body2 = bodies1.Item("Body.2")

            shapeFactory1 = part1.ShapeFactory
            hybridShapes1 = body2.HybridShapes
            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(k))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")
            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            hole1.ThreadingMode = 1
            length1 = hole1.Diameter
            length1.Value = Pin_data12
            hole1.Reverse()
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            limit1.LimitMode = 3

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_die_plate_down_plane")
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference3
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            stripper_plate_pin_number += 1
            remove1.Name = "Stripper-plate-Pin-" + str(stripper_plate_pin_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "HP: " + str(Stripper_machining_instructions_pin_hole) + "-%%C" + str(
            length1.Value) + "割, 單+0.005(合銷)"

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
        plate_width = length1.Value   # 390
        length3 = parameters1.Item("lower_die_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        # ==========================依模板尺寸決定Pin直徑大小==========================
        if plate_length < 500:
            Pinsize = 10  # Pin直徑
        elif 500 < plate_length < 1000:
            Pinsize = 10
        elif 1000 < plate_length < 1500:
            Pinsize = 12
        elif 1500 < plate_length < 2000:
            Pinsize = 12
        elif plate_length > 2000:
            Pinsize = 12
        # ==========================依模板尺寸決定Pin直徑大小==========================

        # ==========================判斷Pin長度、Pin孔尺寸==========================
        if Pinsize == 10:
            Pin_data[1][3] = 10  # pin直徑
            Pin_data[2][3] = (lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height) * 0.66  # pin長度
            Pinhole_data[1][3] = 10 + 2  # pin沉孔直徑
            Pinhole_data[2][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height  # pin孔總深度
            Pinhole_data[3][3] = 250  # pin與pin間的間距
        elif Pinsize == 12:
            Pin_data[1][3] = 12
            Pin_data[2][3] = (lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height) * 0.66
            Pinhole_data[1][3] = 12 + 2
            Pinhole_data[2][3] = lower_die_cavity_plate_height + lower_pad_height + lower_die_set_height
            Pinhole_data[3][3] = 300
        # ==========================判斷Pin長度、Pin孔尺寸==========================

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

        # ==========================建Pin點==========================
        lower_die_machining_instructions_pin_hole = 0
        PinQuantity[3] = ((-int(-int(length1.Value - 15 * 2) / Pinhole_data[3][3])) + 1) * 2
        for j in range(1, PinQuantity[3] + 1):
            lower_die_machining_instructions_pin_hole += 1
            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity[3] / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length1.Value - C2
            elif j > (PinQuantity[3] / 2 + 1):
                X_Coordinate = C1 + Pinhole_data[3][3] * (j - PinQuantity[3] / 2 - 1)
            elif j != 1 and j != (PinQuantity[3] / 2 + 1):
                X_Coordinate = C1 + Pinhole_data[3][3] * (j - 1)
            # elif j == 2:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = length1.Value - C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = length1.Value - C1
            #     Y_Coordinate = length2.Value - C2
            #     n =7

            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate,
                                                                          -(lower_die_cavity_plate_height * 1 / 2))
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()

            # ==========================Pin第二點(決定方向點)==========================
            hybridShapePointCoord2 = hybridShapeFactory1.AddNewPointCoord(
                X_Coordinate, Y_Coordinate, -lower_die_cavity_plate_height)
            reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord2.PtRef = reference5
            body1.InsertHybridShape(hybridShapePointCoord2)
            part1.InWorkObject = hybridShapePointCoord2
            part1.InWorkObject.Name = "Pin_dir_point_" + str(j)
            part1.Update()
            # ==========================Pin第二點(決定方向點)==========================
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        R = 0
        lower_die_pin_number = 0
        for k in range(1, PinQuantity[3] + 1):
            R += 1
            document = catapp.ActiveDocument
            part1 = document.Part
            bodies1 = part1.Bodies
            body1 = bodies1.Add()
            body1.Name = "body_remove_" + str(R)
            part1.Update()

            body2 = bodies1.Item("Body.2")
            hybridShapeFactory1 = part1.HybridShapeFactory
            shapeFactory1 = part1.ShapeFactory
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_die_plate_up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 10)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0
            hole1.ThreadingMode = 1
            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length4 = hole1.Diameter
            length4.Value = Pin_data[1][3]
            length5 = limit1.dimension
            length5.Value = lower_die_cavity_plate_height
            part1.InWorkObject.Name = "Pin_Hole_" + str(k)
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            lower_die_pin_number += 1
            remove1.Name = "Lower-die-Bolt-spacer-" + str(lower_die_pin_number)
            part1.Update()
        # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\HP")
        strParam1.Value = "HP: " + str(lower_die_machining_instructions_pin_hole) + "-%%C" + str(
            Pin_data[1][3]) + "割, 單+0.005(合銷)"  # 加工說明

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

    Pin_data13 = Pin_data[1][3]
    Pin_data23 = Pin_data[2][3]
    Pinhole_data13 = Pinhole_data[1][3]
    Pinhole_data23 = Pinhole_data[2][3]
    Pinhole_data33 = Pinhole_data[3][3]
    PinQuantity3 = PinQuantity[3]
    return Pin_data13, Pin_data23, Pinhole_data13, Pinhole_data23, Pinhole_data33, PinQuantity3


def lower_pad(Pin_data13, Pinhole_data33, PinQuantity3):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents

    for i in range(1, plate_line_number + 1):
        partDocument1 = document.Open(save_path + "lower_pad_" + str(i) + ".CATPart")
        part1 = partDocument1.Part
        parameters1 = part1.Parameters
        length1 = parameters1.Item("lower_pad_" + str(i) + "\\plate_length")
        plate_length = length1.Value  # 159.4
        length2 = parameters1.Item("lower_pad_" + str(i) + "\\plate_width")
        plate_width = length2.Value   # 390
        length3 = parameters1.Item("lower_pad_" + str(i) + "\\plate_height")
        if length3.Value > 0:
            plate_height = length3.Value
        elif length3.Value < 0:
            plate_height = -length3.Value
        part1.Update()

        if 200 <= plate_width < 1000:
            Quantity = 4

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

        # ==========================建Pin點==========================
        lower_pad_machining_instructions_pin_hole = 0
        for j in range(1, PinQuantity3 + 1):
            lower_pad_machining_instructions_pin_hole += 1
            C1 = 15
            C2 = 55

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
                n = 1
            elif j == (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = plate_length - C2
            elif j > (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data33 * (j - PinQuantity3 / 2 - 1)
            elif j != 1 and j != (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data33 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = plate_width - C1
            #     Y_Coordinate = C2
            #     n = 3
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = plate_length - C2
            #     n = 5
            # elif j == 4:
            #     X_Coordinate = plate_width - C1
            #     Y_Coordinate = plate_length - C2
            #     n = 7

            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標
            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        R = 0
        lower_pad_pin_number = 0
        for k in range(1, PinQuantity3 + 1):
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

            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("up_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hole1 = shapeFactory1.AddNewHoleFromRefPoint(reference1, reference2, 32)
            hole1.Type = 0
            hole1.AnchorMode = 0
            hole1.BottomType = 0

            limit1 = hole1.BottomLimit
            limit1.LimitMode = 0
            length1 = hole1.Diameter
            length1.Value = Pin_data13 + 0.5

            hole1.ThreadingMode = 1
            hole1.ThreadSide = 0
            hole1.BottomType = 2
            limit1.LimitMode = 3

            hybridShapePlaneOffset2 = hybridShapes1.Item("down_plane")
            reference3 = part1.CreateReferenceFromObject(hybridShapePlaneOffset2)
            limit1.LimitingElement = reference3
            part1.Update()

            part1.InWorkObject = body2
            remove1 = shapeFactory1.AddNewRemove(body1)
            lower_pad_pin_number += 1
            remove1.Name = "Lower-pad-Bolt-spacer-" + str(lower_pad_pin_number)
            part1.Update()
            # ==========================挖孔==========================

        strParam1 = parameters1.Item("Properties\\A")
        strParam1.Value = "HP: " + str(lower_pad_machining_instructions_pin_hole) + "-%%C" + str(
            length1.Value) + "割, 單+0.01(合銷)"

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


def lower_die_set(Pin_data13, Pin_data23, Pinhole_data13, Pinhole_data23, Pinhole_data33, PinQuantity3):
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.Documents
    partDocument1 = document.Open(save_path + "lower_die_set.CATPart")
    part1 = partDocument1.Part
    parameters1 = part1.Parameters

    length1 = parameters1.Item("plate_length")
    plate_length = length1.Value  # 309.4
    length2 = parameters1.Item("plate_width")
    plate_width = length2.Value   # 450
    length3 = parameters1.Item("plate_height")
    if length3.Value > 0:
        plate_height = length3.Value
        plate_height = 80
    elif length3.Value < 0:
        plate_height = -length3.Value
        plate_height = -80
    part1.Update()

    document = catapp.ActiveDocument
    part1 = document.Part
    bodies1 = part1.Bodies

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

        # ==========================建Pin點==========================
        lower_die_set_machining_instructions_pin_hole = 0
        for j in range(1, PinQuantity3 + 1):
            lower_die_set_machining_instructions_pin_hole += 1
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

            C1 = 15 + lower_die_set_lower_die_X + lower_up_die_set_X + b
            C2 = 55 + lower_die_set_lower_die_Y  # + lower_up_die_set_Y

            if j == 1:
                X_Coordinate = C1
                Y_Coordinate = C2
            elif j == (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + 30
                Y_Coordinate = length1.Value - C2
            elif j > (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data33 * (j - PinQuantity3 / 2 - 1)
            elif j != 1 and j != (PinQuantity3 / 2 + 1):
                X_Coordinate = C1 + Pinhole_data33 * (j - 1)
            # elif j == 2:
            #     X_Coordinate = C1 + length1.Value - 30
            #     Y_Coordinate = C2
            # elif j == 3:
            #     X_Coordinate = C1
            #     Y_Coordinate = lower_die_set_plate_width - C2
            # elif j == 4:
            #     X_Coordinate = C1 + length1.Value - 30
            #     Y_Coordinate = lower_die_set_plate_width - C2

            pin_point_X[j] = X_Coordinate  # 孔位置X座標
            pin_point_Y[j] = Y_Coordinate  # 孔位置Y座標

            hybridShapePointCoord1 = hybridShapeFactory1.AddNewPointCoord(X_Coordinate, Y_Coordinate, 0)
            reference2 = part1.CreateReferenceFromObject(hybridShapeExtremum1)
            hybridShapePointCoord1.PtRef = reference2
            body1.InsertHybridShape(hybridShapePointCoord1)
            part1.InWorkObject = hybridShapePointCoord1
            part1.InWorkObject.Name = "Pin_point_" + str(j)
            part1.Update()
        # ==========================建Pin點==========================

        # ==========================挖孔==========================
        part1 = partDocument1.Part
        bodies1 = part1.Bodies
        body1 = bodies1.Add()
        body1.Name = "pin_remove_Body"
        part1.Update()

        R = 0
        lower_die_set_pin_number = 0
        for k in range(1, PinQuantity3 + 1):
            R += 1

            shapeFactory1 = part1.ShapeFactory
            body2 = bodies1.Item("PartBody")
            hybridShapes1 = body2.HybridShapes

            hybridShapePointCoord1 = hybridShapes1.Item("Pin_point_" + str(R))
            hybridShapePlaneOffset1 = hybridShapes1.Item("down_plane")

            reference1 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference2 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)

            hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference1, reference2, True,
                                                                              Pin_data13 / 2)
            hybridShapeCircleCtrRad1.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad1)
            part1.InWorkObject = hybridShapeCircleCtrRad1
            part1.Update()

            reference3 = part1.CreateReferenceFromObject(hybridShapePointCoord1)
            reference4 = part1.CreateReferenceFromObject(hybridShapePlaneOffset1)
            hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference3, reference4, True,
                                                                              Pinhole_data13 / 2)

            hybridShapeCircleCtrRad2.SetLimitation(1)
            body1.InsertHybridShape(hybridShapeCircleCtrRad2)
            part1.InWorkObject = hybridShapeCircleCtrRad2
            part1.Update()

            reference5 = part1.CreateReferenceFromName("")
            pocket1 = shapeFactory1.AddNewPocketFromRef(reference5, upper_die_set_height)

            reference6 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference6)

            reference7 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
            pocket1.SetProfileElement(reference7)
            part1.Update()

            reference8 = part1.CreateReferenceFromName("")
            pocket2 = shapeFactory1.AddNewPocketFromRef(reference8, lower_die_set_height * 0.5)
            reference9 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference9)

            reference10 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)
            pocket2.SetProfileElement(reference10)
            part1.Update()

        part1.InWorkObject = body2
        remove1 = shapeFactory1.AddNewRemove(body1)

        lower_die_set_pin_number += 1
        remove1.Name = "Lower-die-set-Bolt-" + str(lower_die_set_pin_number)
        part1.Update()
        # ==========================挖孔==========================

    strParam1 = parameters1.Item("Properties\\HP")
    strParam1.Value = "HP: " + str(lower_die_set_machining_instructions_pin_hole) + "-%%C" + str(
        length1.Value * 2) + "鑽穿 ,背面逃孔 %%C　12.0深(合銷CNC鑽銷)"

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


Plate_Pin_Hole()

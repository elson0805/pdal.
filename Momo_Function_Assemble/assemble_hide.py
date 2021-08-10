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


def assemble_hide():  # 隱藏組力拘束
    catapp = win32.Dispatch('CATIA.Application')
    document = catapp.ActiveDocument
    # ===============================搜尋拘束==============================
    selection1 = document.Selection
    selection1.Search("Name=Constraints,all")
    visPropertySet1 = selection1.VisProperties
    visPropertySet1 = visPropertySet1.Parent
    bSTR1 = visPropertySet1.Name
    visPropertySet1.SetShow(1)
    selection1.Clear()
    # ===============================搜尋拘束==============================

    # ===============================搜尋組立起點==============================
    selection1.Search("Name=bned_up_forming_boit_point_*,all")
    visPropertySet2 = selection1.VisProperties
    bSTR2 = visPropertySet2.Name
    visPropertySet2.SetShow(1)
    selection1.Clear()
    # ===============================搜尋組立起點==============================


assemble_hide()

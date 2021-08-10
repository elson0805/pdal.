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
all_part_name = ['', 'op20_cut_punch_01', 'op30_cut_punch_01', 'op30_cut_punch_02', 'op30_cut_punch_03',
                 'op30_cut_punch_04', 'op40_cut_punch_01', 'op60_cut_punch_01', 'op60_cut_punch_02',
                 'op60_cut_punch_03', 'op60_cut_punch_04', 'op70_cut_punch_01', 'op10_SJAS_6_6_01', 'op10_SJAS_6_6_02',
                 'Binder_Plate_30', 'Binder_Plate_40', 'Binder_Plate_50', "op10_A_punch_insert_02",
                 "op20_cut_cavity_insert_01", "op30_cut_cavity_insert_01", "op40_cut_cavity_insert_01",
                 "op60_cut_cavity_insert_01", "op70_cut_cavity_insert_01", "op10_A_punch_QR_Splint_insert_02",
                 "op10_A_punch_QR_Stripper_insert_02", "lower_die_1", "lower_pad_1", "Stripper_1", "Stop_plate_1",
                 "Splint_1", "up_plate_1", "lower_die_set", "upper_die_set"]
all_part_number = 32

# =====================將所有零件名稱存至文字檔=========================
x = (save_path + "all_output_part_name.txt")
with open(x, 'w') as f:
    # f.write(all_part_number)
    all_part_name[0] = str(all_part_number)
    for i in all_part_name:
        f.writelines(i)
        f.writelines('\n')
# =====================將所有零件名稱存至文字檔=========================

catapp = win32.Dispatch('CATIA.Application')
documents1 = catapp.Documents
productDocument1 = documents1.Add("Product")
product1 = productDocument1.Product
products1 = product1.Products

# =====================取特定行=========================
f = open(save_path + "all_output_part_name.txt", 'r')
vntLines = []
for line in f:  # 讀取記事本內每行內容
    vntLines.append(line.strip("\n"))
# =====================取特定行=========================

# =====================匯入各零件檔=========================
arrayOfVariantOfBSTR1 = ['']
i = int()
for j in vntLines:
    if j == str(all_part_number):
        continue
    arrayOfVariantOfBSTR1.append(save_path + j + ".CATPart")
    i += 1
products1Variant = products1
products1Variant.AddComponentsFromFiles(arrayOfVariantOfBSTR1, "All")
product1.Update()
# =====================匯入各零件檔=========================

# =====================重整組立視角=========================
specsAndGeomWindow1 = catapp.ActiveWindow
viewer3D1 = specsAndGeomWindow1.ActiveViewer
viewer3D1.Reframe()
viewpoint3D1 = viewer3D1.Viewpoint3D
# =====================重整組立視角=========================

productDocument1.SaveAs(save_path + "Product1.CATProduct")  # 存檔

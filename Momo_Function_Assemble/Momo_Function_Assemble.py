import csv
import win32com.client as win32
import openpyxl
import os
import Locking_assemble
import Pin_assemble
import Inner_Guiding_post_assemble
import SBT_assemble
import Plate_out_Guide_posts
import out_Guide_posts_locking_assemble
import limiting_assembly
import allotype_punch_assemble
import assemble_hide

catapp = win32.Dispatch('CATIA.Application')
document = catapp.ActiveDocument


# with open(Locking_assemble) as f:  # 螺栓
#     exec(f.read())
#
# with open(Pin_assemble) as f:  # 合銷
#     exec(f.read())
#
# with open(Inner_Guiding_post_assemble) as f:  # 內導柱/套
#     exec(f.read())
#
# with open(SBT_assemble) as f:  # 等高螺栓
#     exec(f.read())

# with open(Pilot_Punch_assemble) as f:  # 引導沖
#     exec(f.read())
#
# with open(Stripper_punch_assemble) as f:  # 脫料釘
#     exec(f.read())

# with open(Plate_out_Guide_posts) as f:  # 外導柱/套
#     exec(f.read())
#
# with open(out_Guide_posts_locking_assemble) as f:  # 外導柱螺栓
#     exec(f.read())

# with open(Lifter_assemble) as f:  # 浮升銷(水箱蓋無
#     exec(f.read())
#
# with open(sensor_assemble) as f:  # 感測器螺栓/合銷
#     exec(f.read())
#
# with open(Nithgen_spring_assemble) as f:  # 氮氣彈簧
#     exec(f.read())
#
# with open(keyway_assemble) as f:  # 向上折彎定位鍵
#     exec(f.read())

# with open(limiting_assembly) as f:  # 限位柱
#     exec(f.read())

# with open(leveling_block_assembly) as f:  # 整平塊
#     exec(f.read())
#
# with open(Bending_punch_up_assembly) as f:  # 向上折彎
#     exec(f.read())
#
# with open(Bending_punch_down_assembly) as f:  # 向下折彎
#     exec(f.read())

# with open(allotype_punch_assemble) as f:  # 異型沖頭
#     exec(f.read())
#
# with open(assemble_hide) as f:  # 隱藏組力拘束
#     exec(f.read())

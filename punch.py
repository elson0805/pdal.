import csv
import win32com.client as win32
import openpyxl

# catapp = win32.Dispatch('CATIA.Application')
# document = catapp.ActiveDocument

# g = now_plate_line_number
for now_op_number in range(1, total_op_number + 1):
    n = now_op_number
    op_number = 10 * n
    # ==========補強沖頭==========
    if plate_line_Reinforcement_cut_line[g][n] > 0:
        with open(QR_punch_Reinforcement) as f:
            exec(f.read())

    # ==========下料切斷沖頭_下==========
    if plate_line_cut_punch_d_cutting_number[g][n] > 0:
        with open(punch_d_cutting) as f:
            exec(f.read())

    # ==========下料切斷沖頭_上==========
    if plate_line_cut_punch_u_cutting_number[g][n] > 0:
        with open(cut_punch_u_cutting) as f:
            exec(f.read())

    # ==========沖切沖頭==========
    if plate_line_cut_line_number[g][n] > 0:
        with open(cut_punch) as f:
            exec(f.read())

    # ==========肩形台階沖頭==========
    if plate_line_cut_line_number[g][n] > 0:
        with open(Riveting_Punch) as f:
            exec(f.read())

    # ==========T形異形沖==========
    if plate_line_unnomal_cut_line_T_number[g][n] > 0:
        with open(unnomal_cut_punch_T) as f:
            exec(f.read())

    # ==========I形異形沖==========
    if plate_line_unnomal_cut_line_I_number[g][n] > 0:
        with open(unnomal_cut_punch_I) as f:
            exec(f.read())

    # ==========M形異形沖==========
    if plate_line_unnomal_cut_line_M_number[g][n] > 0:
        with open(unnomal_cut_punch_M) as f:
            exec(f.read())

    # ==========成形沖頭_沖頭==========
    if plate_line_forming_punch_surface_number[g][n] > 0:
        with open(forming_cavity) as f:
            exec(f.read())

    # ==========成形沖頭_模穴==========
    if plate_line_forming_cavity_surface_number[g][n] > 0:
        with open(forming_punch) as f:
            exec(f.read())

    # ==========異型沖頭==========
    if plate_line_allotype_cut_line_number[g][n] > 0:
        with open(allotype_cut_punch) as f:
            exec(f.read())

    # ==========↓ 快拆沖頭==========
    # ==========沖切沖頭_右==========
    if plate_line_right_quickly_remove_cut_line_number[g][n] > 0:
        with open(quickly_remove_punch) as f:
            exec(f.read())

    # ==========沖切沖頭_左==========
    if plate_line_left_quickly_remove_cut_line_number[g][n] > 0:
        with open(quickly_remove_punch) as f:
            exec(f.read())

    # ==========沖切沖頭_上==========
    if plate_line_up_quickly_remove_cut_line_number[g][n] > 0:
        with open(quickly_remove_punch) as f:
            exec(f.read())

    # ==========沖切沖頭_下==========
    if plate_line_down_quickly_remove_cut_line_number[g][n] > 0:
        with open(quickly_remove_punch) as f:
            exec(f.read())

    # # ==========折彎沖頭_右==========
    # if plate_line_right_quickly_remove_bending_surface_number[g][n] > 0:
    #     with open(quickly_remove_punch) as f:
    #         exec(f.read())
    #
    # # ==========折彎沖頭_左==========
    # if plate_line_left_quickly_remove_bending_surface_number[g][n] > 0:
    #     with open(quickly_remove_punch) as f:
    #         exec(f.read())
    #
    # # ==========折彎沖頭_上==========
    # if plate_line_up_quickly_remove_bending_surface_number[g][n] > 0:
    #     with open(quickly_remove_punch) as f:
    #         exec(f.read())
    #
    # # ==========折彎沖頭_下==========
    # if plate_line_down_quickly_remove_bending_surface_number[g][n] > 0:
    #     with open(quickly_remove_punch) as f:
    #         exec(f.read())
    # ==========↑快拆沖頭==========

    # ==========打凸包沖頭_左==========
    if plate_line_emboss_forming_punch_left_surface_number[g][n] > 0:
        with open(emboss_forming_punch_left) as f:
            exec(f.read())

    # ==========打凸包沖頭_右==========
    if plate_line_emboss_forming_punch_right_surface_number[g][n] > 0:
        with open(emboss_forming_punch_right) as f:
            exec(f.read())

    # ==========半沖切沖頭==========
    if plate_line_half_cut_line_number[g][n] > 0:
        with open(half_cut_punch) as f:
            exec(f.read())

    # if plate_line_A_punch_number[g][n] > 0:
    #     with open() as f:
    #         exec(f.read())

    # if plate_line_bending_punch_surface[g][n] > 0:  # 折彎
    #     with open() as f:
    #         exec(f.read())

    # ==========整形沖頭==========
    # if plate_line_bending_punch_surface_number[g][n] > 0:
    #     with open(F_bending_punch) as f:
    #         exec(f.read())

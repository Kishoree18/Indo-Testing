Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Threading
Imports Inventor
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports Org.BouncyCastle.Math.EC
Public Class Main_Panel


    'Public Shared connectionString As String = "Server=CTPL-HW-183\SQLEXPRESS02;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123" 'OUR URL
    'MY URL
    'Public Shared connectionString As String =
    '"Server=tcp:CTPL-HW-198,1433;" &
    '"Initial Catalog=Indo_tech;" &
    '"User ID=sqladmin;" &
    '"Password=Cadopt@123;" &
    '"Connect Timeout=10;"

    'INFENT URL
    Public Shared connectionString As String =
    "Server=tcp:CTPL-HW-183,1433;" &
    "Initial Catalog=Indo_tech;" &
    "User ID=sqladmin;" &
    "Password=Cadopt@1234#Sql;" &
    "Encrypt=True;" &
    "TrustServerCertificate=True;" &
    "Connect Timeout=10;"

    'MAGENDRAN URL
    'Public Shared connectionString As String =
    '"Server=HP-DES-25\SQLEXPRESS,1433;" &
    '"Initial Catalog=Indo_tech;" &
    '"User ID=sqladmin;" &
    '"Password=indotech_app;" &
    '"Encrypt=True;" &
    '"TrustServerCertificate=True;" &
    '"Connect Timeout=10;"

    'ARAVINTH URL
    'Public Shared connectionString As String =
    '"Server=INDO-DES-001\SQLEXPRESS,1433;" &
    '"Initial Catalog=Indo_tech;" &
    '"User ID=sqladmin;" &
    '"Password=indotech_app;" &
    '"Encrypt=True;" &
    '"TrustServerCertificate=True;" &
    '"Connect Timeout=10;"

    'Public Shared connectionString As String = "Server=192.168.8.80,1433;Database=Indo_tech;User ID=sa;Password=Indo@123" 'Indotech CURRENT URL

    'INVENTOR

    Public Shared Timer As Timer

    Public Shared ThisApplication As Inventor.Application
    Public Shared user As String
    Public Shared INFO As Inventor.Document
    Public Shared BOTTOM_TANK_PARAMETER As Inventor.Document
    Public Shared Directory As String
    Public Shared libraryPath As String

    Public Shared replaceUB As New ArrayList()

    'project info
    Public Shared dwg_titleV As String
    Public Shared wk_odr_NoV As String
    Public Shared customerV As String
    Public Shared ratingV As String
    Public Shared voltageV As String
    Public Shared OutputlocationV As String
    Public Shared workorder_list As String

    Public Shared SelectoriginaldirectoryV As String


    'Basic design data
    Public Shared core_limp As String
    Public Shared core_dia As String
    Public Shared stack_thk As String
    Public Shared max_step As String
    Public Shared leg_centre_1 As String
    Public Shared leg_centre_2 As String
    Public Shared end_limb_1 As String
    Public Shared Tp_btm_el_1 As String
    Public Shared window_ht As String
    Public Shared Max_winding_od As String
    Public Shared vector_group As String
    Public Shared type_of_cooling As String
    Public Shared transform_ft_view As String
    Public Shared OLTC_pkt_position As String

    Public Shared Gasket_type As String
    Public Shared Flange_type As String


    'bottom tank

    Public Shared btm_tk_typeV As String
    Public Shared btm_tk_curb_proV As String
    Public Shared curb_sizev As String
    Public Shared TT_curb_sizeV As String
    Public Shared btm_plt_htV As String
    Public Shared btm_tk_i_lenV As String
    Public Shared btm_plt_lenV As String
    Public Shared btm_plt_thkV As String
    Public Shared btm_tk_i_wdhV As String
    Public Shared btm_plt_widthV As String
    Public Shared bt_tk_wdh_2V As String
    Public Shared Side_Plt_thkV As String
    Public Shared Cus_htV As String

    Public Shared OLTC_pck_lenV As String
    Public Shared OLTC_pck_wdhV As String
    Public Shared tk_curb_wdhV As String
    Public Shared Curb_wdh_LVV As String
    Public Shared Curb_OLTC_EndV As String
    Public Shared Curb_wdh_DV As String
    Public Shared BP_X1V As String
    Public Shared BP_X2V As String
    Public Shared BP_X3V As String
    Public Shared BP_R1V As String
    Public Shared GPX1V As String
    Public Shared GPX2V As String
    Public Shared GPY1V As String
    Public Shared GPY2V As String
    Public Shared BTM_TK_Lifting_lugsV As String


    Public Shared btm_len_chkbxV As String
    Public Shared btm_plt_wdt_ck_bxV As String



    'UB
    Public Shared No_of_RailsV As String
    Public Shared No_of_Rails1V As String
    Public Shared CC_LENGTHV As String
    Public Shared roller_gauge_cmbV As String
    Public Shared roller_orientationV As String
    Public Shared Roller_XaxisV As String
    Public Shared Roller_YaxisV As String
    Public Shared Roller_CustomV As String


    Public Shared MP_lenV As String
    Public Shared MP_WthV As String
    Public Shared MP_ThkV As String
    Public Shared BS_RV As String
    Public Shared BS_LV As String
    Public Shared Dim_L1V As String
    Public Shared T2V As String
    'Public Shared Dim_T1V As String

    Public Shared BS1_WDH_1V As String
    Public Shared BS2_WDH_1V As String
    Public Shared BS3_WDH_1V As String

    Public Shared BS1_HT_WSV As String
    Public Shared BS2_HT_WSV As String
    Public Shared BS3_HT_WSV As String

    Public Shared BS1_THK_1V As String
    Public Shared BS2_THK_1V As String
    Public Shared BS3_THK_1V As String

    Public Shared prob_support_cmbV As String
    Public Shared prob_noV As String
    Public Shared PROB_HTV As String

    Public Shared Prob_X1axisV As String
    Public Shared Prob_X2axisV As String
    Public Shared Prob_Y1axisV As String
    Public Shared Prob_Y2axisV As String

    Public Shared HS_CC_lenV As String
    Public Shared Distance_AV As String
    Public Shared Distance_BV As String
    Public Shared XV As String

    'Curb Details
    Public Shared curb_bolt_sizeV As String
    Public Shared HV_pitch_sizeV As String
    Public Shared LV_pitch_sizeV As String
    Public Shared OLTC1_pitch_sizeV As String
    Public Shared OLTC2_pitch_sizeV As String
    Public Shared OLTC3_pitch_sizeV As String
    Public Shared OLTC_OPP_pitch_sizeV As String
    Public Shared HV_HPV As String
    Public Shared LV_HPV As String
    Public Shared OLTC1_HPV As String
    Public Shared OLTC2_HPV As String
    Public Shared OLTC3_HPV As String
    Public Shared OLTC_OPP_HPV As String

    Public Shared T2HVPA1 As String
    Public Shared T2HVPAB1 As String
    Public Shared T2HVPAB2 As String
    Public Shared T2HVPAB3 As String


    Public Shared curb_lenV As String
    Public Shared curb_widV As String
    Public Shared bush_pkt_X1axisV As String
    Public Shared bush_pkt_X2axisV As String
    Public Shared bush_pkt_Y1axisV As String
    Public Shared curb_R1V As String


    'FLATSTIFFNER
    'HV
    Public Shared HV_STNR_TYPEV As String
    Public Shared HV_FS_AV As String
    Public Shared HV_FS_BV As String
    Public Shared HV_FS_A1V As String
    Public Shared HV_FS_B1V As String

    'LV
    Public Shared LV_STNR_TYPEV As String
    Public Shared LV_FS_AV As String
    Public Shared LV_FS_BV As String
    Public Shared LV_FS_A1V As String
    Public Shared LV_FS_B1V As String

    'OLTC S1
    Public Shared OLTC_STNR_TYPEV As String
    Public Shared OLTC_S1_AV As String
    Public Shared OLTC_S1_BV As String
    Public Shared OLTC_S1_A1V As String
    Public Shared OLTC_S1_B1V As String
    Public Shared OLTC_S1_TV As String

    Public Shared OS1D1V As String
    Public Shared OS1D2V As String

    'OLTC S2
    Public Shared OLTC_STNR2_TYPEV As String
    Public Shared OLTC_S2_AV As String
    Public Shared OLTC_S2_BV As String
    Public Shared OLTC_S2_A1V As String
    Public Shared OLTC_S2_B1V As String

    Public Shared OS2D1V As String
    Public Shared OS2D2V As String

    'OLTC END
    Public Shared OLTC_END_STNR_TYPV As String
    Public Shared OLTC_END_AV As String
    Public Shared OLTC_END_BV As String
    Public Shared OLTC_END_A1V As String
    Public Shared OLTC_END_B1V As String

    Public Shared OED1V As String
    Public Shared OED2V As String
    Public Shared OED3V As String
    Public Shared OED4V As String

    Public Shared Stiffner_count_BT As String
    Public Shared oltc_stiffner1_count As String
    Public Shared oltc_stiffner2_count As String

    Public Shared stiffner_count_HV As String
    Public Shared HVD1 As String
    Public Shared HVD2 As String
    Public Shared HVD3 As String
    Public Shared HVD4 As String

    Public Shared stiffner_count_LV As String
    Public Shared LVD1 As String
    Public Shared LVD2 As String
    Public Shared LVD3 As String
    Public Shared LVD4 As String



    'JACKING PAD
    'HV
    Public Shared jack_pad_typeV As String
    Public Shared JP_AV As String
    Public Shared JP_BV As String
    Public Shared JP_CV As String
    Public Shared JP_DV As String
    Public Shared JP_EV As String
    Public Shared JP_GV As String
    Public Shared JP_XV As String
    Public Shared JP_TV As String
    Public Shared JP_B1V As String
    Public Shared JP_BB1V As String
    Public Shared JP_B2V As String

    Public Shared NO_TOP_JP_HVV As String
    Public Shared JP_T_LV As String
    Public Shared JP_T_HV As String
    Public Shared JP_T_TV As String



    Public Shared NO_OF_JP_HVV As String
    Public Shared HV_JPD1V As String
    Public Shared HV_JPD2V As String
    Public Shared HV_JPD3V As String
    Public Shared HV_JPD4V As String
    Public Shared JP_Y_DISTV As String

    'LV
    Public Shared jack_pad_type1 As String
    Public Shared JP_LV_AV As String
    Public Shared JP_LV_BV As String
    Public Shared JP_LV_CV As String
    Public Shared JP_LV_DV As String
    Public Shared JP_LV_EV As String
    Public Shared JP_LV_FV As String
    Public Shared JP_LV_GV As String
    Public Shared JP_LV_HV As String
    Public Shared JP_LV_CHV As String
    Public Shared JP_LV_TV As String


    Public Shared NO_TOP_JP_LVV As String
    Public Shared JP_T_L_LVV As String
    Public Shared JP_T_H_LVV As String
    Public Shared JP_T_T_LVV As String




    Public Shared NO_JP_LVV As String
    Public Shared LV_JP1V As String
    Public Shared LV_JP2V As String
    Public Shared LV_JP3V As String
    Public Shared LV_JP4V As String




    Public Shared LV_STNR_THKV As String
    Public Shared LV_F_STNR_AV As String
    Public Shared LV_F_STNR_BV As String
    Public Shared LV_F_STNR_DIAV As String
    Public Shared LV_F_STNR_R1V As String
    Public Shared LV_F_STNR_R2V As String
    Public Shared LV_F_STNR_C As String
    Public Shared LV_F_STNR_D As String
    Public Shared LV_F_STNR_E As String
    Public Shared LV_F_STNR_F As String
    Public Shared CURB_HOLECUT_LVV As String
    Public Shared LV_CURB_HOLE_WV As String
    Public Shared LV_CURB_HOLE_HV As String
    Public Shared LV_DIM1V As String
    Public Shared LV_DIM2V As String
    Public Shared LV_DIM3V As String




    Public Shared LV_F_STNR_CV As String
    Public Shared LV_F_STNR_DV As String
    Public Shared LV_F_STNR_EV As String
    Public Shared LV_F_STNR_FV As String


    Public Shared OLTC_STFR_THKV As String
    Public Shared NO_OLTC_STNRV As String
    Public Shared OLTC_F_STNR_AV As String
    Public Shared OLTC_F_STNR_BV As String
    Public Shared OLTC_F_STNR_CV As String
    Public Shared OLTC_F_STNR_DV As String
    Public Shared OLTC_F_STNR_RV As String
    Public Shared OLTC_F_STNR_CHV As String
    Public Shared OLTC_DIM1V As String
    Public Shared OLTC_DIM2V As String
    Public Shared OLTC_DIM3V As String


    Public Shared NO_OLTC_STNR2V As String
    Public Shared OLTC_F_STNR2_TV As String
    Public Shared OLTC_F_STNR2_AV As String
    Public Shared OLTC_F_STNR2_BV As String
    Public Shared OLTC_STNR2_DIM1V As String
    Public Shared OLTC_STNR2_DIM2V As String
    Public Shared OLTC_STNR2_DIM3V As String


    Public Shared NO_OLTC_END_STNRV As String

    Public Shared OLTC_END_THKV As String
    Public Shared OLTC_END_DIM1V As String
    Public Shared OLTC_END_DIM2V As String
    Public Shared OLTC_END_DIM3V As String
    Public Shared OLTC_END_DIM4V As String
    Public Shared OLTC_END_DIM5V As String



    'OTHER PARTS 1


    Public Shared SAMP_PIPE_SIZEV As String
    Public Shared SAMP_FLG_TYPEV As String
    Public Shared SAM_VAL_POSV As String
    Public Shared SV_XAXISV As String
    Public Shared SV_YAXISV As String


    Public Shared FILTER_VALUVE_SIZEV As String
    Public Shared FILTER_VAL_POSV As String
    Public Shared FILTRE_FLG_TYPEV As String
    Public Shared FV_XAXISV As String
    Public Shared FV_YAXISV As String



    Public Shared DGA_valve_reqV As String

    Public Shared DRAIN_VALVE_SIZEV As String
    Public Shared DRAIN_FLG_TYPEV As String
    Public Shared DRAIN_VALVE_POSV As String
    Public Shared DV_XAXISV As String
    Public Shared DV_YAXISV As String


    Public Shared DGA_VALVE_SIZEV As String
    Public Shared DGA_VALVE_POSV As String
    Public Shared DGA_FLG_TYPEV As String
    Public Shared DGA_XAXISV As String
    Public Shared DGA_YAXISV As String


    Public Shared NIFPS_REQV As String
    Public Shared NIFPS_MAKEV As String
    Public Shared NIFPS_VALSIZE As String



    Public Shared N2_filling_req_HVV As String
    Public Shared N2_filling_req_LVV As String
    Public Shared N2_filling_req_OLTC_oppV As String
    Public Shared N2_filling_req_OLTCV As String
    Public Shared N2_qty_HVV As String
    Public Shared N2_qty_LVV As String
    Public Shared N2_qty_OLTC_oppV As String
    Public Shared N2_qty_OLTCV As String
    Public Shared HV_X1_AXISV As String
    Public Shared LV_X1_AXISV As String
    Public Shared OLTC_opp_X1_AXISV As String
    Public Shared OLTC_X1_AXISV As String
    Public Shared HV_X2_AXISV As String
    Public Shared LV_X2_AXISV As String
    Public Shared OLTC_opp_X2_AXISV As String
    Public Shared OLTC_X2_AXISV As String
    Public Shared HV_X3_AXISV As String
    Public Shared LV_X3_AXISV As String
    Public Shared OLTC_opp_X3_AXISV As String
    Public Shared OLTC_X3_AXISV As String
    Public Shared HV_Y1_AXISV As String
    Public Shared LV_Y1_AXISV As String
    Public Shared OLTC_opp_Y1_AXISV As String
    Public Shared OLTC_Y1_AXISV As String

    Public Shared GAUDE_PIN_SIZEV As String
    Public Shared GAUDE_PIN_POSV As String
    Public Shared X1V As String
    Public Shared X2V As String
    Public Shared Y1V As String
    Public Shared Y2V As String

    Public Shared EARTHINGPAD_HVV As String
    Public Shared EARTHINGPAD_POSV As String
    Public Shared HV_XAXISV As String
    Public Shared HV_YAXISV As String
    Public Shared LV_XAXISV As String
    Public Shared LV_YAXISV As String
    Public Shared earthing_pad_HV_type As String
    Public Shared earthing_pad_LV_type As String
    Public Shared NO_OF_PORT_BOT As String

    Public Shared COOLER_PIPE_SIZEV As String
    Public Shared Bottom_CP_req_HV As String
    Public Shared Bottom_CP_req_LV As String
    Public Shared Bottom_CP_req_OE As String
    Public Shared Bottom_CP_req_OS As String
    Public Shared NO_OF_CP_HV As String
    Public Shared NO_OF_CP_LV As String
    Public Shared NO_OF_CP_OE As String
    Public Shared NO_OF_CP_OS As String
    Public Shared CP_HV_X1_AXIS As String
    Public Shared CP_LV_X1_AXIS As String
    Public Shared CP_OE_X1_AXIS As String
    Public Shared CP_OS_X1_AXIS As String
    Public Shared CP_HV_X2_AXIS As String
    Public Shared CP_LV_X2_AXIS As String
    Public Shared CP_OE_X2_AXIS As String
    Public Shared CP_OS_X2_AXIS As String
    Public Shared CP_HV_X3_AXIS As String
    Public Shared CP_LV_X3_AXIS As String
    Public Shared CP_OE_X3_AXIS As String
    Public Shared CP_OS_X3_AXIS As String
    Public Shared CP_HV_Y1_AXIS As String
    Public Shared CP_LV_Y1_AXIS As String
    Public Shared CP_OE_Y1_AXIS As String
    Public Shared CP_OS_Y1_AXIS As String

    '=================================================================
    'Top Tank

    Public Shared T_T_CURB_PROV As String
    Public Shared T_T_CURB_SIZEV As String
    Public Shared T_T_CURB_BOLT_SIZEV As String
    Public Shared T_T_PLT_THKV As String
    Public Shared T_IN_LENV As String
    Public Shared T_TANK_WTHV As String
    Public Shared T_TANK_WTH2V As String
    Public Shared TNK_HTV As String
    Public Shared T_TOP_CVR_THKV As String
    Public Shared T_SLOPE_SIDEV As String
    Public Shared tank_slope_valV As String
    Public Shared SLOPE_CUSV As String

    Public Shared NT_XV As String
    Public Shared NT_HV As String

    Public Shared BUSH_POCK_TYPEV As String
    Public Shared Bushing_poc_AV As String
    Public Shared Bushing_poc_BV As String
    Public Shared Bushing_poc_CV As String
    Public Shared BUSH_POC_RADV As String
    Public Shared TAP_CHANGER_TYPEV As String
    Public Shared TAP_makeV As String
    Public Shared TAP_model_noV As String
    Public Shared OLTC_POS_XV As String
    Public Shared OLTC_POS_YV As String
    Public Shared TIE_IN_REQV As String
    Public Shared Tie_in_resistV As String
    Public Shared MOUNTINGV As String
    Public Shared OLTC_POC_LENV As String
    Public Shared OLTC_POC_WDHV As String
    Public Shared OLTC_POC_HTV As String
    Public Shared OLTC_HEAD_VANTV As String
    Public Shared DM_BOX_ORINV As String
    Public Shared UP_TO_CLV As String
    Public Shared VP_TO_CLV As String
    Public Shared WP_TO_CLV As String
    Public Shared LV_TERMV As String
    Public Shared LV_TERM_OPNV As String
    Public Shared LV_T1V As String
    Public Shared LV_T2V As String
    Public Shared OLTC_conservator_type As String

    'HV TERMINAL
    Public Shared HV_Termiantion As String
    Public Shared HV_Termiantion_Mounting As String

    Public Shared HV_bush_rating_U As String
    Public Shared HV_bush_A_UV As String
    Public Shared HV_bush_BCT_U As String
    Public Shared HV_turret_typ_U As String
    Public Shared HV_turret_dia_U As String
    Public Shared HV_turret_ht_U As String
    Public Shared HV_ht_CT_MP_U As String
    Public Shared HV_ht_CT_box_U As String
    Public Shared HV_ID_CT_MP_U As String
    Public Shared HV_bush_MTG_U As String
    Public Shared HV_U_CT_ter_box As String
    Public Shared HV_U_CT_ter_box_angle As String
    Public Shared HV_U_inspection_cover_cmb As String
    Public Shared HV_U_inspection_cover_angle As String
    Public Shared HV_U_equalizing_cover_angle As String
    Public Shared HV_ht_turret_body_plt_U As String
    Public Shared HV_ht_sml_turret_body_U As String
    Public Shared HV_welded_trunk_ht_U As String
    Public Shared HV_turret_ang_U As String

    Public Shared HV_turret_ang_U1 As String
    Public Shared HV_turret_ang_V1 As String
    Public Shared HV_turret_ang_W1 As String

    Public Shared Bush_Mounting_UL As String
    Public Shared LVN1_Mounting As String
    Public Shared LVN2_Mounting As String

    Public Shared Bush_Mounting_UV As String
    Public Shared CT_BOX_UV As String
    Public Shared IC_partno_UV As String

    Public Shared HV_bush_rating_V As String
    Public Shared HV_bush_A_VV As String
    Public Shared HV_bush_BCT_V As String
    Public Shared HV_turret_typ_V As String
    Public Shared HV_turret_dia_V As String
    Public Shared HV_turret_ht_V As String
    Public Shared HV_ht_CT_MP_V As String
    Public Shared HV_ht_CT_box_V As String
    Public Shared HV_ID_CT_MP_V As String
    Public Shared HV_bush_MTG_V As String
    Public Shared HV_V_CT_ter_box As String
    Public Shared HV_V_CT_ter_box_angle As String
    Public Shared HV_V_inspection_cover_cmb As String
    Public Shared HV_V_inspection_cover_angle As String
    Public Shared HV_V_equalizing_cover_angle As String
    Public Shared HV_ht_turret_body_plt_V As String
    Public Shared HV_ht_sml_turret_body_V As String
    Public Shared HV_welded_trunk_ht_V As String
    Public Shared HV_turret_ang_V As String

    Public Shared Bush_Mounting_VL As String

    Public Shared Bush_Mounting_VV As String
    Public Shared CT_BOX_VV As String
    Public Shared IC_partno_VV As String

    Public Shared HV_bush_rating_W As String
    Public Shared HV_bush_A_WV As String
    Public Shared HV_bush_BCT_W As String
    Public Shared HV_turret_typ_W As String
    Public Shared HV_turret_dia_W As String
    Public Shared HV_turret_ht_W As String
    Public Shared HV_ht_CT_MP_W As String
    Public Shared HV_ht_CT_box_W As String
    Public Shared HV_ID_CT_MP_W As String
    Public Shared HV_bush_MTG_W As String
    Public Shared HV_W_CT_ter_box As String
    Public Shared HV_W_CT_ter_box_angle As String
    Public Shared HV_W_inspection_cover_cmb As String
    Public Shared HV_W_inspection_cover_angle As String
    Public Shared HV_W_equalizing_cover_angle As String
    Public Shared HV_ht_turret_body_plt_W As String
    Public Shared HV_ht_sml_turret_body_W As String
    Public Shared HV_welded_trunk_ht_W As String
    Public Shared HV_turret_ang_W As String

    Public Shared Bush_Mounting_wL As String

    Public Shared Bush_Mounting_WV As String
    Public Shared CT_BOX_WV As String
    Public Shared IC_partno_WV As String

    'HVN Termination
    Public Shared HVN_Termination As String
    Public Shared HVN_Termination_Mounting As String
    Public Shared HVN_Termination_X As String
    Public Shared HVN_Termination_Y As String
    Public Shared MOUNTING_HVN As String
    Public Shared MOUNT_LVU As String
    Public Shared MOUNT_LVV As String
    Public Shared MOUNT_LVW As String
    Public Shared turret_incl_ang As String
    Public Shared HVN_TERM_Air_release_plug As String

    Public Shared HVN_MOUNTSIDE As String
    Public Shared HVN_TYP As String



    Public Shared HV_UH As String
    Public Shared HV_VH As String
    Public Shared HV_WH As String

    Public Shared HVN_H As String

    Public Shared LV_UH As String
    Public Shared LV_VH As String
    Public Shared LV_WH As String

    Public Shared LVN1_H As String
    Public Shared LVN2_H As String

    Public Shared HVN_bush_rating_U As String
    Public Shared HVN_bush_BCT_U As String
    Public Shared HVN_turret_ty_U As String
    Public Shared HVN_turret_dia_U As String
    Public Shared HVN_turret_ht_U As String
    Public Shared HVN_ht_CT_MP_U As String
    Public Shared HVN_ht_CT_box_U As String
    Public Shared HVN_ID_CT_MP_U As String
    Public Shared HVN_bush_MTG_pad_U As String
    Public Shared HVN_CT_terminal_U As String
    Public Shared HVN_CT_terminal_ang_U As String
    Public Shared HVN_inspection_cover_U As String
    Public Shared HVN_inspection_cover_ang_U As String
    Public Shared HVN_equaliznig_ang_U As String
    Public Shared HVN_welded_trunk_ht_U As String
    Public Shared HVN_ht_turret_bd_pl_U As String
    Public Shared HVN_ht_small_turret_U As String
    Public Shared HVN_turret_ang_U As String
    Public Shared HVN_CTT_TYPV As String
    Public Shared HVN_IC_TYPV As String

    'LV Termination
    Public Shared LV_Termination As String
    Public Shared LV_Termination_Mounting As String
    Public Shared LV1_termination As String
    Public Shared LV2_termination As String

    Public Shared LV_bush_rating_U As String
    Public Shared LV_bush_BCT_U As String
    Public Shared LV_turret_type_U As String
    Public Shared LV_turret_dia_U As String
    Public Shared LV_turret_ht_U As String
    Public Shared LV_ht_CT_MP_U As String
    Public Shared LV_ht_CT_box_U As String
    Public Shared LV_ID_CT_MP_U As String
    Public Shared LV_bush_MTG_pad_U As String
    Public Shared LV_CT_terminal_box_U As String
    Public Shared LV_CT_terminal_box_ang_U As String
    Public Shared LV_inspection_cover_U As String
    Public Shared LV_inspection_cover_ang_U As String
    Public Shared LV_equalizing_cover_U As String
    Public Shared LV_welded_trunk_ht_U As String
    Public Shared LV_ht_turret_body_U As String
    Public Shared LV_ht_small_turret_U As String
    Public Shared LV_turret_angle_U As String

    Public Shared LV_bush_rating_V As String
    Public Shared LV_bush_BCT_V As String
    Public Shared LV_turret_type_V As String
    Public Shared LV_turret_dia_V As String
    Public Shared LV_turret_ht_V As String
    Public Shared LV_ht_CT_MP_V As String
    Public Shared LV_ht_CT_box_V As String
    Public Shared LV_ID_CT_MP_V As String
    Public Shared LV_bush_MTG_pad_V As String
    Public Shared LV_CT_terminal_box_V As String
    Public Shared LV_CT_terminal_box_ang_V As String
    Public Shared LV_inspection_cover_V As String
    Public Shared LV_inspection_cover_ang_V As String
    Public Shared LV_equalizing_cover_V As String
    Public Shared LV_welded_trunk_ht_V As String
    Public Shared LV_ht_turret_body_V As String
    Public Shared LV_ht_small_turret_V As String
    Public Shared LV_turret_angle_V As String

    Public Shared LV_bush_rating_W As String
    Public Shared LV_bush_BCT_W As String
    Public Shared LV_turret_type_W As String
    Public Shared LV_turret_dia_W As String
    Public Shared LV_turret_ht_W As String
    Public Shared LV_ht_CT_MP_W As String
    Public Shared LV_ht_CT_box_W As String
    Public Shared LV_ID_CT_MP_W As String
    Public Shared LV_bush_MTG_pad_W As String
    Public Shared LV_CT_terminal_box_W As String
    Public Shared LV_CT_terminal_box_ang_W As String
    Public Shared LV_inspection_cover_W As String
    Public Shared LV_inspection_cover_ang_W As String
    Public Shared LV_equalizing_cover_W As String
    Public Shared LV_welded_trunk_ht_W As String
    Public Shared LV_ht_turret_body_W As String
    Public Shared LV_ht_small_turret_W As String
    Public Shared LV_turret_angle_W As String

    Public Shared LV_U1_CT_TYPV As String
    Public Shared LV_V1_CT_TYPV As String
    Public Shared LV_W1_CT_TYPV As String
    Public Shared LV_U1_IC_TYPV As String
    Public Shared LV_V1_IC_TYPV As String
    Public Shared LV_W1_IC_TYPV As String


    Public Shared LVN1_CT_TYP As String
    Public Shared LVN2_CT_TYP As String
    Public Shared LVN1_IC_TYP As String
    Public Shared LVN2_IC_TYP As String



    'BAREBUSHING
    Public Shared LV1_W As String

    Public Shared LV2_W As String

    Public Shared LV1_V As String
    Public Shared LV2_V As String
    Public Shared LV1_U As String
    Public Shared LV2_U As String
    Public Shared LVN1_C As String
    Public Shared LVN2 As String
    Public Shared LVN_C As String
    Public Shared LVBC As String
    Public Shared LVNBC As String
    Public Shared LVN_Mounting As String
    Public Shared LVN_POS As String
    Public Shared LVN_X As String
    Public Shared LVN_Y As String
    Public Shared LV1_TT As String
    Public Shared LV2_TT As String


    'LV1 Cable Box

    Public Shared LVN_CB_Z2 As String
    Public Shared LVN_CB_TP As String
    Public Shared LVN_CB_Z As String
    Public Shared LVN_CB_X As String
    Public Shared LVN_SIDE As String


    Public Shared LV1_CB As String
    Public Shared LV1_CB_Mounting As String
    Public Shared LV1_CB_Position As String
    Public Shared LV1_CB_Clearance As String
    Public Shared LV1_CB_PH As String
    Public Shared LV1_CB_ER As String
    Public Shared LV1_terminat As String
    Public Shared LV2_terminat As String
    Public Shared LV1_trunk_offset As String
    Public Shared LV1_trunk_height As String
    Public Shared LV1_bush_rating_U As String
    Public Shared LV1_bush_mountain_pad_U As String
    Public Shared LV_bushing_U As String
    Public Shared LV1_welded_trunk_ht_U As String
    Public Shared LV1_CT_terminal_box_U As String
    Public Shared LV1_inspec_cover1_U As String
    Public Shared LV1_inspec_cover2_U As String
    Public Shared LV1_top_gusset_U As String
    Public Shared LV1_bot_gusset_U As String
    Public Shared LV1_side1_gusset_U As String
    Public Shared LV1_side2_gusset_U As String

    'LV2 Cable Box
    Public Shared LV2_CB As String
    Public Shared LV2_CB_Mounting As String
    Public Shared LV2_CB_Position As String
    Public Shared LV2_CB_Clearance As String
    Public Shared LV2_CB_PH As String
    Public Shared LV2_CB_ER As String
    Public Shared LV1_term As String
    Public Shared LV2_term As String
    Public Shared LV2_trunk_offset As String
    Public Shared LV2_trunk_ht As String
    Public Shared LV2_bush_rating As String
    Public Shared LV2_bush_mounting_pad As String
    Public Shared LV2_bushing As String
    Public Shared LV2_welded_trunk_ht As String
    Public Shared LV2_CT_terminal_box As String
    Public Shared LV2_inspec_cover1 As String
    Public Shared LV2_inspec_cover2 As String
    Public Shared LV2_top_gusset As String
    Public Shared LV2_bot_gusset As String
    Public Shared LV2_side1_gusset As String
    Public Shared LV2_side2_gusset As String

    'LVN Termination
    Public Shared LVN_Termination As String
    Public Shared LVN_Termination_Mounting As String
    Public Shared LVN1_Termination As String
    Public Shared LVN2_Termination As String

    Public Shared LVN1_bush_rating As String
    Public Shared LVN1_bush_BCT As String
    Public Shared LVN1_turret_type As String
    Public Shared LVN1_turret_dia As String
    Public Shared LVN1_turret_ht As String
    Public Shared LVN1_ht_CT_MP As String
    Public Shared LVN1_ht_CT_box As String
    Public Shared LVN1_ID_CT_MP As String
    Public Shared LVN1_bush_MTG_pad As String
    Public Shared LVN1_CT_terminal_box As String
    Public Shared LVN1_CT_terminal_box_ang As String
    Public Shared LVN1_inspection_cover As String
    Public Shared LVN1_inspection_cover_ang As String
    Public Shared LVN1_equalizing_cover As String
    Public Shared LVN1_welded_trunk_ht As String
    Public Shared LVN1_ht_turret_body As String
    Public Shared LVN1_ht_small_turret_body As String
    Public Shared LVN1_turret_angle As String

    Public Shared LVN2_bush_rating As String
    Public Shared LVN2_bush_BCT As String
    Public Shared LVN2_turret_type As String
    Public Shared LVN2_turret_dia As String
    Public Shared LVN2_turret_ht As String
    Public Shared LVN2_ht_CT_MP As String
    Public Shared LVN2_ht_CT_box As String
    Public Shared LVN2_ID_CT_MP As String
    Public Shared LVN2_bush_MTG_pad As String
    Public Shared LVN2_CT_terminal_box As String
    Public Shared LVN2_CT_terminal_box_ang As String
    Public Shared LVN2_inspection_cover As String
    Public Shared LVN2_inspection_cover_ang As String
    Public Shared LVN2_equalizing_cover As String
    Public Shared LVN2_welded_trunk_ht As String
    Public Shared LVN2_ht_turret_body As String
    Public Shared LVN2_ht_small_turret_body As String
    Public Shared LVN2_turret_angle As String

    'TV Termination
    Public Shared TV_Termination As String
    Public Shared TV_Termination_Mounting As String
    'U
    Public Shared TV_bush_rating_U As String
    Public Shared TV_bush_BCT_U As String
    Public Shared TV_turret_type_U As String
    Public Shared TV_turret_dia_U As String
    Public Shared TV_turret_ht_U As String
    Public Shared TV_ht_CT_MP_U As String
    Public Shared TV_ht_CT_box_U As String
    Public Shared TV_ID_CT_MP_U As String
    Public Shared TV_bush_MTG_U As String
    Public Shared TV_CT_terminal_box_U As String
    Public Shared TV_CT_terminal_box_ang_U As String
    Public Shared TV_inspection_cover_U As String
    Public Shared TV_inspection_cover_ang_U As String
    Public Shared TV_equalizing_cover_U As String
    Public Shared TV_welded_trunk_ht_U As String
    Public Shared TV_ht_turre_body_U As String
    Public Shared TV_ht_small_turret_body_U As String
    Public Shared TV_turret_angle_U As String
    'V1
    Public Shared TV_bush_rating_V1 As String
    Public Shared TV_bush_BCT_V1 As String
    Public Shared TV_turret_type_V1 As String
    Public Shared TV_turret_dia_V1 As String
    Public Shared TV_turret_ht_V1 As String
    Public Shared TV_ht_CT_MP_V1 As String
    Public Shared TV_ht_CT_box_V1 As String
    Public Shared TV_ID_CT_MP_V1 As String
    Public Shared TV_bush_MTG_V1 As String
    Public Shared TV_CT_terminal_box_V1 As String
    Public Shared TV_CT_terminal_box_ang_V1 As String
    Public Shared TV_inspection_cover_V1 As String
    Public Shared TV_inspection_cover_ang_V1 As String
    Public Shared TV_equalizing_cover_V1 As String
    Public Shared TV_welded_trunk_ht_V1 As String
    Public Shared TV_ht_turre_body_V1 As String
    Public Shared TV_ht_small_turret_body_V1 As String
    Public Shared TV_turret_angle_V1 As String
    'V2
    Public Shared TV_bush_rating_V2 As String
    Public Shared TV_bush_BCT_V2 As String
    Public Shared TV_turret_type_V2 As String
    Public Shared TV_turret_dia_V2 As String
    Public Shared TV_turret_ht_V2 As String
    Public Shared TV_ht_CT_MP_V2 As String
    Public Shared TV_ht_CT_box_V2 As String
    Public Shared TV_ID_CT_MP_V2 As String
    Public Shared TV_bush_MTG_V2 As String
    Public Shared TV_CT_terminal_box_V2 As String
    Public Shared TV_CT_terminal_box_ang_V2 As String
    Public Shared TV_inspection_cover_V2 As String
    Public Shared TV_inspection_cover_ang_V2 As String
    Public Shared TV_equalizing_cover_V2 As String
    Public Shared TV_welded_trunk_ht_V2 As String
    Public Shared TV_ht_turre_body_V2 As String
    Public Shared TV_ht_small_turret_body_V2 As String
    Public Shared TV_turret_angle_V2 As String
    'W
    Public Shared TV_bush_rating_W As String
    Public Shared TV_bush_BCT_W As String
    Public Shared TV_turret_type_W As String
    Public Shared TV_turret_dia_W As String
    Public Shared TV_turret_ht_W As String
    Public Shared TV_ht_CT_MP_W As String
    Public Shared TV_ht_CT_box_W As String
    Public Shared TV_ID_CT_MP_W As String
    Public Shared TV_bush_MTG_W As String
    Public Shared TV_CT_terminal_box_W As String
    Public Shared TV_CT_terminal_box_ang_W As String
    Public Shared TV_inspection_cover_W As String
    Public Shared TV_inspection_cover_ang_W As String
    Public Shared TV_equalizing_cover_W As String
    Public Shared TV_welded_trunk_ht_W As String
    Public Shared TV_ht_turre_body_W As String
    Public Shared TV_ht_small_turret_body_W As String
    Public Shared TV_turret_angle_W As String

    'Stiffner1

    Public Shared BS1_A As String
    Public Shared BS1_A1 As String
    Public Shared BS1_B As String
    Public Shared BS1_C As String
    Public Shared BS1_D As String
    Public Shared BS1_E As String
    Public Shared BS1_F As String
    Public Shared BS1_G As String
    Public Shared BS1_H As String
    Public Shared BS1_I As String
    Public Shared BS1_J As String
    Public Shared BS1_K As String

    Public Shared HV_DIM1 As String
    Public Shared HV_DIM2 As String
    Public Shared HV_DIM3 As String
    Public Shared HV_DIM4 As String
    Public Shared HV_DIM5 As String
    Public Shared HV_DIM6 As String
    Public Shared HV_pocket_gusset As String



    Public Shared LV_S1A As String
    Public Shared LV_S1B As String
    Public Shared LV_S1C As String
    Public Shared LV_S1D As String

    Public Shared LV2_S1A As String
    Public Shared LV2_S1B As String
    Public Shared LV2_S1C As String
    Public Shared LV2_S1D As String

    Public Shared LH1_S1A As String
    Public Shared LH1_S1B As String
    Public Shared LH1_S1C As String
    Public Shared LH1_S1D As String

    Public Shared LV3_S1A As String
    Public Shared LV3_S1B As String
    Public Shared LV3_S1C As String
    Public Shared LV3_S1D As String

    Public Shared LH2_S1A As String
    Public Shared LH2_S1B As String
    Public Shared LH2_S1C As String
    Public Shared LH2_S1D As String

    Public Shared LV_D1 As String
    Public Shared LV_D2 As String
    Public Shared LV_D3 As String
    Public Shared LV_D4 As String
    Public Shared LV_D5 As String
    Public Shared LV_D6 As String

    Public Shared OLTC_OPP_box_stiffner As String
    Public Shared OLTC_BS1_QTY As String
    Public Shared OLTC_BS2_QTY As String
    Public Shared OLTC_BS3_QTY As String
    Public Shared OLTC_BS1_WDTH As String
    Public Shared OLTC_BS2_WDTH As String
    Public Shared OLTC_BS3_WDTH As String
    Public Shared OLTC_BS1_HT As String
    Public Shared OLTC_BS2_HT As String
    Public Shared OLTC_BS3_HT As String
    Public Shared OLTC_BS1_THK As String
    Public Shared OLTC_BS2_THK As String
    Public Shared OLTC_BS3_THK As String
    Public Shared OLTC_D1 As String
    Public Shared OLTC_D2 As String
    Public Shared OLTC_D3 As String
    Public Shared OLTC_D4 As String
    Public Shared OLTC_D5 As String
    Public Shared OLTC_D6 As String

    Public Shared HV_LL_cmb As String
    Public Shared LV_LL_cmb As String
    Public Shared OLTC_LL_cmb As String
    Public Shared OLTC_OPP_LL_cmb As String

    Public Shared HV_LL_H1 As String
    Public Shared HV_LL_H2 As String
    Public Shared HV_LL_H3 As String
    Public Shared HV_LL_H4 As String
    Public Shared HV_LL_H5 As String

    Public Shared LV_LL_H1 As String
    Public Shared LV_LL_H2 As String
    Public Shared LV_LL_H3 As String
    Public Shared LV_LL_H4 As String
    Public Shared LV_LL_H5 As String

    Public Shared OLTC_LL_H1 As String
    Public Shared OLTC_LL_H2 As String

    Public Shared OLTC_OPP_LL_H1 As String
    Public Shared OLTC_OPP_LL_H2 As String

    Public Shared gcd_part_cmb As String
    Public Shared gcd_part_x As String
    Public Shared gcd_part_y As String

    Public Shared IC_HV_REQ As String
    Public Shared IC_HV_COUNT As String
    Public Shared IC_HV_X1 As String
    Public Shared IC_HV_X2 As String
    Public Shared IC_HV_X3 As String
    Public Shared IC_HV_Y1 As String
    Public Shared IC_HV_Y2 As String
    Public Shared IC_HV_Y3 As String

    Public Shared IC_HVOP_REQ As String
    Public Shared IC_HVOP_COUNT As String
    Public Shared IC_HVOP_X1 As String
    Public Shared IC_HVOP_X2 As String
    Public Shared IC_HVOP_Y1 As String
    Public Shared IC_HVOP_Y2 As String

    Public Shared IC_LV_REQ As String
    Public Shared IC_LV_COUNT As String
    Public Shared IC_LV_X As String
    Public Shared IC_LV_Y As String

    Public Shared IC_LVOP_REQ As String
    Public Shared IC_LVOP_COUNT As String
    Public Shared IC_LVOP_X As String
    Public Shared IC_LVOP_Y As String

    Public Shared IC_HV_TYPE As String
    Public Shared IC_HVOP_TYPE As String
    Public Shared IC_LV_TYPE As String
    Public Shared IC_LVOP_TYPE As String

    Public Shared top_tank_ladder_dis As String
    Public Shared top_tank_ladder_dis_LV As String
    Public Shared top_tank_ladder_dis_OLTC As String
    Public Shared top_tank_ladder_dis_OLTC_OPP As String
    Public Shared top_tank_ladder_ht As String
    Public Shared LADDER_MOUNTING As String

    Public Shared top_tank_ladder_HV_chk As String
    Public Shared top_tank_ladder_LV_chk As String
    Public Shared top_tank_ladder_OLTC_chk As String
    Public Shared top_tank_ladder_OLTCOPP_chk As String

    Public Shared rating_plate_HV As String
    Public Shared rating_plate_LV As String
    Public Shared rating_plate_OLTC As String
    Public Shared rating_plate_OLTC_OPP As String

    Public Shared rating_plate_HV_X As String
    Public Shared rating_plate_LV_X As String
    Public Shared rating_plate_OLTC_X As String
    Public Shared rating_plate_OLTC_OPP_X As String

    Public Shared rating_plate_HV_X2 As String
    Public Shared rating_plate_LV_X2 As String
    Public Shared rating_plate_OLTC_X2 As String
    Public Shared rating_plate_OLTC_OPP_X2 As String

    Public Shared rating_plate_HV_X3 As String
    Public Shared rating_plate_LV_X3 As String
    Public Shared rating_plate_OLTC_X3 As String
    Public Shared rating_plate_OLTC_OPP_X3 As String

    Public Shared rating_plate_HV_Y As String
    Public Shared rating_plate_LV_Y As String
    Public Shared rating_plate_OLTC_Y As String
    Public Shared rating_plate_OLTC_OPP_Y As String

    Public Shared rating_plate_HV_Y2 As String
    Public Shared rating_plate_LV_Y2 As String
    Public Shared rating_plate_OLTC_Y2 As String
    Public Shared rating_plate_OLTC_OPP_Y2 As String

    Public Shared rating_plate_HV_Y3 As String
    Public Shared rating_plate_LV_Y3 As String
    Public Shared rating_plate_OLTC_Y3 As String
    Public Shared rating_plate_OLTC_OPP_Y3 As String

    Public Shared OLTC_flat_stiffner As String
    Public Shared OLTC_FS1_OTY As String
    Public Shared OLTC_FS2_OTY As String
    Public Shared OLTC_FS3_OTY As String
    Public Shared OLTC_FS4_OTY As String
    Public Shared OLTC_FS1_WDTH As String
    Public Shared OLTC_FS2_WDTH As String
    Public Shared OLTC_FS3_WDTH As String
    Public Shared OLTC_FS4_WDTH As String
    Public Shared OLTC_FS1_THK As String
    Public Shared OLTC_FS2_THK As String
    Public Shared OLTC_FS3_THK As String
    Public Shared OLTC_FS4_THK As String
    Public Shared OLTC_DIM1 As String
    Public Shared OLTC_DIM2 As String
    Public Shared OLTC_DIM3 As String
    Public Shared OLTC_DIM4 As String
    Public Shared OLTC_DIM5 As String
    Public Shared OLTC_DIM6 As String

    'Stiffner2

    Public Shared no_flat_stiff_U As String
    Public Shared flat_stiff1_U As String
    Public Shared flat_stiff2_U As String
    Public Shared flat_stiff3_U As String
    Public Shared flat_stiff4_U As String
    Public Shared flat_stiff5_U As String
    Public Shared mid_insp_cvr_U As String
    Public Shared top_insp_cvr_U As String

    Public Shared no_flat_stiff_V As String
    Public Shared flat_stiff1_V As String
    Public Shared flat_stiff2_V As String
    Public Shared flat_stiff3_V As String
    Public Shared flat_stiff4_V As String
    Public Shared flat_stiff5_V As String
    Public Shared mid_insp_cvr_V As String
    Public Shared top_insp_cvr_V As String

    Public Shared LV_box_stiff As String
    Public Shared LV_box_stiff_typ As String
    Public Shared LV_box_stiff_ht As String
    Public Shared LV_box_stiff_wdt As String
    Public Shared LV_box_stiff_thk As String
    Public Shared slop_flat_stiff_thk As String
    Public Shared slop_flat_stiff_plate_thk As String

    Public Shared DIM1 As String
    Public Shared DIM2 As String
    Public Shared DIM3 As String
    Public Shared DIM4 As String

    Public Shared box_stiff_OLTC_END As String
    Public Shared box_stiff__typ_OLTC_END As String
    Public Shared box_stiff_ht_OLTC_END As String
    Public Shared box_stiff_wdt_OLTC_END As String
    Public Shared box_stiff_thk_OLTC_END As String
    Public Shared stiff_below_slope As String

    Public Shared DIM1_OLTC_END As String
    Public Shared DIM2_OLTC_END As String
    Public Shared DIM3_OLTC_END As String
    Public Shared DIM4_OLTC_END As String

    Public Shared flat_Stiff1_OLTC_SIDE As String
    Public Shared flat_Stiff2_OLTC_SIDE As String
    Public Shared flat_Stiff3_OLTC_SIDE As String
    Public Shared flat_Stiff4_OLTC_SIDE As String
    Public Shared box_stiff_thk_OLTC_SIDE As String

    Public Shared DIM1_OLTC_SIDE As String
    Public Shared DIM2_OLTC_SIDE As String
    Public Shared DIM3_OLTC_SIDE As String
    Public Shared DIM4_OLTC_SIDE As String

    'Top Cover
    Public Shared top_cover_types As String
    Public Shared HV_W_dist_X_axis As String
    Public Shared OLTC_flange_dist_X_axis As String
    Public Shared OLTC_flange_dist_Y_axis As String
    Public Shared HV_UVW_dist_Y_axis As String


    Public Shared TC_BSWV As String
    Public Shared TC_BSHV As String
    Public Shared TC_BSTV As String
    Public Shared TC_BSLV As String
    Public Shared TC_FSHV As String
    Public Shared TC_FSTV As String




    Public Shared dist1 As String
    Public Shared dist2 As String
    Public Shared dist3 As String
    Public Shared dist4 As String
    Public Shared dist5 As String

    Public Shared FS_2V As String
    Public Shared FS_3V As String
    Public Shared FS_4V As String
    Public Shared FS_5V As String
    Public Shared FS_6V As String

    Public Shared HV_turret_U_phase_cntr As String
    Public Shared HV_turret_V_phase_cntr As String
    Public Shared HV_turret_W_phase_cntr As String

    Public Shared CPF_REQ As String
    Public Shared CPF_COUNT As String
    Public Shared CPF_X1 As String
    Public Shared CPF_X2 As String
    Public Shared CPF_X3 As String
    Public Shared CPF_X4 As String
    Public Shared CPF_X5 As String
    Public Shared CPF_Y1 As String
    Public Shared CPF_Y2 As String
    Public Shared CPF_Y3 As String
    Public Shared CPF_Y4 As String
    Public Shared CPF_Y5 As String

    Public Shared BUCHHOLZ_RELAY_REQ As String
    Public Shared BUCHHOLZ_RELAY_X As String
    Public Shared BUCHHOLZ_RELAY_Y As String

    Public Shared NO_OF_PORT_TOP_HV As String
    Public Shared NO_OF_PORT_TOP_LV As String

    'Accessories1
    'PRD Mounting
    Public Shared PRD_mounting As String
    Public Shared PRD_x1 As String
    Public Shared PRD_x2 As String
    Public Shared PRD_x3 As String
    Public Shared PRD_Y1 As String
    Public Shared PRD_Y2 As String
    Public Shared PRD_Y3 As String

    'Buffer Locking
    Public Shared Buffer_Locking As String
    Public Shared Buffer_Locking_x1 As String
    Public Shared Buffer_Locking_x2 As String
    Public Shared Buffer_Locking_x3 As String
    Public Shared Buffer_Locking_y1 As String
    Public Shared Buffer_Locking_y2 As String
    Public Shared Buffer_Locking_y3 As String

    'Inspection cover for core earthing box
    Public Shared Inspection_cover As String
    Public Shared ins_cover_size As String
    Public Shared Inspection_cover_x1 As String
    Public Shared Inspection_cover_x2 As String
    Public Shared Inspection_cover_x3 As String
    Public Shared Inspection_cover_y1 As String
    Public Shared Inspection_cover_y2 As String
    Public Shared Inspection_cover_y3 As String

    'Safety Pole
    Public Shared Safety_pole As String
    Public Shared Safety_pole_x1 As String
    Public Shared Safety_pole_x2 As String
    Public Shared Safety_pole_x3 As String
    Public Shared Safety_pole_x4 As String
    Public Shared Safety_pole_y1 As String
    Public Shared Safety_pole_y2 As String
    Public Shared Safety_pole_y3 As String
    Public Shared Safety_pole_y4 As String

    'Other Accessories
    Public Shared CF_TE_X As String
    Public Shared CF_TE_Y As String
    Public Shared SPRR_MOUNTING_X As String
    Public Shared SPRR_MOUNTING_Y As String
    Public Shared GOR_CP_X As String
    Public Shared GOR_CP_Y As String

    'Marshallin box
    Public Shared Marshallin_box_Mounting As String
    Public Shared Marshallin_box_X1 As String
    Public Shared Marshallin_box_Y1 As String

    'Accessories2
    'Thermometer pocket
    Public Shared Thermometer_pocket As String
    Public Shared Thermometer_pocket_X1 As String
    Public Shared Thermometer_pocket_X2 As String
    Public Shared Thermometer_pocket_X3 As String
    Public Shared Thermometer_pocket_X4 As String
    Public Shared Thermometer_pocket_X5 As String
    Public Shared Thermometer_pocket_X6 As String
    Public Shared Thermometer_pocket_X7 As String
    Public Shared Thermometer_pocket_Y1 As String
    Public Shared Thermometer_pocket_Y2 As String
    Public Shared Thermometer_pocket_Y3 As String
    Public Shared Thermometer_pocket_Y4 As String
    Public Shared Thermometer_pocket_Y5 As String
    Public Shared Thermometer_pocket_Y6 As String
    Public Shared Thermometer_pocket_Y7 As String

    'Air Releasing Plug
    Public Shared air_release_plug As String
    Public Shared Air_release_plug_X1 As String
    Public Shared Air_release_plug_X2 As String
    Public Shared Air_release_plug_Y1 As String
    Public Shared Air_release_plug_Y2 As String

    'Vaccum Braking Value
    Public Shared Vaccum_braking_valve As String
    Public Shared Vaccum_braking_valve_X As String
    Public Shared Vaccum_braking_valve_Y As String

    'NIFPES Sensor Box
    Public Shared NIFPES_sensor_box As String
    Public Shared NIFPES_sensor_box_X As String
    Public Shared NIFPES_sensor_box_Y As String

    Public Shared quick_drain_val_1 As String
    Public Shared quick_drain_val_x As String
    Public Shared quick_drain_val_y As String

    'Impact recorder
    Public Shared Impact_recorder_req As String
    Public Shared Impact_recorder_qty As String
    Public Shared Impact_recorder_mounting As String
    Public Shared Impact_recorder_X As String
    Public Shared Impact_recorder_Y As String

    'OLTC Mounting
    Public Shared OLTC_MAKE As String
    Public Shared OLTC_MODEL As String
    Public Shared OLTC_MOUNTING_X As String
    Public Shared OLTC_MOUNTING_Y As String

    'Accessories3
    'ODS
    Public Shared ODS_REQ As String
    Public Shared ODS_MAKE As String
    Public Shared ODS_SIZE As String
    Public Shared ODS_POS As String
    Public Shared ODS_MOUNTING As String
    Public Shared ODS_X As String
    Public Shared ODS_Y As String

    'DGA
    Public Shared DGA_REQ As String
    Public Shared DGA_MAKE As String
    Public Shared DGA_SIZE As String
    Public Shared DGA_POS As String
    Public Shared DGA_MOUNTING As String
    Public Shared DGA_X As String
    Public Shared DGA_Y As String

    'Thermosyphon Filter
    Public Shared Thermosyphon_filter_Req As String
    Public Shared Thermosyphon_filter_Make As String
    Public Shared Thermosyphon_filter_Size As String
    Public Shared Thermosyphon_filter_POS As String
    Public Shared Thermosyphon_filter_Mounting As String
    Public Shared Thermosyphon_filter_X As String
    Public Shared Thermosyphon_filter_Y As String

    'UHF Sensor Valve
    Public Shared UHF_Sensor_Valve_REQ As String
    Public Shared UHF_Sensor_Valve_MAKE As String
    Public Shared UHF_Sensor_Valve_SIZE As String
    Public Shared UHF_Sensor_Valve_POS As String
    Public Shared UHF_QTY As String
    Public Shared UHF_Sensor_Valve_Mounting As String
    Public Shared UHF_Sensor_Valve_X As String
    Public Shared UHF_Sensor_Valve_Y As String

    'Moisture Sensor
    Public Shared Moisture_sensor_REQ As String
    Public Shared Moisture_sensor_MAKE As String
    Public Shared Moisture_sensor_SIZE As String
    Public Shared Moisture_sensor_POS As String
    Public Shared Moisture_sensor_Mounting As String
    Public Shared Moisture_sensor_X As String
    Public Shared Moisture_sensor_Y As String

    'FOS
    Public Shared FOS_REQ As String
    Public Shared FOS_MAKE As String
    Public Shared FOS_Channel_Mounting As String
    Public Shared FOS_Box_Mounting As String
    Public Shared FOS_Inspection_cover As String
    Public Shared FOS_X1 As String
    Public Shared FOS_X2 As String
    Public Shared FOS_Y1 As String
    Public Shared FOS_Y2 As String
    Public Shared FOS_MTG_REQ As String
    Public Shared FOS_MTG_POS As String
    Public Shared FOS_MTG_X As String
    Public Shared FOS_MTG_Y As String

    'N2 Cylinder
    Public Shared N2_cylinder_REQ As String
    Public Shared N2_cylinder_QTY As String
    Public Shared N2_cylinder_MP As String
    Public Shared N2_cylinder_AR As String


    Public Shared ODS_X1 As String
    Public Shared ODS_Y1 As String
    Public Shared ODS_Y2 As String

    Public Shared DGA_X1 As String
    Public Shared DGA_Y1 As String
    Public Shared DGA_Y2 As String

    'UHF
    Public Shared UHF_Y1 As String
    Public Shared UHF_X1 As String
    Public Shared UHF_X2 As String
    Public Shared UHF_X3 As String

    Public Shared N2_X As String
    Public Shared N2_Y As String

    Public Shared N2_X1 As String
    Public Shared N2_Y1 As String


    'MS
    Public Shared Sensor_X As String
    Public Shared Sensor_Y As String

    'Accesories5
    Public Shared MPS_REQ As String
    Public Shared MPS_POS As String
    Public Shared MPS_X As String
    Public Shared MPS_Y As String

    Public Shared LBA_REQ As String
    Public Shared LBA_HV_COUNT As String
    Public Shared LBA_HVX1 As String
    Public Shared LBA_HVX2 As String
    Public Shared LBA_HVY1 As String
    Public Shared LBA_HVY2 As String
    Public Shared LBA_LV_COUNT As String
    Public Shared LBA_HV_X1 As String
    Public Shared LBA_HV_X2 As String
    Public Shared LBA_HV_Y1 As String
    Public Shared LBA_HV_Y2 As String

    Public Shared TOLG_REQ As String
    Public Shared TOLG_POS As String
    Public Shared TOLG_X As String
    Public Shared TOLG_Y As String

    Public Shared GCDM_REQ As String
    Public Shared GCDM_POS As String
    Public Shared GCDM_X As String
    Public Shared GCDM_Y As String

    Public Shared TFV_REQ As String
    Public Shared TFV_POS As String
    Public Shared TFV_X As String
    Public Shared TFV_Y As String

    Public Shared TSV_REQ As String
    Public Shared TSV_POS As String
    Public Shared TSV_X As String
    Public Shared TSV_Y As String

    Public Shared TQDV_REQ As String
    Public Shared TQDV_POS As String
    Public Shared TQDV_X As String
    Public Shared TQDV_Y As String

    Public Shared LBA_TYPE As String

    'Cooler Bank

    'Cooler Bank Details
    Public Shared cooler_bnk As String
    'Public Shared CB_MOUNT As String
    Public Shared cooler_bnk_pos1 As String
    Public Shared cooler_bnk_pos2 As String
    Public Shared cooler_bnk_pos3 As String
    Public Shared typ_of_cooling As String
    Public Shared cooler_pipe_siz As String
    Public Shared Cooling_Fans As String
    Public Shared Cooling_size As String
    Public Shared Bank1_main As String
    Public Shared Bank2_main As String
    Public Shared Bank3_main As String
    Public Shared Bank1_standby As String
    Public Shared Bank2_standby As String
    Public Shared Bank3_standby As String
    Public Shared Radiator_size_C_C As String
    Public Shared Radiator_size_W As String
    Public Shared Radiator_size_FINS As String
    Public Shared Radiator_size_QTY As String
    Public Shared Radiator_pipe_size As String
    Public Shared Ground_Support_Reg As String
    Public Shared Radiator_drain1 As String
    Public Shared Radiator_drain2 As String
    Public Shared Qty_Bank1 As String
    Public Shared Qty_Bank2 As String
    Public Shared Qty_Bank3 As String

    'Cooler Pipe & Header Details
    'Top Cooler pipe
    Public Shared cooler_pipe_size As String
    Public Shared CB_MOUNT As String
    Public Shared no_of_port As String
    Public Shared butterfly_valve_size As String
    Public Shared expansion_bellows As String
    Public Shared expansion_bellows_size As String
    Public Shared TCP_HV_X1 As String
    Public Shared TCP_HV_X2 As String
    Public Shared TCP_HV_X3 As String
    Public Shared TCP_LV_X1 As String
    Public Shared TCP_LV_X2 As String
    Public Shared TCP_LV_X3 As String
    Public Shared TCP_HV_Y1 As String
    Public Shared TCP_HV_Y2 As String
    Public Shared TCP_HV_Y3 As String
    Public Shared TCP_LV_Y1 As String
    Public Shared TCP_LV_Y2 As String
    Public Shared TCP_LV_Y3 As String
    Public Shared header_pipe_size_bot As String
    Public Shared NO_SUPV As String
    Public Shared TCP_LV_Y6 As String
    Public Shared TCP_LV_Y4 As String
    Public Shared TCP_LV_Y5 As String

    Public Shared Z_1 As String
    Public Shared Z_2 As String
    Public Shared Z_3 As String
    Public Shared Z_4 As String

    Public Shared V_1 As String
    Public Shared H_U As String
    Public Shared H_V As String
    Public Shared RCC As String


    'CB2
    Public Shared B2_NO_P As String
    Public Shared B2_X1 As String
    Public Shared B2_Y1 As String
    Public Shared B2_X2 As String
    Public Shared B2_Y2 As String
    Public Shared B2_X3 As String
    Public Shared B2_Y3 As String


    Public Shared B2_Z1 As String
    Public Shared B2_Z2 As String
    Public Shared B2_Z3 As String
    Public Shared B2_Z4 As String

    Public Shared B2_BX1 As String
    Public Shared B2_BX2 As String
    Public Shared B2_BY As String

    Public Shared B2_D1 As String
    Public Shared B2_D2 As String
    Public Shared B2_D3 As String
    Public Shared B2_D4 As String
    Public Shared B2_Y As String

    Public Shared B2_V1 As String
    Public Shared B2_U As String
    Public Shared B2_V As String
    Public Shared B2_R_C_C As String




    Public Shared B3_NO_P As String
    Public Shared B3_NO_GSUP As String
    Public Shared B2_NO_GSUP As String



    Public Shared B3_X1 As String
    Public Shared B3_X2 As String
    Public Shared B3_X3 As String
    Public Shared B3_Y1 As String
    Public Shared B3_Y2 As String
    Public Shared B3_Y3 As String


    Public Shared B3_BX1 As String
    Public Shared B3_BX2 As String
    Public Shared B3_BY As String

    Public Shared B3_D1 As String
    Public Shared B3_D2 As String
    Public Shared B3_D3 As String
    Public Shared B3_D4 As String
    Public Shared B3_Y As String

    Public Shared B3_V1 As String
    Public Shared B3_U As String
    Public Shared B3_V As String
    Public Shared B3_R_C_C As String


    Public Shared B3_Z1 As String
    Public Shared B3_Z2 As String
    Public Shared B3_Z3 As String
    Public Shared B3_Z4 As String






    'ONAN
    Public Shared ONAN_no_port As String
    Public Shared typ_of_cooling_cb1 As String
    Public Shared Cooler_Pipe_size_ONAN As String
    Public Shared ONAN_butterfly_valve_size As String
    Public Shared ONAN_expansion_bellows As String
    Public Shared ONAN_expansion_bellows_size As String
    Public Shared ONAN_cooler_pipe_supp_qty As String

    'ONAN/ONAF
    Public Shared bot_Pipe_size_ONAF As String
    Public Shared TYP_COOLING_ONAF As String
    Public Shared NO_OF_PORT_ONAF As String
    Public Shared BF_VL_ONAF As String
    Public Shared EXP_BELOW_ONAF As String
    Public Shared EXP_BLW_SIZE_ONAF As String
    Public Shared IFRAME_ONAF As String

    Public Shared P1L_LV As String
    Public Shared L1_LV As String
    Public Shared P2L_LV As String
    Public Shared P3L_LV As String
    Public Shared L2_LV As String
    Public Shared P4L_LV As String
    Public Shared L4_LV As String
    Public Shared P5L1_LV As String

    Public Shared P1L_HV As String
    Public Shared L1_HV As String
    Public Shared P2L_HV As String
    Public Shared P3L_HV As String
    Public Shared L2_HV As String
    Public Shared P4L_HV As String
    Public Shared L4_HV As String
    Public Shared P5L1_HV As String

    Public Shared BL1_LV As String
    Public Shared BL1_HV As String

    Public Shared X1LV As String
    Public Shared X2LV As String
    Public Shared X3LV As String
    Public Shared Y1LV As String
    Public Shared Y2LV As String

    Public Shared X1HV As String
    Public Shared X2HV As String
    Public Shared X3HV As String
    Public Shared Y1HV As String
    Public Shared Y2HV As String

    Public Shared ID1 As String
    Public Shared ID2 As String

    Public Shared BC_H1 As String
    Public Shared BC_H2 As String
    Public Shared BC_H3 As String

    'ONAF




    Public Shared ODAF_cooler_pipe_size As String
    Public Shared ODAF_no_of_port As String
    Public Shared ODAF_no_of_port1 As String
    Public Shared ODAF_butterfly_valve_size As String
    Public Shared ODAF_expansion_bellows As String
    Public Shared ODAF_expansion_bellows_size As String

    Public Shared P1L_LVF As String
    Public Shared L1_LVF As String
    Public Shared P2L_LVF As String
    Public Shared P3L_LVF As String
    Public Shared L2_LVF As String
    Public Shared P4L_LVF As String
    Public Shared L4_LVF As String
    Public Shared P5L1_LVF As String

    Public Shared P1L_HVF As String
    Public Shared L1_HVF As String
    Public Shared P2L_HVF As String
    Public Shared P3L_HVF As String
    Public Shared L2_HVF As String
    Public Shared P4L_HVF As String
    Public Shared L4_HVF As String
    Public Shared P5L1_HVF As String

    Public Shared BL1_LVF As String
    Public Shared BL1_HVF As String

    Public Shared X1LVF As String
    Public Shared X2LVF As String
    Public Shared X3LVF As String
    Public Shared Y1LVF As String
    Public Shared Y2LVF As String

    Public Shared X1HVF As String
    Public Shared X2HVF As String
    Public Shared X3HVF As String
    Public Shared Y1HVF As String
    Public Shared Y2HVF As String

    Public Shared ID1F As String
    Public Shared ID2F As String

    Public Shared BC_H1F As String
    Public Shared BC_H2F As String
    Public Shared BC_H3F As String




    'TOP COOLER PIPE

    Public Shared P1L_LVT As String
    Public Shared P2L_LVT As String
    Public Shared P3L_LVT As String
    Public Shared P4L_LVT As String
    Public Shared P5L_LVT As String

    Public Shared P1L_HVT As String
    Public Shared P2L_HVT As String
    Public Shared P3L_HVT As String
    Public Shared P4L_HVT As String
    Public Shared P5L_HVT As String


    Public Shared X1LVT As String
    Public Shared X2LVT As String
    Public Shared X3LVT As String
    Public Shared Y1LVT As String

    Public Shared X1HVT As String
    Public Shared X2HVT As String
    Public Shared X3HVT As String
    Public Shared Y1HVT As String

    Public Shared L1_FVT As String
    Public Shared H1_FVT As String
    Public Shared H2_FVT As String
    Public Shared HL1_FVT As String

    'AFRAME ARRANGEMENT

    Public Shared X100 As String
    Public Shared H1100 As String
    Public Shared U1100 As String
    Public Shared U2100 As String
    Public Shared UT100 As String
    Public Shared D1100 As String
    Public Shared D2100 As String


    'Top & Bottom Header

    Public Shared HP_SIZE_HV As String
    Public Shared RAD_CC_HV As String
    Public Shared RAD_P_SIZE_HV As String
    Public Shared N0_RAD_HV As String

    Public Shared AIR_R_PG_HV As String
    Public Shared AIR_PLUG_HV As String

    Public Shared T_P_HV As String
    Public Shared TM_POC_VAL_HV As String
    Public Shared FCD_HV As String
    Public Shared FCD_VAL_HV As String
    Public Shared CP_LEN_HV As String
    Public Shared HP_LEN_HV As String


    Public Shared HP_SIZE_LV As String
    Public Shared RAD_CC_LV As String
    Public Shared RAD_P_SIZE_LV As String
    Public Shared N0_RAD_LV As String

    Public Shared AIR_R_PG_LV As String
    Public Shared AIR_PLUG_LV As String

    Public Shared T_P_LV As String
    Public Shared TM_POC_VAL_LV As String
    Public Shared FCD_LV As String
    Public Shared FCD_VAL_LV As String
    Public Shared CP_LEN_LV As String
    Public Shared HP_LEN_LV As String



    Public Shared HP_SIZE_OOS As String
    Public Shared RAD_CC_OOS As String
    Public Shared RAD_P_SIZE_OOS As String
    Public Shared N0_RAD_OOS As String

    Public Shared AIR_R_PG_OOS As String
    Public Shared AIR_PLUG_OOS As String

    Public Shared T_P_OOS As String
    Public Shared TM_POC_VAL_OOS As String
    Public Shared FCD_OOS As String
    Public Shared FCD_VAL_OOS As String
    Public Shared CP_LEN_OOS As String
    Public Shared HP_LEN_OOS As String


    Public Shared HP_SIZE_OS As String
    Public Shared RAD_CC_OS As String
    Public Shared RAD_P_SIZE_OS As String
    Public Shared N0_RAD_OS As String

    Public Shared AIR_R_PG_OS As String
    Public Shared AIR_PLUG_OS As String

    Public Shared T_P_OS As String
    Public Shared TM_POC_VAL_OS As String
    Public Shared FCD_OS As String
    Public Shared FCD_VAL_OS As String
    Public Shared CP_LEN_OS As String
    Public Shared HP_LEN_OS As String



    'A Frame Proposal
    'A FRAME
    Public Shared H1 As String
    Public Shared H2 As String
    Public Shared H3 As String
    Public Shared H4 As String
    Public Shared H5 As String
    Public Shared H6 As String
    Public Shared L1 As String

    Public Shared ANG As String

    Public Shared A1 As String

    Public Shared A As String

    Public Shared D As String

    Public Shared DG As String

    Public Shared DL As String
    Public Shared A_D1 As String
    Public Shared A_D2 As String

    Public Shared ISMC_W As String
    Public Shared ISMC_H As String
    Public Shared ISMC_T As String
    Public Shared ISMC_L As String





    Public Shared ABX1 As String
    Public Shared ABX2 As String
    Public Shared ABX3 As String
    Public Shared A_L1 As String

    Public Shared ATY1 As String
    Public Shared ATY2 As String
    Public Shared APL1 As String
    Public Shared APL2 As String
    Public Shared APL3 As String
    Public Shared APL4 As String
    Public Shared APL5 As String
    Public Shared APL6 As String
    Public Shared ABY1 As String
    Public Shared ABY2 As String
    Public Shared APY1 As String
    Public Shared APY2 As String
    Public Shared D1A As String
    Public Shared D2A As String
    Public Shared D3A As String
    Public Shared D4A As String
    Public Shared D5A As String
    Public Shared D6A As String
    Public Shared D7A As String
    Public Shared D8A As String

    Public Shared bk1_main As String
    Public Shared bk2_main As String
    Public Shared bk3_main As String
    Public Shared Coolingfan_2x50 As String
    Public Shared Coolingfansize_2x50 As String

    Public Shared b1_SBY As String
    Public Shared b2_SBY As String
    Public Shared b3_SBY As String

    'LADDER
    Public Shared CB_ladder_HV_chk As String
    Public Shared CB_ladder_LV_chk As String

    'Conservator & it's support
    Public Shared Conservator_type1 As String
    Public Shared Conservator_type2 As String
    Public Shared Conservator_type3 As String
    Public Shared Conservator_dia As String
    Public Shared Conservator_len As String
    Public Shared Conservator_Pos1 As String
    Public Shared Conservator_Pos2 As String
    Public Shared Conservator_Bracket As String
    Public Shared Conservator_supp_dis_B1 As String
    Public Shared Conservator_supp_dis_B2 As String
    Public Shared LB_R_pipe_centeradder_C As String
    Public Shared Conservator_drain_valve As String
    Public Shared Conservator_mounting As String
    Public Shared Aircel_rapture_relay As String
    Public Shared MOLG_SIZE As String
    Public Shared IITL_Monogram_plate As String
    Public Shared B_R_Pipe_Connection As String
    Public Shared Breather_pipe_connection As String
    Public Shared Air_release_pipe_connection As String

    'OLTC Conservator
    Public Shared OLTC_Conservator_type1 As String
    Public Shared OLTC_Conservator_type2 As String
    Public Shared OLTC_conservator_dia As String
    Public Shared OLTC_conservator_pos As String
    Public Shared OLTC_cons_position As String
    Public Shared OLTC_Conservator_supp As String
    Public Shared OLTC_Conservator_supp_dist_X As String
    Public Shared OLTC_Conservator_supp_dist_Y As String
    Public Shared OSR_pipe_ceneter As String
    Public Shared OLTC_Conservator_drain_valve As String
    Public Shared OLTC_MOLG_size As String
    Public Shared OLTC_Breather_pipe_connection As String
    Public Shared OLTC_suction_pipe_connection As String
    Public Shared Conservator_breather As String
    Public Shared Conservator_breather_type As String

    Public Shared H12X100 As String
    Public Shared U12X100 As String
    Public Shared U22X100 As String
    Public Shared U32X100 As String













    Friend Sub InventorConnect()
        Try
            ThisApplication = GetObject(, "Inventor.Application")
        Catch ex As Exception
            MsgBox("Unable to connect to Inventor.")
        End Try
    End Sub

    Private Sub dashBtn_Click(sender As Object, e As EventArgs) Handles dashBtn.Click

        Try
            Me.Hide()
            Create_New_User.ShowDialog()
        Catch ex As Exception
            MsgBox("ERR : " & ex.ToString)
        End Try

        'user = userTxt.Text
        'Dim secondForm As New Form2()

        '' Set the 'CreatedBy' property to the entered username
        'secondForm.CreatedBy = user

        '' Show Form2
        'secondForm.Show()
        'Dim createUserForm As New Form2()
        'createUserForm.LoggedInUser = LoggedInUser
        'MessageBox.Show("LoggedInUser set to: " & LoggedInUser)
        'createUserForm.ShowDialog()
        'Me.Show()

    End Sub

    Private Sub user_btn_Click(sender As Object, e As EventArgs) Handles user_btn.Click

        Try

            If userTxt.Text = Nothing Or passTxt.Text = Nothing Then
                MsgBox("Enter username and password to proceed")
            Else
                'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
                'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
                Dim connection As New SqlConnection(connectionString)

                Try




                    connection.Open()

                    Dim username As String = userTxt.Text
                    Dim password As String = passTxt.Text
                    Dim role As String = GetUserRole(username, password)
                    If role <> "Admin" Then
                        dashBtn.Enabled = False
                    End If




                    Dim query As String = "SELECT * FROM log WHERE Username=@username AND Password=@password"
                    Dim cmd As New SqlCommand(query, connection)
                    cmd.Parameters.AddWithValue("@username", userTxt.Text)
                    cmd.Parameters.AddWithValue("@password", passTxt.Text)

                    Dim reader As SqlDataReader = cmd.ExecuteReader()
                    Dim cred As String
                    If reader.HasRows Then
                        reader.Read()
                        cred = reader("Credential").ToString()
                        usertype_label.Text = cred

                        If cred = "Admin" Then
                            dashBtn.Enabled = True
                        ElseIf cred = "User" Then
                            dashBtn.Enabled = False
                        End If

                        If reader.HasRows Then
                            edit_prj_btn.Enabled = True
                            crt_prj_btn.Enabled = True
                        Else
                            edit_prj_btn.Enabled = False
                            crt_prj_btn.Enabled = False
                        End If
                    Else
                        MsgBox("Incorrect username or password")
                    End If

                    reader.Close()
                Catch ex As Exception
                    MessageBox.Show("Data Base Is Not connected, kindly check the Database connection")
                Finally
                    connection.Close()
                End Try
            End If

        Catch ex As Exception
            MsgBox("Data Base is not connected,kindly check the database connection")
        End Try
    End Sub

    Private Sub crt_prj_btn_Click(sender As Object, e As EventArgs) Handles crt_prj_btn.Click

        Me.Hide()
        Create_Project.ShowDialog()

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles edit_prj_btn.Click

        Me.Hide()
        Edit_Project.ShowDialog()

    End Sub

    'Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    '    usertype_label.Text = String.Empty

    '    Try

    '        Using conn As New SqlConnection(connectionString)

    '            MessageBox.Show(
    '            "Trying SQL Server:" &
    '            System.Environment.NewLine &
    '            conn.DataSource &
    '            System.Environment.NewLine &
    '            System.Environment.NewLine &
    '            "Database:" &
    '            System.Environment.NewLine &
    '            conn.Database,
    '            "SQL Connection Debug",
    '            MessageBoxButtons.OK,
    '            MessageBoxIcon.Information
    '        )

    '            conn.Open()

    '            MessageBox.Show(
    '            "Database connected successfully." &
    '            System.Environment.NewLine &
    '            System.Environment.NewLine &
    '            "Server: " & conn.DataSource &
    '            System.Environment.NewLine &
    '            "Database: " & conn.Database,
    '            "Database Connection",
    '            MessageBoxButtons.OK,
    '            MessageBoxIcon.Information
    '        )

    '        End Using

    '    Catch ex As SqlException

    '        MessageBox.Show(
    '        "Database connection failed." &
    '        System.Environment.NewLine &
    '        System.Environment.NewLine &
    '        "SQL Error Number: " & ex.Number.ToString() &
    '        System.Environment.NewLine &
    '        System.Environment.NewLine &
    '        "Message:" &
    '        System.Environment.NewLine &
    '        ex.Message,
    '        "Database Connection Error",
    '        MessageBoxButtons.OK,
    '        MessageBoxIcon.Error
    '    )

    '        System.Windows.Forms.Application.Exit()

    '    Catch ex As Exception

    '        MessageBox.Show(
    '        "Database connection failed." &
    '        System.Environment.NewLine &
    '        System.Environment.NewLine &
    '        ex.Message,
    '        "Database Connection Error",
    '        MessageBoxButtons.OK,
    '        MessageBoxIcon.Error
    '    )

    '        System.Windows.Forms.Application.Exit()

    '    End Try

    'End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        usertype_label.Text = String.Empty

        Try

            Using conn As New SqlConnection(connectionString)
                conn.Open()
            End Using

        Catch ex As Exception
            MessageBox.Show("Database is not connected,Please check the connection and try again!!", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            System.Environment.Exit(0)
        End Try

    End Sub


    Private Function GetUserRole(username As String, password As String) As String

        'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        Dim connection As New SqlConnection(connectionString)
        Try
            connection.Open()
            Dim query As String = "SELECT Credential FROM log WHERE Username=@username and Password=@password"
            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@username", username)
            cmd.Parameters.AddWithValue("@password", password)
            Dim role As String = Convert.ToString(cmd.ExecuteReader())
            Return role
        Catch ex As Exception
            MessageBox.Show("Database error: " & ex.Message)
            Return String.Empty
        Finally
            connection.Close()
        End Try
    End Function

    Private Sub passTxt_TextChanged(sender As Object, e As EventArgs) Handles passTxt.TextChanged
        'Dim username As String = userTxt.Text
        'Dim password As String = passTxt.Text
        'Dim role As String = GetUserRole(username, password)
        'If role <> "Admin" Then
        '    dashBtn.Enabled = False
        'End If
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        user = userTxt.Text

    End Sub

    Friend Shared Function OLTC_pkt_pos() As Object
        Throw New NotImplementedException()
    End Function


    Private Sub GetPlanes(ByVal Occurrence As Inventor.ComponentOccurrence,
                     ByRef BaseYZ As Inventor.WorkPlane,
                     ByRef BaseXZ As Inventor.WorkPlane,
                     ByRef BaseXY As Inventor.WorkPlane)
        ' Get the work planes from the definition of the occurrence.
        ' These will be in the context of the part or subassembly, not
        ' the top-level assembly, which is what we need to return.
        BaseYZ = Occurrence.Definition.WorkPlanes.Item(1)
        BaseXZ = Occurrence.Definition.WorkPlanes.Item(2)
        BaseXY = Occurrence.Definition.WorkPlanes.Item(3)

        ' Create proxies for these planes.  This will act as the work
        ' plane in the context of the top-level assembly.
        Occurrence.CreateGeometryProxy(BaseYZ, BaseYZ)
        Occurrence.CreateGeometryProxy(BaseXZ, BaseXZ)
        Occurrence.CreateGeometryProxy(BaseXY, BaseXY)
    End Sub


End Class

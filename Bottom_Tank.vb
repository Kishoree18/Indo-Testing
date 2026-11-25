Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Security.Policy
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Inventor

Public Class Bottom_Tank

    Public Shared ThisApplication As Inventor.Application
    Public Shared MainMaster As Inventor.Document
    Private ReadOnly TabContro5 As Object
    Public Shared Parameters As New Dictionary(Of String, String)
    Public Shared OLTC_pck_len1
    Public Shared OLTC_pck_wdh1

    Private OLTC_pkt_position As String = ""

    Private Sub Load_OLTC_Pkt_Position()

        Try

            Dim wk_odr_No1 As String =
                Create_Project.wk_ord_No

            Using connection As New SqlConnection(
                Main_Panel.connectionString
            )

                connection.Open()

                Dim query As String =
                    "SELECT OLTC_pkt_position " &
                    "FROM Basic_Input " &
                    "WHERE wk_odr_No = @wk_odr_No"

                Using cmd As New SqlCommand(
                    query,
                    connection
                )

                    cmd.Parameters.AddWithValue(
                        "@wk_odr_No",
                        wk_odr_No1
                    )

                    Dim result As Object =
                        cmd.ExecuteScalar()

                    If result IsNot Nothing AndAlso
                       result IsNot DBNull.Value Then

                        OLTC_pkt_position =
                            result.ToString().Trim()

                    Else

                        OLTC_pkt_position =
                            ""

                    End If

                End Using

            End Using

        Catch ex As Exception

            MsgBox(
                "Error while getting OLTC packet position:" &
                vbCrLf &
                ex.ToString()
            )

        End Try

    End Sub


    Public Sub save_btm_tank_inputs()

        'botom tank

        Main_Panel.btm_tk_typeV = btm_tk_type.SelectedItem
        Main_Panel.btm_tk_curb_proV = btm_tk_curb_pro.SelectedItem
        Main_Panel.curb_sizev = curb_size.SelectedItem

        Main_Panel.btm_plt_htV = btm_plt_ht.SelectedItem
        Main_Panel.btm_tk_i_lenV = btm_tk_i_len.Text
        Main_Panel.btm_plt_lenV = btm_plt_len.Text
        Main_Panel.btm_plt_thkV = btm_plt_thk.Text
        Main_Panel.btm_tk_i_wdhV = btm_tk_i_wdh.Text
        Main_Panel.btm_plt_widthV = btm_plt_width.Text

        Main_Panel.Cus_htV = Cus_ht.Text
        Main_Panel.Side_Plt_thkV = Side_Plt_thk.Text
        Main_Panel.OLTC_pck_lenV = OLTC_pck_len.Text
        Main_Panel.OLTC_pck_wdhV = OLTC_pck_wdh.Text
        Main_Panel.tk_curb_wdhV = tk_curb_wdh.Text
        Main_Panel.Curb_wdh_LVV = Curb_wdh_LV.Text
        Main_Panel.Curb_wdh_DV = Curb_wdh_D.Text
        Main_Panel.Curb_OLTC_EndV = Curb_OLTC_End.Text

        Main_Panel.GPX1V = GPX1.Text
        Main_Panel.GPX2V = GPX2.Text
        Main_Panel.GPY1V = GPY1.Text
        Main_Panel.GPY2V = GPY2.Text

        Main_Panel.btm_len_chkbxV = btm_len_chkbx.Checked
        Main_Panel.btm_plt_wdt_ck_bxV = btm_plt_wdt_ck_bx.Checked


        'underbase
        Main_Panel.No_of_RailsV = No_of_Rails.SelectedItem
        Main_Panel.No_of_Rails1V = No_of_Rails1.SelectedItem
        Main_Panel.CC_LENGTHV = CC_LENGTH.SelectedItem
        Main_Panel.Roller_CustomV = Roller_Custom.Text
        Main_Panel.roller_gauge_cmbV = roller_gauge_cmb.SelectedItem
        Main_Panel.roller_orientationV = roller_orientation.SelectedItem

        Main_Panel.XV = X.Text

        Main_Panel.BS_RV = BS_R.Text
        Main_Panel.BS_LV = BS_L.Text

        Main_Panel.MP_lenV = MP_len.Text
        Main_Panel.MP_WthV = MP_Wth.Text
        Main_Panel.MP_ThkV = MP_Thk.Text
        Main_Panel.Dim_L1V = Dim_L1.Text
        Main_Panel.T2V = T2.Text
        'Main_Panel.Dim_T1V = Dim_T1.Text

        Main_Panel.BS1_WDH_1V = BS1_WDH_1.Text
        Main_Panel.BS2_WDH_1V = BS2_WDH_1.Text
        Main_Panel.BS3_WDH_1V = BS3_WDH_1.Text

        Main_Panel.BS1_HT_WSV = BS1_HT_WS.Text
        Main_Panel.BS2_HT_WSV = BS2_HT_WS.Text
        Main_Panel.BS3_HT_WSV = BS3_HT_WS.Text

        Main_Panel.BS1_THK_1V = BS1_THK_1.Text
        Main_Panel.BS2_THK_1V = BS2_THK_1.Text
        Main_Panel.BS3_THK_1V = BS3_THK_1.Text

        Main_Panel.prob_support_cmbV = prob_support_cmb.SelectedItem
        Main_Panel.prob_noV = prob_no.SelectedItem
        Main_Panel.PROB_HTV = PROB_HT.Text
        Main_Panel.Prob_X1axisV = Prob_X1axis.Text
        Main_Panel.Prob_X2axisV = Prob_X2axis.Text
        Main_Panel.Prob_Y1axisV = Prob_Y1axis.Text
        Main_Panel.Prob_Y2axisV = Prob_Y2axis.Text

        Main_Panel.HS_CC_lenV = HS_CC_len.Text
        Main_Panel.Distance_AV = Distance_A.Text
        Main_Panel.Distance_BV = Distance_A.Text



        'FLAT STIFFNER
        Main_Panel.HV_STNR_TYPEV = HV_STNR_TYPE.SelectedItem
        Main_Panel.HV_FS_AV = HV_FS_A.Text
        Main_Panel.HV_FS_BV = HV_FS_B.Text
        Main_Panel.HV_FS_A1V = HV_FS_A1.Text
        Main_Panel.HV_FS_B1V = HV_FS_B1.Text

        'LV
        Main_Panel.LV_STNR_TYPEV = LV_STNR_TYPE.SelectedItem
        Main_Panel.LV_FS_AV = LV_FS_A.Text
        Main_Panel.LV_FS_BV = LV_FS_B.Text
        Main_Panel.LV_FS_A1V = LV_FS_A1.Text
        Main_Panel.LV_FS_B1V = LV_FS_B1.Text




        'OLTC S1
        Main_Panel.OLTC_STNR_TYPEV = OLTC_STNR_TYPE.Text
        Main_Panel.OLTC_S1_AV = OLTC_S1_A.Text
        Main_Panel.OLTC_S1_BV = OLTC_S1_B.Text
        Main_Panel.OLTC_S1_A1V = OLTC_S1_A1.Text
        Main_Panel.OLTC_S1_B1V = OLTC_S1_B1.Text
        Main_Panel.OLTC_S1_TV = OLTC_S1_T.Text

        Main_Panel.OS1D1V = OS1D1.Text
        Main_Panel.OS1D2V = OS2D2.Text

        'OLTC S2
        Main_Panel.OLTC_STNR2_TYPEV = OLTC_STNR2_TYPE.Text
        Main_Panel.OLTC_S2_AV = OLTC_S2_A.Text
        Main_Panel.OLTC_S2_BV = OLTC_S2_B.Text
        Main_Panel.OLTC_S2_A1V = OLTC_S2_A1.Text

        Main_Panel.OS2D1V = OS2D1.Text
        Main_Panel.OS2D2V = OS2D2.Text


        'OLTC END
        Main_Panel.OLTC_END_STNR_TYPV = OLTC_END_STNR_TYP.Text
        Main_Panel.OLTC_END_AV = OLTC_END_A.Text
        Main_Panel.OLTC_END_BV = OLTC_END_B.Text
        Main_Panel.OLTC_END_A1V = OLTC_END_A1.Text

        Main_Panel.OED1V = OED1.Text
        Main_Panel.OED2V = OED2.Text
        Main_Panel.OED3V = OED3.Text
        Main_Panel.OED4V = OED4.Text

        Main_Panel.Stiffner_count_BT = Stiffner_count_BT.SelectedItem
        Main_Panel.oltc_stiffner1_count = oltc_stiffner1_count.SelectedItem
        Main_Panel.oltc_stiffner2_count = oltc_stiffner2_count.SelectedItem

        Main_Panel.stiffner_count_HV = stiffner_count_HV.SelectedItem
        Main_Panel.HVD1 = HVD1.Text
        Main_Panel.HVD2 = HVD2.Text
        Main_Panel.HVD3 = HVD3.Text
        Main_Panel.HVD4 = HVD4.Text

        Main_Panel.stiffner_count_LV = stiffner_count_LV.SelectedItem
        Main_Panel.LVD1 = LVD1.Text
        Main_Panel.LVD2 = LVD2.Text
        Main_Panel.LVD3 = LVD3.Text
        Main_Panel.LVD4 = LVD4.Text

        'jackingpad
        'hv
        Main_Panel.jack_pad_typeV = jack_pad_type.SelectedItem
        Main_Panel.JP_AV = JP_A.Text
        Main_Panel.JP_BV = JP_B.Text
        Main_Panel.JP_CV = JP_C.Text
        Main_Panel.JP_DV = JP_D.Text
        Main_Panel.JP_EV = JP_E.Text
        Main_Panel.JP_GV = JP_G.Text

        Main_Panel.JP_TV = JP_T.Text
        Main_Panel.JP_B1V = JP_B1.Text
        Main_Panel.JP_BB1V = JP_BB1.Text
        Main_Panel.JP_B2V = JP_B2.Text

        Main_Panel.NO_TOP_JP_HVV = NO_TOP_JP_HV.SelectedItem
        Main_Panel.JP_T_LV = JP_T_L.Text
        Main_Panel.JP_T_HV = JP_T_H.Text
        Main_Panel.JP_T_TV = JP_T_T.Text

        Main_Panel.NO_OF_JP_HVV = NO_OF_JP_HV.SelectedItem
        Main_Panel.HV_JPD1V = HV_JPD1.Text
        Main_Panel.HV_JPD2V = HV_JPD2.Text
        Main_Panel.HV_JPD3V = HV_JPD3.Text
        Main_Panel.HV_JPD4V = HV_JPD4.Text
        Main_Panel.JP_Y_DISTV = JP_Y_DIST.Text

        'LV
        Main_Panel.jack_pad_type1 = jack_pad_type1.SelectedItem
        Main_Panel.JP_LV_AV = JP_LV_A.Text
        Main_Panel.JP_LV_BV = JP_LV_B.Text
        Main_Panel.JP_LV_CV = JP_LV_C.Text
        Main_Panel.JP_LV_DV = JP_LV_D.Text
        Main_Panel.JP_LV_EV = JP_LV_E.Text
        'Main_Panel.JP_LV_FV = JP_LV_F.Text
        Main_Panel.JP_LV_GV = JP_LV_G.Text

        Main_Panel.JP_LV_CHV = JP_LV_B1.Text
        Main_Panel.JP_LV_TV = JP_LV_T.Text

        Main_Panel.NO_TOP_JP_LVV = NO_TOP_JP_LV.Text
        Main_Panel.JP_T_L_LVV = JP_T_L_LV.Text
        Main_Panel.JP_T_H_LVV = JP_T_H_LV.Text
        Main_Panel.JP_T_T_LVV = JP_T_T_LV.Text

        Main_Panel.NO_JP_LVV = NO_JP_LV.SelectedItem
        Main_Panel.LV_JP1V = LV_JP1.Text
        Main_Panel.LV_JP2V = LV_JP2.Text
        Main_Panel.LV_JP3V = LV_JP3.Text
        Main_Panel.LV_JP4V = LV_JP4.Text

        'FLAT STIFFNER
        Main_Panel.HV_STNR_TYPEV = HV_STNR_TYPE.SelectedItem
        'Main_Panel.HV_STNR_THKV = HV_STNR_THK.Text
        'Main_Panel.F_STNR_AV = F_STNR_A.Text
        'Main_Panel.F_STNR_BV = F_STNR_B.Text
        'Main_Panel.F_STNR_DIAV = F_STNR_DIA.Text
        'Main_Panel.F_STNR_R1V = F_STNR_R1.Text
        'Main_Panel.F_STNR_R2V = F_STNR_R2.Text
        'Main_Panel.F_STNR_C = F_STNR_C.Text
        'Main_Panel.F_STNR_D = F_STNR_D.Text
        'Main_Panel.F_STNR_E = F_STNR_E.Text
        'Main_Panel.F_STNR_F = F_STNR_F.Text
        'Main_Panel.CURB_HOLECUTV = CURB_HOLECUT.SelectedItem
        'Main_Panel.CURB_HOLE_WV = CURB_HOLE_W.Text
        'Main_Panel.CURB_HOLE_HV = CURB_HOLE_H.Text
        'Main_Panel.HV_DIM1V = HV_DIM1.Text
        'Main_Panel.HV_DIM2V = HV_DIM2.Text
        'Main_Panel.HV_DIM3V = HV_DIM3.Text

        'Main_Panel.F_STNR_CV = F_STNR_C.Text
        'Main_Panel.F_STNR_DV = F_STNR_D.Text
        'Main_Panel.F_STNR_EV = F_STNR_E.Text
        'Main_Panel.F_STNR_FV = F_STNR_F.Text

        Main_Panel.LV_STNR_TYPEV = LV_STNR_TYPE.SelectedItem
        'Main_Panel.LV_STNR_THKV = LV_STNR_THK.Text
        'Main_Panel.LV_F_STNR_AV = LV_F_STNR_A.Text
        'Main_Panel.LV_F_STNR_BV = LV_F_STNR_B.Text
        'Main_Panel.LV_F_STNR_DIAV = LV_F_STNR_DIA.Text
        'Main_Panel.LV_F_STNR_R1V = LV_F_STNR_R1.Text
        'Main_Panel.LV_F_STNR_R2V = LV_F_STNR_R2.Text
        'Main_Panel.LV_F_STNR_C = LV_F_STNR_C.Text
        'Main_Panel.LV_F_STNR_D = LV_F_STNR_D.Text
        'Main_Panel.LV_F_STNR_E = LV_F_STNR_E.Text
        'Main_Panel.LV_F_STNR_F = LV_F_STNR_F.Text
        'Main_Panel.CURB_HOLECUT_LVV = CURB_HOLECUT_LV.SelectedItem
        'Main_Panel.LV_CURB_HOLE_WV = LV_CURB_HOLE_W.Text
        'Main_Panel.LV_CURB_HOLE_HV = LV_CURB_HOLE_H.Text
        'Main_Panel.LV_DIM1V = LV_DIM1.Text
        'Main_Panel.LV_DIM2V = LV_DIM2.Text
        'Main_Panel.LV_DIM3V = LV_DIM3.Text

        'Main_Panel.LV_F_STNR_CV = LV_F_STNR_C.Text
        'Main_Panel.LV_F_STNR_DV = LV_F_STNR_D.Text
        'Main_Panel.LV_F_STNR_EV = LV_F_STNR_E.Text
        'Main_Panel.LV_F_STNR_FV = LV_F_STNR_F.Text

        Main_Panel.OLTC_STNR_TYPEV = OLTC_STNR_TYPE.SelectedItem
        'Main_Panel.OLTC_STFR_THKV = OLTC_STFR_THK.Text
        'Main_Panel.NO_OLTC_STNRV = NO_OLTC_STNR.SelectedItem
        'Main_Panel.OLTC_F_STNR_AV = OLTC_F_STNR_A.Text
        'Main_Panel.OLTC_F_STNR_BV = OLTC_F_STNR_B.Text
        'Main_Panel.OLTC_F_STNR_CV = OLTC_F_STNR_C.Text
        'Main_Panel.OLTC_F_STNR_DV = OLTC_F_STNR_D.Text
        'Main_Panel.OLTC_F_STNR_RV = OLTC_F_STNR_R.Text
        'Main_Panel.OLTC_F_STNR_CHV = OLTC_F_STNR_CH.Text
        'Main_Panel.OLTC_DIM1V = OLTC_DIM1.Text
        'Main_Panel.OLTC_DIM2V = OLTC_DIM2.Text
        'Main_Panel.OLTC_DIM3V = OLTC_DIM3.Text

        Main_Panel.OLTC_STNR2_TYPEV = OLTC_STNR2_TYPE.SelectedItem
        'Main_Panel.NO_OLTC_STNR2V = NO_OLTC_STNR2.SelectedItem
        'Main_Panel.OLTC_F_STNR2_TV = OLTC_F_STNR2_T.Text
        'Main_Panel.OLTC_F_STNR2_AV = OLTC_F_STNR2_A.Text
        'Main_Panel.OLTC_F_STNR2_BV = OLTC_F_STNR2_B.Text
        'Main_Panel.OLTC_STNR2_DIM1V = OLTC_STNR2_DIM1.Text
        'Main_Panel.OLTC_STNR2_DIM2V = OLTC_STNR2_DIM2.Text
        'Main_Panel.OLTC_STNR2_DIM3V = OLTC_STNR2_DIM3.Text

        Main_Panel.OLTC_END_STNR_TYPV = OLTC_END_STNR_TYP.SelectedItem
        'Main_Panel.NO_OLTC_END_STNRV = NO_OLTC_END_STNR.SelectedItem
        'Main_Panel.OLTC_END_AV = OLTC_END_A.Text
        'Main_Panel.OLTC_END_BV = OLTC_END_B.Text
        'Main_Panel.OLTC_END_THKV = OLTC_END_THK.Text
        'Main_Panel.OLTC_END_DIM1V = OLTC_END_DIM1.Text
        'Main_Panel.OLTC_END_DIM2V = OLTC_END_DIM2.Text
        'Main_Panel.OLTC_END_DIM3V = OLTC_END_DIM3.Text
        'Main_Panel.OLTC_END_DIM4V = OLTC_END_DIM4.Text
        'Main_Panel.OLTC_END_DIM4V = OLTC_END_DIM5.Text

        'OTHER PARTS1
        Main_Panel.SAMP_PIPE_SIZEV = SAMP_PIPE_SIZE.SelectedItem
        Main_Panel.SAM_VAL_POSV = SAM_VAL_POS.SelectedItem
        Main_Panel.SV_XAXISV = SV_XAXIS.Text

        Main_Panel.FILTER_VALUVE_SIZEV = FILTER_VALUVE_SIZE.SelectedItem
        Main_Panel.FILTER_VAL_POSV = FILTER_VAL_POS.SelectedItem
        Main_Panel.FV_XAXISV = FV_XAXIS.Text

        Main_Panel.DRAIN_VALVE_SIZEV = DRAIN_VALVE_SIZE.SelectedItem
        Main_Panel.DRAIN_VALVE_POSV = DRAIN_VALVE_POS.SelectedItem
        Main_Panel.DV_XAXISV = DV_XAXIS.Text
        Main_Panel.DV_YAXISV = DV_YAXIS.Text


        Main_Panel.NIFPS_REQV = NIFPS_REQ.Text

        Main_Panel.NIFPS_VALSIZE = NIFPS_VALSIZE.Text

        Main_Panel.N2_filling_req_HVV = N2_filling_req_HV.SelectedItem
        Main_Panel.N2_filling_req_LVV = N2_filling_req_LV.SelectedItem
        Main_Panel.N2_filling_req_OLTC_oppV = N2_filling_req_OLTC_opp.SelectedItem
        Main_Panel.N2_filling_req_OLTCV = N2_filling_req_OLTC.SelectedItem
        Main_Panel.N2_qty_HVV = N2_qty_HV.SelectedItem
        Main_Panel.N2_qty_LVV = N2_qty_LV.SelectedItem
        Main_Panel.N2_qty_OLTC_oppV = N2_qty_OLTC_opp.SelectedItem
        Main_Panel.N2_qty_OLTCV = N2_qty_OLTC.SelectedItem
        Main_Panel.HV_X1_AXISV = HV_X1_AXIS.Text
        Main_Panel.LV_X1_AXISV = LV_X1_AXIS.Text
        Main_Panel.OLTC_opp_X1_AXISV = OLTC_opp_X1_AXIS.Text
        Main_Panel.OLTC_X1_AXISV = OLTC_X1_AXIS.Text
        Main_Panel.HV_X2_AXISV = HV_X2_AXIS.Text
        Main_Panel.LV_X2_AXISV = LV_X2_AXIS.Text
        Main_Panel.OLTC_opp_X2_AXISV = OLTC_opp_X2_AXIS.Text
        Main_Panel.OLTC_X2_AXISV = OLTC_X2_AXIS.Text
        Main_Panel.HV_X3_AXISV = HV_X3_AXIS.Text
        Main_Panel.LV_X3_AXISV = LV_X3_AXIS.Text
        Main_Panel.OLTC_opp_X3_AXISV = OLTC_opp_X3_AXIS.Text
        Main_Panel.OLTC_X3_AXISV = OLTC_X3_AXIS.Text

        Main_Panel.GAUDE_PIN_SIZEV = GAUDE_PIN_SIZE.SelectedItem


        Main_Panel.EARTHINGPAD_HVV = EARTHINGPAD_HV.SelectedItem
        Main_Panel.EARTHINGPAD_POSV = EARTHINGPAD_POS.SelectedItem
        Main_Panel.HV_XAXISV = HV_XAXIS.Text
        Main_Panel.HV_YAXISV = HV_YAXIS.Text
        Main_Panel.LV_XAXISV = LV_XAXIS.Text
        Main_Panel.LV_YAXISV = LV_YAXIS.Text

        Main_Panel.earthing_pad_HV_type = earthing_pad_HV_type.SelectedItem
        Main_Panel.earthing_pad_LV_type = earthing_pad_LV_type.SelectedItem

        Main_Panel.NO_OF_PORT_BOT = NO_OF_PORT_BOT.SelectedItem

    End Sub



    Public Sub Load_btm_tank_inputs()

        'bottom tank

        btm_tk_type.SelectedItem = Main_Panel.btm_tk_typeV
        btm_tk_curb_pro.SelectedItem = Main_Panel.btm_tk_curb_proV
        curb_size.SelectedItem = Main_Panel.curb_sizev

        btm_plt_ht.SelectedItem = Main_Panel.btm_plt_htV
        btm_tk_i_len.Text = Main_Panel.btm_tk_i_lenV
        btm_plt_len.Text = Main_Panel.btm_plt_lenV
        btm_plt_thk.Text = Main_Panel.btm_plt_thkV
        btm_tk_i_wdh.Text = Main_Panel.btm_tk_i_wdhV
        btm_plt_width.Text = Main_Panel.btm_plt_widthV
        Side_Plt_thk.Text = Main_Panel.Side_Plt_thkV
        OLTC_pck_len.Text = Main_Panel.OLTC_pck_lenV
        OLTC_pck_wdh.Text = Main_Panel.OLTC_pck_wdhV
        Cus_ht.Text = Main_Panel.Cus_htV
        tk_curb_wdh.Text = Main_Panel.tk_curb_wdhV
        Curb_wdh_LV.Text = Main_Panel.Curb_wdh_LVV
        Curb_OLTC_End.Text = Main_Panel.Curb_OLTC_EndV
        Curb_wdh_D.Text = Main_Panel.Curb_wdh_DV

        GPX1.Text = Main_Panel.GPX1V
        GPX2.Text = Main_Panel.GPX2V
        GPY1.Text = Main_Panel.GPY1V
        GPY2.Text = Main_Panel.GPY2V

        btm_len_chkbx.Checked = Main_Panel.btm_len_chkbxV
        btm_plt_wdt_ck_bx.Checked = Main_Panel.btm_plt_wdt_ck_bxV

        'underbase
        No_of_Rails.SelectedItem = Main_Panel.No_of_RailsV
        No_of_Rails1.SelectedItem = Main_Panel.No_of_Rails1V
        CC_LENGTH.SelectedItem = Main_Panel.CC_LENGTHV
        Roller_Custom.Text = Main_Panel.Roller_CustomV
        roller_gauge_cmb.SelectedItem = Main_Panel.roller_gauge_cmbV
        roller_orientation.SelectedItem = Main_Panel.roller_orientationV

        X.Text = Main_Panel.XV

        MP_len.Text = Main_Panel.MP_lenV
        MP_Wth.Text = Main_Panel.MP_WthV
        MP_Thk.Text = Main_Panel.MP_ThkV
        Dim_L1.Text = Main_Panel.Dim_L1V

        BS1_WDH_1.Text = Main_Panel.BS1_WDH_1V
        BS2_WDH_1.Text = Main_Panel.BS2_WDH_1V

        BS_R.Text = Main_Panel.BS_RV
        BS_L.Text = Main_Panel.BS_LV
        'BS3_WDH_1.Text = Main_Panel.BS3_WDH_1V

        BS1_HT_WS.Text = Main_Panel.BS1_HT_WSV
        BS2_HT_WS.Text = Main_Panel.BS2_HT_WSV
        'BS3_HT_WS.Text = Main_Panel.BS3_HT_WSV

        BS1_THK_1.Text = Main_Panel.BS1_THK_1V
        BS2_THK_1.Text = Main_Panel.BS2_THK_1V
        'BS3_THK_1.Text = Main_Panel.BS3_THK_1V

        prob_support_cmb.SelectedItem = Main_Panel.prob_support_cmbV
        prob_no.SelectedItem = Main_Panel.prob_noV
        PROB_HT.Text = Main_Panel.PROB_HTV
        Prob_X1axis.Text = Main_Panel.Prob_X1axisV
        Prob_X2axis.Text = Main_Panel.Prob_X2axisV
        Prob_Y1axis.Text = Main_Panel.Prob_Y1axisV
        Prob_Y2axis.Text = Main_Panel.Prob_Y2axisV

        HS_CC_len.Text = Main_Panel.HS_CC_lenV
        Distance_A.Text = Main_Panel.Distance_AV
        Distance_B.Text = Main_Panel.Distance_BV

        'FLAT STIFFNER
        HV_STNR_TYPE.SelectedItem = Main_Panel.HV_STNR_TYPEV
        HV_FS_A.Text = Main_Panel.HV_FS_AV
        HV_FS_B.Text = Main_Panel.HV_FS_BV
        HV_FS_A1.Text = Main_Panel.HV_FS_A1V
        HV_FS_B1.Text = Main_Panel.HV_FS_B1V

        'LV
        LV_STNR_TYPE.SelectedItem = Main_Panel.LV_STNR_TYPEV
        LV_FS_A.Text = Main_Panel.LV_FS_AV
        LV_FS_B.Text = Main_Panel.LV_FS_BV
        LV_FS_A1.Text = Main_Panel.LV_FS_A1V
        LV_FS_B1.Text = Main_Panel.LV_FS_B1V

        'OLTC S1
        OLTC_STNR_TYPE.Text = Main_Panel.OLTC_STNR_TYPEV
        OLTC_S1_A.Text = Main_Panel.OLTC_S1_AV
        OLTC_S1_B.Text = Main_Panel.OLTC_S1_BV
        OLTC_S1_A1.Text = Main_Panel.OLTC_S1_A1V
        OLTC_S1_B1.Text = Main_Panel.OLTC_S1_B1V
        OLTC_S1_T.Text = Main_Panel.OLTC_S1_TV


        OS1D1.Text = Main_Panel.OS1D1V
        OS1D2.Text = Main_Panel.OS1D2V

        'OLTC S2
        OLTC_STNR2_TYPE.Text = Main_Panel.OLTC_STNR2_TYPEV
        OLTC_S2_A.Text = Main_Panel.OLTC_S2_AV
        OLTC_S2_B.Text = Main_Panel.OLTC_S2_BV
        OLTC_S2_A1.Text = Main_Panel.OLTC_S2_A1V

        OS2D1.Text = Main_Panel.OS2D1V
        OS2D2.Text = Main_Panel.OS2D2V

        'OLTC END
        OLTC_END_STNR_TYP.Text = Main_Panel.OLTC_END_STNR_TYPV
        OLTC_END_A.Text = Main_Panel.OLTC_END_AV
        OLTC_END_B.Text = Main_Panel.OLTC_END_BV
        OLTC_END_A1.Text = Main_Panel.OLTC_END_A1V

        OED1.Text = Main_Panel.OED1V
        OED2.Text = Main_Panel.OED2V
        OED3.Text = Main_Panel.OED3V
        OED4.Text = Main_Panel.OED4V

        Stiffner_count_BT.SelectedItem = Main_Panel.Stiffner_count_BT
        oltc_stiffner1_count.SelectedItem = Main_Panel.oltc_stiffner1_count
        oltc_stiffner2_count.SelectedItem = Main_Panel.oltc_stiffner2_count

        stiffner_count_HV.SelectedItem = Main_Panel.stiffner_count_HV
        HVD1.Text = Main_Panel.HVD1
        HVD2.Text = Main_Panel.HVD2
        HVD3.Text = Main_Panel.HVD3
        HVD4.Text = Main_Panel.HVD4

        stiffner_count_LV.SelectedItem = Main_Panel.stiffner_count_LV
        LVD1.Text = Main_Panel.LVD1
        LVD2.Text = Main_Panel.LVD2
        LVD3.Text = Main_Panel.LVD3
        LVD4.Text = Main_Panel.LVD4

        'JACKING PAD
        'HV


        jack_pad_type.SelectedItem = Main_Panel.jack_pad_typeV
        JP_A.Text = Main_Panel.JP_AV
        JP_B.Text = Main_Panel.JP_BV
        JP_C.Text = Main_Panel.JP_CV
        JP_D.Text = Main_Panel.JP_DV
        JP_E.Text = Main_Panel.JP_EV
        JP_G.Text = Main_Panel.JP_GV

        JP_T.Text = Main_Panel.JP_TV
        JP_B1.Text = Main_Panel.JP_B1V
        JP_BB1.Text = Main_Panel.JP_BB1V
        JP_B2.Text = Main_Panel.JP_B2V

        NO_TOP_JP_HV.SelectedItem = Main_Panel.NO_TOP_JP_HVV
        JP_T_L.Text = Main_Panel.JP_T_LV
        JP_T_H.Text = Main_Panel.JP_T_HV
        JP_T_T.Text = Main_Panel.JP_T_TV

        NO_OF_JP_HV.SelectedItem = Main_Panel.NO_OF_JP_HVV
        HV_JPD1.Text = Main_Panel.HV_JPD1V
        HV_JPD2.Text = Main_Panel.HV_JPD2V
        HV_JPD3.Text = Main_Panel.HV_JPD3V
        HV_JPD4.Text = Main_Panel.HV_JPD4V
        JP_Y_DIST.Text = Main_Panel.JP_Y_DISTV






        'LV
        jack_pad_type1.SelectedItem = Main_Panel.jack_pad_type1
        JP_LV_A.Text = Main_Panel.JP_LV_AV
        JP_LV_B.Text = Main_Panel.JP_LV_BV
        JP_LV_C.Text = Main_Panel.JP_LV_CV
        JP_LV_D.Text = Main_Panel.JP_LV_DV
        JP_LV_E.Text = Main_Panel.JP_LV_EV
        'JP_LV_F.Text = Main_Panel.JP_LV_FV
        JP_LV_G.Text = Main_Panel.JP_LV_GV

        JP_LV_B1.Text = Main_Panel.JP_LV_CHV
        JP_LV_T.Text = Main_Panel.JP_LV_TV

        NO_TOP_JP_LV.Text = Main_Panel.NO_TOP_JP_LVV
        JP_T_L_LV.Text = Main_Panel.JP_T_L_LVV
        JP_T_H_LV.Text = Main_Panel.JP_T_H_LVV
        JP_T_T_LV.Text = Main_Panel.JP_T_T_LVV

        NO_JP_LV.SelectedItem = Main_Panel.NO_JP_LVV
        LV_JP1.Text = Main_Panel.LV_JP1V
        LV_JP2.Text = Main_Panel.LV_JP2V
        LV_JP3.Text = Main_Panel.LV_JP3V
        LV_JP4.Text = Main_Panel.LV_JP4V

        'FLATSTIFFF
        'HV_STNR_TYPE.SelectedItem = Main_Panel.HV_STNR_TYPEV
        'HV_STNR_THK.Text = Main_Panel.HV_STNR_THKV
        'F_STNR_A.Text = Main_Panel.F_STNR_AV
        'F_STNR_B.Text = Main_Panel.F_STNR_BV
        'F_STNR_DIA.Text = Main_Panel.F_STNR_DIAV
        'F_STNR_R1.Text = Main_Panel.F_STNR_R1V
        'F_STNR_R2.Text = Main_Panel.F_STNR_R2V
        'F_STNR_C.Text = Main_Panel.F_STNR_C
        'F_STNR_D.Text = Main_Panel.F_STNR_D
        'F_STNR_E.Text = Main_Panel.F_STNR_E
        'F_STNR_F.Text = Main_Panel.F_STNR_F
        'CURB_HOLECUT.SelectedItem = Main_Panel.CURB_HOLECUTV
        'CURB_HOLE_W.Text = Main_Panel.CURB_HOLE_WV
        'CURB_HOLE_H.Text = Main_Panel.CURB_HOLE_HV
        'HV_DIM1.Text = Main_Panel.HV_DIM1V
        'HV_DIM2.Text = Main_Panel.HV_DIM2V
        'HV_DIM3.Text = Main_Panel.HV_DIM3V

        'LV_STNR_TYPE.SelectedItem = Main_Panel.LV_STNR_TYPEV
        'LV_STNR_THK.Text = Main_Panel.LV_STNR_THKV
        'LV_F_STNR_A.Text = Main_Panel.LV_F_STNR_AV
        'LV_F_STNR_B.Text = Main_Panel.LV_F_STNR_BV
        'LV_F_STNR_DIA.Text = Main_Panel.LV_F_STNR_DIAV
        'LV_F_STNR_R1.Text = Main_Panel.LV_F_STNR_R1V
        'LV_F_STNR_R2.Text = Main_Panel.LV_F_STNR_R2V
        'LV_F_STNR_C.Text = Main_Panel.LV_F_STNR_C
        'LV_F_STNR_D.Text = Main_Panel.LV_F_STNR_D
        'LV_F_STNR_E.Text = Main_Panel.LV_F_STNR_E
        'LV_F_STNR_F.Text = Main_Panel.LV_F_STNR_F
        'CURB_HOLECUT_LV.SelectedItem = Main_Panel.CURB_HOLECUT_LVV
        'LV_CURB_HOLE_W.Text = Main_Panel.LV_CURB_HOLE_WV
        'LV_CURB_HOLE_H.Text = Main_Panel.LV_CURB_HOLE_HV
        'LV_DIM1.Text = Main_Panel.LV_DIM1V
        'LV_DIM2.Text = Main_Panel.LV_DIM2V
        'LV_DIM3.Text = Main_Panel.LV_DIM3V

        'OLTC_STNR_TYPE.SelectedItem = Main_Panel.OLTC_STNR_TYPEV
        'OLTC_STFR_THK.Text = Main_Panel.OLTC_STFR_THKV
        'NO_OLTC_STNR.SelectedItem = Main_Panel.NO_OLTC_STNRV
        'OLTC_F_STNR_A.Text = Main_Panel.OLTC_F_STNR_AV
        'OLTC_F_STNR_B.Text = Main_Panel.OLTC_F_STNR_BV
        'OLTC_F_STNR_C.Text = Main_Panel.OLTC_F_STNR_CV
        'OLTC_F_STNR_D.Text = Main_Panel.OLTC_F_STNR_DV
        'OLTC_F_STNR_R.Text = Main_Panel.OLTC_F_STNR_RV
        'OLTC_F_STNR_CH.Text = Main_Panel.OLTC_F_STNR_CHV
        'OLTC_DIM1.Text = Main_Panel.OLTC_DIM1V
        'OLTC_DIM2.Text = Main_Panel.OLTC_DIM2V
        'OLTC_DIM3.Text = Main_Panel.OLTC_DIM3V

        'OLTC_STNR2_TYPE.SelectedItem = Main_Panel.OLTC_STNR2_TYPEV
        'NO_OLTC_STNR2.SelectedItem = Main_Panel.NO_OLTC_STNR2V
        'OLTC_F_STNR2_T.Text = Main_Panel.OLTC_F_STNR2_TV
        'OLTC_F_STNR2_A.Text = Main_Panel.OLTC_F_STNR2_AV
        'OLTC_F_STNR2_B.Text = Main_Panel.OLTC_F_STNR2_BV
        'OLTC_STNR2_DIM1.Text = Main_Panel.OLTC_STNR2_DIM1V
        'OLTC_STNR2_DIM2.Text = Main_Panel.OLTC_STNR2_DIM2V
        'OLTC_STNR2_DIM3.Text = Main_Panel.OLTC_STNR2_DIM3V

        'OLTC_END_STNR_TYP.SelectedItem = Main_Panel.OLTC_END_STNR_TYPV
        'NO_OLTC_END_STNR.SelectedItem = Main_Panel.NO_OLTC_END_STNRV
        'OLTC_END_A.Text = Main_Panel.OLTC_END_AV
        'OLTC_END_B.Text = Main_Panel.OLTC_END_BV
        'OLTC_END_THK.Text = Main_Panel.OLTC_END_THKV
        'OLTC_END_DIM1.Text = Main_Panel.OLTC_END_DIM1V
        'OLTC_END_DIM2.Text = Main_Panel.OLTC_END_DIM2V
        'OLTC_END_DIM3.Text = Main_Panel.OLTC_END_DIM3V
        'OLTC_END_DIM4.Text = Main_Panel.OLTC_END_DIM4V
        'OLTC_END_DIM5.Text = Main_Panel.OLTC_END_DIM4V

        'F_STNR_C.Text = Main_Panel.F_STNR_CV
        'F_STNR_D.Text = Main_Panel.F_STNR_DV
        'F_STNR_E.Text = Main_Panel.F_STNR_EV
        'F_STNR_F.Text = Main_Panel.F_STNR_FV

        'LV_F_STNR_C.Text = Main_Panel.LV_F_STNR_CV
        'LV_F_STNR_D.Text = Main_Panel.LV_F_STNR_DV
        'LV_F_STNR_E.Text = Main_Panel.LV_F_STNR_EV
        'LV_F_STNR_F.Text = Main_Panel.LV_F_STNR_FV

        'OTHERPARTS 1
        SAMP_PIPE_SIZE.SelectedItem = Main_Panel.SAMP_PIPE_SIZEV
        SAM_VAL_POS.SelectedItem = Main_Panel.SAM_VAL_POSV
        SV_XAXIS.Text = Main_Panel.SV_XAXISV

        FILTER_VALUVE_SIZE.SelectedItem = Main_Panel.FILTER_VALUVE_SIZEV
        FILTER_VAL_POS.SelectedItem = Main_Panel.FILTER_VAL_POSV
        FV_XAXIS.Text = Main_Panel.FV_XAXISV

        DRAIN_VALVE_SIZE.SelectedItem = Main_Panel.DRAIN_VALVE_SIZEV
        DRAIN_VALVE_POS.SelectedItem = Main_Panel.DRAIN_VALVE_POSV
        DV_XAXIS.Text = Main_Panel.DV_XAXISV
        DV_YAXIS.Text = Main_Panel.DV_YAXISV

        NIFPS_REQ.Text = Main_Panel.NIFPS_REQV

        NIFPS_VALSIZE.Text = Main_Panel.NIFPS_VALSIZE
        N2_filling_req_HV.SelectedItem = Main_Panel.N2_filling_req_HVV
        N2_filling_req_LV.SelectedItem = Main_Panel.N2_filling_req_LVV
        N2_filling_req_OLTC_opp.SelectedItem = Main_Panel.N2_filling_req_OLTC_oppV
        N2_filling_req_OLTC.SelectedItem = Main_Panel.N2_filling_req_OLTCV
        N2_qty_HV.SelectedItem = Main_Panel.N2_qty_HVV
        N2_qty_LV.SelectedItem = Main_Panel.N2_qty_LVV
        N2_qty_OLTC_opp.SelectedItem = Main_Panel.N2_qty_OLTC_oppV
        N2_qty_OLTC.SelectedItem = Main_Panel.N2_qty_OLTCV
        HV_X1_AXIS.Text = Main_Panel.HV_X1_AXISV
        LV_X1_AXIS.Text = Main_Panel.LV_X1_AXISV
        OLTC_opp_X1_AXIS.Text = Main_Panel.OLTC_opp_X1_AXISV
        OLTC_X1_AXIS.Text = Main_Panel.OLTC_X1_AXISV
        HV_X2_AXIS.Text = Main_Panel.HV_X2_AXISV
        LV_X2_AXIS.Text = Main_Panel.LV_X2_AXISV
        OLTC_opp_X2_AXIS.Text = Main_Panel.OLTC_opp_X2_AXISV
        OLTC_X2_AXIS.Text = Main_Panel.OLTC_X2_AXISV
        HV_X3_AXIS.Text = Main_Panel.HV_X3_AXISV
        LV_X3_AXIS.Text = Main_Panel.LV_X3_AXISV
        OLTC_opp_X3_AXIS.Text = Main_Panel.OLTC_opp_X3_AXISV
        OLTC_X3_AXIS.Text = Main_Panel.OLTC_X3_AXISV

        GAUDE_PIN_SIZE.SelectedItem = Main_Panel.GAUDE_PIN_SIZEV

        EARTHINGPAD_HV.SelectedItem = Main_Panel.EARTHINGPAD_HVV
        EARTHINGPAD_POS.SelectedItem = Main_Panel.EARTHINGPAD_POSV
        HV_XAXIS.Text = Main_Panel.HV_XAXISV
        HV_YAXIS.Text = Main_Panel.HV_YAXISV
        LV_XAXIS.Text = Main_Panel.LV_XAXISV
        LV_YAXIS.Text = Main_Panel.LV_YAXISV

        earthing_pad_HV_type.SelectedItem = Main_Panel.earthing_pad_HV_type
        earthing_pad_LV_type.SelectedItem = Main_Panel.earthing_pad_LV_type

        NO_OF_PORT_BOT.SelectedItem = Main_Panel.NO_OF_PORT_BOT

    End Sub







    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles BT_closebtn_pg1.Click
    '    save_btm_tank_inputs()
    '    Create_Project.Show()
    '    Me.Close()

    'End Sub

    Private Sub Form3_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Create_Project.Show()
        save_btm_tank_inputs()

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs)

        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles BT_Nextbtn_pg1.Click

        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage8")

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage1")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab("TabPage1")
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab("TabPage3")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab("TabPage2")
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub





    Private Sub Button25_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub





    Private Sub Button28_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab("TabPage4")
    End Sub





    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Next_pg5.Click
        save_btm_tank_inputs()
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles BT_backbtn_pg4.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage2")
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        TabControl4.SelectTab("TabPage19")
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs)
        TabControl4.SelectTab("TabPage19")
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs)
        TabControl4.SelectTab("TabPage19")
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs)
        TabControl4.SelectTab("TabPage17")
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs)
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub BT_backbtn_pg1_Click(sender As Object, e As EventArgs) Handles BT_backbtn_pg1.Click
        save_btm_tank_inputs()
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub BT_backbtn_pg2_Click(sender As Object, e As EventArgs) Handles BT_backbtn_pg2.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage3")
    End Sub

    Private Sub BT_nextbtn_pg2_Click(sender As Object, e As EventArgs) Handles BT_nextbtn_pg2.Click
        save_btm_tank_inputs()
        TabControl3.SelectTab("TabPage7")

        JP_PIC_BOX_2.Image = My.Resources.LV_STIFNER
        'BT_nextbtn_pg2_2.Enabled = True
    End Sub

    Private Sub BT_backbtn_pg2_2_Click(sender As Object, e As EventArgs) Handles BT_backbtn_pg2_2.Click
        save_btm_tank_inputs()
        'TabControl3.SelectTab("TabPage7")
        TabControl3.SelectTab("TabPage5")
    End Sub

    Private Sub BT_nextbtn_pg2_2_Click(sender As Object, e As EventArgs) Handles BT_nextbtn_pg2_2.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage4")
    End Sub



    Private Sub BT_closebtn_pg2_2_Click(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub BT_closebtn_pg5_1_Click(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        Create_Project.Show()
        Me.Close()
    End Sub

    Private Sub BT_nextbtn_pg5_2_Click(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        TabControl4.SelectTab("TabPage18")
        'BT_nextbtn_pg5_3.Enabled = False

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Next_pg3.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage2")
        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage8")
    End Sub




    Private Sub BT_nextbtn_pg5_1_Click(sender As Object, e As EventArgs) Handles BT_nextbtn_pg5_1.Click

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()


        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If


        TabControl1.SelectTab("TabPage3")
    End Sub


    Private Sub BT_backbtn_pg5_1_Click(sender As Object, e As EventArgs) Handles BT_backbtn_pg5_1.Click
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage1")
    End Sub


    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
        save_btm_tank_inputs()
        TabControl1.SelectTab("TabPage4")
    End Sub


    Public Sub Load_BottomTank_BT()

        Try


            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Bottom_Tank_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                btm_tk_type.SelectedItem = reader("btm_tk_type").ToString()
                btm_tk_curb_pro.SelectedItem = reader("btm_tk_curb_pro").ToString()
                btm_tk_curb_pro.SelectedItem = reader("btm_tk_curb_pro").ToString()

                curb_size.SelectedItem = reader("curb_size").ToString()

                btm_plt_ht.SelectedItem = reader("btm_plt_ht").ToString()
                Cus_ht.Text = reader("Cus_ht").ToString()
                btm_tk_i_len.Text = reader("btm_tk_i_len").ToString()
                btm_plt_len.Text = reader("btm_plt_len").ToString()
                btm_plt_thk.Text = reader("btm_plt_thk").ToString()
                btm_plt_width.Text = reader("btm_plt_width").ToString()
                Side_Plt_thk.Text = reader("Side_Plt_thk").ToString()
                btm_tk_i_wdh.Text = reader("btm_tk_i_wdh").ToString()
                
                OLTC_pck_len.Text = reader("OLTC_pck_len").ToString()
                OLTC_pck_wdh.Text = reader("OLTC_pck_wdh").ToString()
                tk_curb_wdh.Text = reader("tk_curb_wdh").ToString()
                Curb_wdh_LV.Text = reader("Curb_wdh_LV").ToString()
                Curb_OLTC_End.Text = reader("Curb_OLTC_End").ToString()

                Curb_wdh_D.Text = reader("Curb_wdh_D").ToString()

                GAUDE_PIN_SIZE.SelectedItem = reader("GAUDE_PIN_SIZE").ToString()
                GPX1.Text = reader("GPX1").ToString()
                GPX2.Text = reader("GPX2").ToString()
                GPY1.Text = reader("GPY1").ToString()
                GPY2.Text = reader("GPY2").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_UnderBase_BT()

        Try


            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from UnderBase_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                No_of_Rails.SelectedItem = reader("No_of_Rails").ToString()
                No_of_Rails1.SelectedItem = reader("No_of_Rails1").ToString()
                CC_LENGTH.SelectedItem = reader("CC_LENGTH").ToString()
                Roller_Custom.Text = reader("Roller_Custom").ToString()
                roller_gauge_cmb.SelectedItem = reader("roller_gauge_cmb").ToString()
                roller_orientation.SelectedItem = reader("roller_orientation").ToString()

                BS1_WDH_1.Text = reader("BS1_WDH_1").ToString()
                BS1_HT_WS.Text = reader("BS1_HT_WS").ToString()
                BS1_THK_1.Text = reader("BS1_THK_1").ToString()
                BS2_WDH_1.Text = reader("BS2_WDH_1").ToString()
                BS2_HT_WS.Text = reader("BS2_HT_WS").ToString()
                BS2_THK_1.Text = reader("BS2_THK_1").ToString()
                BS3_WDH_1.Text = reader("BS3_WDH_1").ToString()
                BS3_HT_WS.Text = reader("BS3_HT_WS").ToString()
                BS3_THK_1.Text = reader("BS3_THK_1").ToString()

                HS_CC_len.Text = reader("HS_CC_len").ToString()
                X.Text = reader("X").ToString()
                Distance_A.Text = reader("Distance_A").ToString()
                Distance_B.Text = reader("Distance_B").ToString()

                MP_len.Text = reader("MP_len").ToString()
                MP_Wth.Text = reader("MP_Wth").ToString()
                MP_Thk.Text = reader("MP_Thk").ToString()
                T2.Text = reader("T2").ToString()
                Dim_L1.Text = reader("Dim_L1").ToString()

                prob_support_cmb.SelectedItem = reader("prob_support_cmb").ToString()
                prob_no.SelectedItem = reader("prob_no").ToString()
                PROB_HT.Text = reader("PROB_HT").ToString()
                Prob_X1axis.Text = reader("Prob_X1axis").ToString()
                Prob_X2axis.Text = reader("Prob_X2axis").ToString()
                Prob_Y1axis.Text = reader("Prob_Y1axis").ToString()
                Prob_Y2axis.Text = reader("Prob_Y2axis").ToString()

                BS_L.Text = reader("BS_L").ToString()
                BS_R.Text = reader("BS_R").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub



    Public Sub Load_FlatStiffner_BT()

        Try

            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from FlatStiffner_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                HV_STNR_TYPE.SelectedItem = reader("HV_STNR_TYPE").ToString()
                HV_FS_A.Text = reader("HV_FS_A").ToString()
                HV_FS_A1.Text = reader("HV_FS_A1").ToString()
                HV_FS_B.Text = reader("HV_FS_B").ToString()
                HV_FS_B1.Text = reader("HV_FS_B1").ToString()

                LV_STNR_TYPE.SelectedItem = reader("LV_STNR_TYPE").ToString()
                LV_FS_A.Text = reader("LV_FS_A").ToString()
                LV_FS_A1.Text = reader("LV_FS_A1").ToString()
                LV_FS_B.Text = reader("LV_FS_B").ToString()
                LV_FS_B1.Text = reader("LV_FS_B1").ToString()

                OLTC_STNR_TYPE.SelectedItem = reader("OLTC_STNR_TYPE").ToString()
                OLTC_S1_A.Text = reader("OLTC_S1_A").ToString()
                OLTC_S1_A1.Text = reader("OLTC_S1_A1").ToString()
                OLTC_S1_B.Text = reader("OLTC_S1_B").ToString()
                OLTC_S1_B1.Text = reader("OLTC_S1_B1").ToString()
                OLTC_S1_T.Text = reader("OLTC_S1_T").ToString()
                OS1D1.Text = reader("OS1D1").ToString()
                OS1D2.Text = reader("OS1D2").ToString()

                OLTC_STNR2_TYPE.SelectedItem = reader("OLTC_STNR2_TYPE").ToString()
                OLTC_S2_A.Text = reader("OLTC_S2_A").ToString()
                OLTC_S2_B.Text = reader("OLTC_S2_B").ToString()
                OLTC_S2_A1.Text = reader("OLTC_S2_A1").ToString()
                OS2D1.Text = reader("OS2D1").ToString()
                OS2D2.Text = reader("OS2D2").ToString()

                OLTC_END_STNR_TYP.SelectedItem = reader("OLTC_END_STNR_TYP").ToString()
                Stiffner_count_BT.SelectedItem = reader("Stiffner_count_BT").ToString()
                oltc_stiffner1_count.SelectedItem = reader("oltc_stiffner1_count").ToString()
                oltc_stiffner2_count.SelectedItem = reader("oltc_stiffner2_count").ToString()
                OLTC_END_A.Text = reader("OLTC_END_A").ToString()
                OLTC_END_B.Text = reader("OLTC_END_B").ToString()
                OLTC_END_A1.Text = reader("OLTC_END_A1").ToString()
                OED1.Text = reader("OED1").ToString()
                OED2.Text = reader("OED2").ToString()
                OED3.Text = reader("OED3").ToString()
                OED4.Text = reader("OED4").ToString()

                stiffner_count_HV.SelectedItem = reader("stiffner_count_HV").ToString()
                HVD1.Text = reader("HVD1").ToString()
                HVD2.Text = reader("HVD2").ToString()
                HVD3.Text = reader("HVD3").ToString()
                HVD4.Text = reader("HVD4").ToString()

                stiffner_count_LV.SelectedItem = reader("stiffner_count_LV").ToString()
                LVD1.Text = reader("LVD1").ToString()
                LVD2.Text = reader("LVD2").ToString()
                LVD3.Text = reader("LVD3").ToString()
                LVD4.Text = reader("LVD4").ToString()

                FS_HV_REQ.SelectedItem = reader("FS_HV_REQ").ToString()
                FS_LV_REQ.SelectedItem = reader("FS_LV_REQ").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub
    '21-07-2025
    Public Sub Load_JackingPad_HV_BT()

        Try

            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from JackingPad_HV_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)

            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                'jack_pad_type.SelectedItem = reader("jack_pad_type").ToString()
                Dim value2 As String = reader("jack_pad_type").ToString().Trim()

                ' Use FindStringExact to avoid type mismatch
                Dim index2 As Integer = jack_pad_type.FindStringExact(value2)
                If index2 >= 0 Then
                    ' Temporarily disable event handlers if needed
                    RemoveHandler jack_pad_type.SelectedIndexChanged, AddressOf jack_pad_type_SelectedIndexChanged
                    jack_pad_type.SelectedIndex = index2
                    AddHandler jack_pad_type.SelectedIndexChanged, AddressOf jack_pad_type_SelectedIndexChanged
                Else
                    MsgBox("Dropdown value not found: " & value2)
                End If
                JP_A.Text = reader("JP_A").ToString()
                JP_B.Text = reader("JP_B").ToString()
                JP_T.Text = reader("JP_T").ToString()
                JP_G.Text = reader("JP_G").ToString()
                JP_E.Text = reader("JP_E").ToString()
                JP_D.Text = reader("JP_D").ToString()
                JP_C.Text = reader("JP_C").ToString()
                JP_B1.Text = reader("JP_B1").ToString()
                JP_BB1.Text = reader("JP_BB1").ToString()
                JP_B2.Text = reader("JP_B2").ToString()



                'NO_TOP_JP_HV.SelectedItem = reader("NO_TOP_JP_HV").ToString()
                Dim value1 As String = reader("NO_TOP_JP_HV").ToString().Trim()

                ' Use FindStringExact to avoid type mismatch
                Dim index1 As Integer = NO_TOP_JP_HV.FindStringExact(value1)
                If index1 >= 0 Then
                    ' Temporarily disable event handlers if needed
                    RemoveHandler NO_TOP_JP_HV.SelectedIndexChanged, AddressOf NO_TOP_JP_HV_SelectedIndexChanged
                    NO_TOP_JP_HV.SelectedIndex = index1
                    AddHandler NO_TOP_JP_HV.SelectedIndexChanged, AddressOf NO_TOP_JP_HV_SelectedIndexChanged
                Else
                    MsgBox("Dropdown value not found: " & value1)
                End If
                JP_T_L.Text = reader("JP_T_L").ToString()
                JP_T_H.Text = reader("JP_T_H").ToString()
                JP_T_T.Text = reader("JP_T_T").ToString()




                Dim value As String = reader("NO_OF_JP_HV").ToString().Trim()

                ' Use FindStringExact to avoid type mismatch
                Dim index As Integer = NO_OF_JP_HV.FindStringExact(value)
                If index >= 0 Then
                    ' Temporarily disable event handlers if needed
                    RemoveHandler NO_OF_JP_HV.SelectedIndexChanged, AddressOf NO_OF_JP_HV_SelectedIndexChanged
                    NO_OF_JP_HV.SelectedIndex = index
                    AddHandler NO_OF_JP_HV.SelectedIndexChanged, AddressOf NO_OF_JP_HV_SelectedIndexChanged
                Else
                    MsgBox("Dropdown value not found: " & value)
                End If

                HV_JPD1.Text = reader("HV_JPD1").ToString()
                HV_JPD2.Text = reader("HV_JPD2").ToString()
                HV_JPD3.Text = reader("HV_JPD3").ToString()
                HV_JPD4.Text = reader("HV_JPD4").ToString()
                JP_Y_DIST.Text = reader("JP_Y_DIST").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_JakingPad_LV_BT()

        Try


            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from JakingPad_LV_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                jack_pad_type1.SelectedItem = reader("jack_pad_type1").ToString()
                JP_LV_A.Text = reader("JP_LV_A").ToString()
                JP_LV_B.Text = reader("JP_LV_B").ToString()
                JP_LV_C.Text = reader("JP_LV_C").ToString()
                JP_LV_D.Text = reader("JP_LV_D").ToString()
                JP_LV_E.Text = reader("JP_LV_E").ToString()
                JP_LV_G.Text = reader("JP_LV_G").ToString()
                JP_LV_T.Text = reader("JP_LV_T").ToString()
                JP_LV_B1.Text = reader("JP_LV_B1").ToString()
                JP_LV_BB1.Text = reader("JP_LV_BB1").ToString()
                JP_LV_B2.Text = reader("JP_LV_B2").ToString()

                NO_TOP_JP_LV.SelectedItem = reader("NO_TOP_JP_LV").ToString()
                JP_T_L_LV.Text = reader("JP_T_L_LV").ToString()
                JP_T_T_LV.Text = reader("JP_T_T_LV").ToString()
                JP_T_H_LV.Text = reader("JP_T_H_LV").ToString()

                NO_JP_LV.SelectedItem = reader("NO_JP_LV").ToString()
                JP_Y_DIST_LV.Text = reader("JP_Y_DIST_LV").ToString()
                LV_JP1.Text = reader("LV_JP1").ToString()
                LV_JP2.Text = reader("LV_JP2").ToString()
                LV_JP3.Text = reader("LV_JP3").ToString()
                LV_JP4.Text = reader("LV_JP4").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Accessories_BT()

        Try

            Dim wk_odr_No1 As String = Create_Project.wk_ord_No

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                SAMP_PIPE_SIZE.SelectedItem = reader("SAMP_PIPE_SIZE").ToString()
                SAM_VAL_POS.SelectedItem = reader("SAM_VAL_POS").ToString()
                SV_XAXIS.Text = reader("SV_XAXIS").ToString()

                FILTER_VALUVE_SIZE.SelectedItem = reader("FILTER_VALUVE_SIZE").ToString()
                FILTER_VAL_POS.SelectedItem = reader("FILTER_VAL_POS").ToString()
                FV_XAXIS.Text = reader("FV_XAXIS").ToString()

                DRAIN_VALVE_SIZE.SelectedItem = reader("DRAIN_VALVE_SIZE").ToString()
                DRAIN_VALVE_POS.SelectedItem = reader("DRAIN_VALVE_POS").ToString()
                DV_XAXIS.Text = reader("DV_XAXIS").ToString()
                DV_YAXIS.Text = reader("DV_YAXIS").ToString()

                NIFPS_REQ.SelectedItem = reader("NIFPS_REQ").ToString()
                NIFPS_VALSIZE.SelectedItem = reader("NIFPS_VALSIZE").ToString()

                N2_filling_req_HV.SelectedItem = reader("N2_filling_req_HV").ToString()
                N2_qty_HV.SelectedItem = reader("N2_qty_HV").ToString()
                HV_X1_AXIS.Text = reader("HV_X1_AXIS").ToString()
                HV_X2_AXIS.Text = reader("HV_X2_AXIS").ToString()
                HV_X3_AXIS.Text = reader("HV_X3_AXIS").ToString()

                N2_filling_req_LV.SelectedItem = reader("N2_filling_req_LV").ToString()
                N2_qty_LV.SelectedItem = reader("N2_qty_LV").ToString()
                LV_X1_AXIS.Text = reader("LV_X1_AXIS").ToString()
                LV_X2_AXIS.Text = reader("LV_X2_AXIS").ToString()
                LV_X3_AXIS.Text = reader("LV_X3_AXIS").ToString()

                N2_filling_req_OLTC_opp.SelectedItem = reader("N2_filling_req_OLTC_opp").ToString()
                N2_qty_OLTC_opp.SelectedItem = reader("N2_qty_OLTC_opp").ToString()
                OLTC_opp_X1_AXIS.Text = reader("OLTC_opp_X1_AXIS").ToString()
                OLTC_opp_X2_AXIS.Text = reader("OLTC_opp_X2_AXIS").ToString()
                OLTC_opp_X3_AXIS.Text = reader("OLTC_opp_X3_AXIS").ToString()

                N2_filling_req_OLTC.SelectedItem = reader("N2_filling_req_OLTC").ToString()
                N2_qty_OLTC.SelectedItem = reader("N2_qty_OLTC").ToString()
                OLTC_X1_AXIS.Text = reader("OLTC_X1_AXIS").ToString()
                OLTC_X2_AXIS.Text = reader("OLTC_X2_AXIS").ToString()
                OLTC_X3_AXIS.Text = reader("OLTC_X3_AXIS").ToString()

                EARTHINGPAD_HV.SelectedItem = reader("EARTHINGPAD_HV").ToString()
                HV_XAXIS.Text = reader("HV_XAXIS").ToString()
                HV_YAXIS.Text = reader("HV_YAXIS").ToString()

                EARTHINGPAD_POS.SelectedItem = reader("EARTHINGPAD_POS").ToString()
                LV_XAXIS.Text = reader("LV_XAXIS").ToString()
                LV_YAXIS.Text = reader("LV_YAXIS").ToString()

                earthing_pad_HV_type.SelectedItem = reader("earthing_pad_HV_type").ToString()
                earthing_pad_LV_type.SelectedItem = reader("earthing_pad_LV_type").ToString()

                NO_OF_PORT_BOT.SelectedItem = reader("NO_OF_PORT_BOT").ToString()

                NO_OF_PORT_BOT_X1.Text = reader("NO_OF_PORT_BOT_X1").ToString()
                NO_OF_PORT_BOT_X2.Text = reader("NO_OF_PORT_BOT_X2").ToString()
                NO_OF_PORT_BOT_X3.Text = reader("NO_OF_PORT_BOT_X3").ToString()

                NO_OF_PORT_BOT_Y1.Text = reader("NO_OF_PORT_BOT_Y1").ToString()


                NO_OF_PORT_BOT_LV.SelectedItem = reader("NO_OF_PORT_BOT_LV").ToString()

                NO_OF_PORT_BOT_X1_LV.Text = reader("NO_OF_PORT_BOT_X1_LV").ToString()
                NO_OF_PORT_BOT_X2_LV.Text = reader("NO_OF_PORT_BOT_X2_LV").ToString()
                NO_OF_PORT_BOT_X3_LV.Text = reader("NO_OF_PORT_BOT_X3_LV").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Function SafeDouble(textbox As System.Windows.Forms.TextBox) As Double
        If IsNumeric(textbox.Text.Trim()) Then
            Return CDbl(textbox.Text)
        Else
            Return 0
        End If
    End Function

    Private Sub Bottom_Tank_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Load_OLTC_Pkt_Position()

        'btm_plt_width.Hide()

        btm_plt_wdt_ck_bx.Hide()

        If Main_Panel.btm_tk_typeV = Nothing Then

            btm_tk_type.SelectedItem = "Type 1"
            btm_tk_curb_pro.SelectedItem = "Type 1"
            curb_size.SelectedItem = "150 x 25"

            btm_plt_ht.SelectedItem = "500"

            CC_LENGTH.SelectedItem = "1485"
            roller_gauge_cmb.SelectedItem = "Single- Dia 280"
            roller_orientation.SelectedItem = "Length Wise"
            prob_support_cmb.SelectedItem = "No"
            No_of_Rails1.SelectedItem = "Option 1"


            prob_no.SelectedItem = "2 Nos"
            No_of_Rails.SelectedItem = "3"


            btm_len_chkbx.Checked = False
            btm_plt_wdt_ck_bx.Checked = False

            ver_channel1_cmb.SelectedItem = "Type 1"
            ver_channel1_Y_axis_cmb.SelectedItem = "Bottom tank width"
            ver_channel2_cmb.SelectedItem = "Type 1"
            ver_channel2_Y_axis_cmb.SelectedItem = "Bottom tank width"
            hor_channel_cmb.SelectedItem = "Type 2"

            SAMP_PIPE_SIZE.SelectedItem = "15mm"
            SAM_VAL_POS.SelectedItem = "HV"
            FILTER_VALUVE_SIZE.SelectedItem = "50mm"
            FILTER_VAL_POS.SelectedItem = "HV"
            DRAIN_VALVE_SIZE.SelectedItem = "50mm"
            DRAIN_VALVE_POS.SelectedItem = "HV"

            Btm_TK_pic_box.Image = My.Resources.BT_T1
            No_of_Rails.SelectedItem = "4"
            CC_LENGTH.SelectedItem = "1000 mm"

            jack_pad_type.SelectedItem = "TYPE 1"
            jack_pad_type1.SelectedItem = "TYPE 1"
            NO_OF_JP_HV.SelectedItem = "4"
            NO_JP_LV.SelectedItem = "4"
            NO_TOP_JP_HV.SelectedItem = "2 Nos"
            NO_TOP_JP_LV.SelectedItem = "2 Nos"
            HV_STNR_TYPE.SelectedItem = "Type 1"
            LV_STNR_TYPE.SelectedItem = "Type 1"
            OLTC_STNR_TYPE.SelectedItem = "Type 1"
            OLTC_END_STNR_TYP.SelectedItem = "Type 1"
            OLTC_STNR2_TYPE.SelectedItem = "Type 1"

            'CURB_HOLECUT.SelectedItem = "Yes"
            'CURB_HOLECUT_LV.SelectedItem = "Yes"
            'NO_OLTC_STNR.SelectedItem = "2"
            'NO_OLTC_STNR2.SelectedItem = "1"
            'NO_OLTC_END_STNR.SelectedItem = "3"

            N2_filling_req_HV.SelectedItem = "No"
            N2_filling_req_LV.SelectedItem = "No"
            N2_filling_req_OLTC_opp.SelectedItem = "No"
            N2_filling_req_OLTC.SelectedItem = "No"

            N2_qty_HV.SelectedItem = "1"
            N2_qty_LV.SelectedItem = "1"
            N2_qty_OLTC_opp.SelectedItem = "1"
            N2_qty_OLTC.SelectedItem = "1"

            GAUDE_PIN_SIZE.SelectedItem = "Ø50x75 Lg"


            EARTHINGPAD_HV.SelectedItem = "Yes"
            EARTHINGPAD_POS.SelectedItem = "Yes"



            bt_tk_l.Hide()
            bt_tk_ht.Hide()

            tk_thk.Hide()
            BT_tk_win.Hide()
            bt_tk_lin.Hide()


            Curb_A.Hide()
            Curb_B.Hide()
            Curb_C.Hide()
            Curb_D.Hide()

            GP_Yaxis_1.Hide()
            GP_Yaxis_2.Hide()
            GP_Xaxis_1.Hide()
            GP_Xaxis_2.Hide()

            OLTC_PKT_WDH.Hide()
            OLTC_PKT_LEN.Hide()
            W1.Hide()
            W2.Hide()

            If btm_tk_type.SelectedItem = "Type 1" Then
                btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

                ' bt_tk_wdh_2.Text = Main_Panel.bt_tk_wdh_2V
            ElseIf btm_tk_type.SelectedItem = "Type 2" Then


            End If

            If CC_LENGTH.SelectedItem = "Custom" Then
                Roller_Custom.Enabled = True

                HS_CC_len.Text = Roller_Custom.Text
            Else
                Roller_Custom.Enabled = False
                HS_CC_len.Text = CC_LENGTH.SelectedItem

            End If



            BS3_THK_1.Enabled = False
            BS3_HT_WS.Enabled = False
            BS3_WDH_1.Enabled = False

        Else

            Load_btm_tank_inputs()

        End If

        HS_CC_len.Enabled = False

        'If btm_tk_type.SelectedItem = "Type 1" Then
        '    btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

        '    ' bt_tk_wdh_2.Text = Main_Panel.bt_tk_wdh_2V
        'Else
        '    bt_tk_wdh_2.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

        '    ' btm_plt_width.Text = Main_Panel.btm_plt_widthV
        'End If

        BS3_THK_1.Enabled = False
        BS3_HT_WS.Enabled = False
        BS3_WDH_1.Enabled = False



        prob_no.Enabled = False

        If btm_len_chkbx.Checked Then

            btm_plt_len.Enabled = True
        Else
            btm_plt_len.Enabled = False
        End If


        'If prob_support_cmb.SelectedItem = "Yes" Then

        '    If prob_no.SelectedItem = "2 Nos" Then

        '        Prob_X1axis.Enabled = True
        '        Prob_X2axis.Enabled = True
        '        prob_H.Enabled = True
        '        prob_no.Enabled = True


        '        Prob_Y1axis.Enabled = False
        '        Prob_Y2axis.Enabled = False

        '    Else prob_no.SelectedItem = "4 Nos"
        '        prob_H.Enabled = True
        '        prob_H.Enabled = True

        '        Prob_X1axis.Enabled = True
        '        Prob_X2axis.Enabled = True

        '        Prob_Y1axis.Enabled = True
        '        Prob_Y2axis.Enabled = True
        '    End If

        'End If



        'UB
        If No_of_Rails.SelectedItem = "2" Then

            ' No_of_Rails1.Items.Clear()
            'No_of_Rails1.Items.Add("Option 1")
            ' No_of_Rails1.Items.Add("Option 2")

            'No_of_Rails1.SelectedItem = "Option 1"
            No_of_Rails1.Enabled = True
            Distance_A.Enabled = False
            Distance_B.Enabled = False
            prob_support_cmb.SelectedItem = "Yes"
            prob_support_cmb.Enabled = True

            If No_of_Rails1.SelectedItem = "Option 1" Then

                If prob_support_cmb.SelectedItem = "Yes" Then

                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False


                ElseIf prob_support_cmb.SelectedItem = "No" Then
                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = False
                    Prob_X2axis.Enabled = False
                    PROB_HT.Enabled = False
                    prob_no.Enabled = False


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False

                End If

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                If prob_support_cmb.SelectedItem = "Yes" Then

                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    'prob_no.Enabled = True


                    Prob_Y1axis.Enabled = True
                    Prob_Y2axis.Enabled = True


                ElseIf prob_support_cmb.SelectedItem = "No" Then
                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = False
                    Prob_X2axis.Enabled = False
                    PROB_HT.Enabled = False
                    prob_no.Enabled = False


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False

                End If



            End If





        ElseIf No_of_Rails.SelectedItem = "3" Then

            No_of_Rails1.Enabled = False

            Distance_A.Enabled = True
            Distance_B.Enabled = True


            prob_support_cmb.SelectedItem = "No"
            prob_support_cmb.Enabled = False
            prob_no.Enabled = False
            PROB_HT.Enabled = False
            Prob_X1axis.Enabled = False
            Prob_X2axis.Enabled = False
            Prob_Y1axis.Enabled = False
            Prob_Y2axis.Enabled = False

        End If

        '' Convert text values to Double safely
        'Dim val1 As Double = 0
        'Dim val2 As Double = 0

        'If IsNumeric(BS1_HT_WS.Text) Then
        '    val1 = CDbl(BS1_HT_WS.Text)
        'End If

        'If IsNumeric(MP_Thk.Text) Then
        '    val2 = CDbl(MP_Thk.Text)
        'End If

        '' Calculate the result
        'Dim result As Double = val1 - val2

        '' Set the result in third textbox
        'HV_FS_B.Text = result.ToString()

        'flat stif
        Dim result As Double = SafeDouble(BS1_HT_WS) - SafeDouble(MP_Thk)
        HV_FS_B.Text = result.ToString("F2")


        'If CC_LENGTH.SelectedItem = "Custom" Then

        '    HV_FS_A1.Text = ((btm_tk_i_wdh.Text + HV_FS_A.Text) - (Roller_Custom.Text / 2 + T2.Text))

        'Else

        '    HV_FS_A1.Text = ((btm_tk_i_wdh.Text + HV_FS_A.Text) - (CC_LENGTH.SelectedItem / 2 + T2.Text))
        'End If


        'If btm_plt_ht.SelectedItem = "Custom" Then
        '    HV_FS_B1.Text = ((Cus_ht.Text) + (HV_FS_B.Text))


        'Else

        '    HV_FS_B1.Text = ((btm_plt_ht.SelectedItem) + (HV_FS_B.Text))
        'End If



        HV_X1_AXIS.Enabled = False
        HV_X2_AXIS.Enabled = False
        HV_X3_AXIS.Enabled = False

        LV_X1_AXIS.Enabled = False
        LV_X2_AXIS.Enabled = False
        LV_X3_AXIS.Enabled = False

        OLTC_opp_X1_AXIS.Enabled = False
        OLTC_opp_X2_AXIS.Enabled = False
        OLTC_opp_X3_AXIS.Enabled = False

        OLTC_X1_AXIS.Enabled = False
        OLTC_X2_AXIS.Enabled = False
        OLTC_X3_AXIS.Enabled = False

        'otherparts 1
        If N2_filling_req_HV.SelectedItem = "Yes" Then

            N2_qty_HV.Enabled = True



            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

            If N2_qty_HV.SelectedItem = "1" Then


                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = False
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "2" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "3" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_HV.SelectedItem = "No" Then
            N2_qty_HV.Enabled = False
            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False


        End If

        'lv
        If N2_filling_req_LV.SelectedItem = "Yes" Then

            N2_qty_LV.Enabled = True
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

            If N2_qty_LV.SelectedItem = "1" Then


                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = False
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "2" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "3" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_LV.SelectedItem = "No" Then
            N2_qty_LV.Enabled = False
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

        End If


        'oltc opp
        If N2_filling_req_OLTC_opp.SelectedItem = "Yes" Then

            N2_qty_OLTC_opp.Enabled = True
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False

            If N2_qty_OLTC_opp.SelectedItem = "1" Then


                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = False
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "2" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "3" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = True

            End If


        ElseIf N2_filling_req_OLTC_opp.SelectedItem = "No" Then
            N2_qty_OLTC_opp.Enabled = False
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False


        End If


        'oltc
        If N2_filling_req_OLTC.SelectedItem = "Yes" Then

            N2_qty_OLTC.Enabled = True

            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

            If N2_qty_OLTC.SelectedItem = "1" Then


                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = False
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "2" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "3" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = True

            End If

        ElseIf N2_filling_req_OLTC.SelectedItem = "No" Then
            N2_qty_OLTC.Enabled = False
            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False


        End If

        Load_BottomTank_BT()
        Load_UnderBase_BT()

        Load_FlatStiffner_BT()
        Load_JackingPad_HV_BT()
        Load_JakingPad_LV_BT()
        Load_Accessories_BT()
    End Sub

    Private Sub btm_tk_type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles btm_tk_type.SelectedIndexChanged

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        W1.Hide()
        W2.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then
            btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30


        Else
            btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

            'btm_plt_width.Text = Main_Panel.btm_plt_widthV
        End If


        If btm_tk_type.SelectedItem = "Type 1" Then
            Btm_TK_pic_box.Image = My.Resources.BT_T1

            If btm_plt_wdt_ck_bx.Checked Then
                btm_plt_width.Enabled = True

            Else
                btm_plt_width.Enabled = False

            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then
            Btm_TK_pic_box.Image = My.Resources.BT_T2

            If btm_plt_wdt_ck_bx.Checked Then
                btm_plt_width.Enabled = True

            Else

                btm_plt_width.Enabled = False
            End If


        End If

    End Sub

    Private Sub btm_tk_curb_pro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles btm_tk_curb_pro.SelectedIndexChanged

        Load_OLTC_Pkt_Position()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()
        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_R1

            ElseIf btm_tk_curb_pro.SelectedItem = "Type 2" Then
                Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2LHV 'T2_BTM_TK_LWT1
            End If



        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_L
            ElseIf btm_tk_curb_pro.SelectedItem = "Type 2" Then
                Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2RHV
            End If

        End If



        'If Basic_inputs.OLTC_pkt_position = "Left" Then

        '    'MsgBox("1")


        '    If btm_tk_curb_pro.SelectedItem = "Type 1" Then
        '        Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_R1

        '    Else

        '        Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2LHV1

        '    End If



        'Else


        '    'MsgBox("1.2")

        '    If btm_tk_curb_pro.SelectedItem = "Type 1" Then
        '        Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_L

        '    Else

        '        Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2RHV

        '    End If


        'End If







    End Sub


    Private Sub Roller_CC_Length_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_WDH_1_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_HT_WS_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_THK_1_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_Wth_2_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_HT_LS_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub BS_THK_2_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub HS_CC_len_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub UB_Height_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_height1
    End Sub

    Private Sub Distance_A_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub Distance_B_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub MP_len_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub MP_Wth_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub Dim_L1_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub Dim_B_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.UB_new1
    End Sub

    Private Sub ComboBox3_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.pulling_eye
    End Sub

    Private Sub Pulling_eyes_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.pulling_eye
    End Sub

    Private Sub Pulling_eyes_len_Click(sender As Object, e As EventArgs)
        UB_picture_box.Image = My.Resources.pulling_eye
    End Sub

    Private Sub JP_A_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_B_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_C_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub
    Private Sub JP_D_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_E_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_F_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_G_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub JP_H_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub


    Private Sub JP_CH_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub


    Private Sub JP_T_Click(sender As Object, e As EventArgs) Handles JP_A.Click
        JP_PIC_BOX.Image = My.Resources.JACKINGPAD2
    End Sub

    Private Sub Gusset_HV_B_txt_Click(sender As Object, e As EventArgs) Handles JP_T_L.Click
        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1

    End Sub

    Private Sub TextBox22_Click(sender As Object, e As EventArgs) Handles JP_T_H.Click
        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1
    End Sub

    Private Sub JP_T_T_Click(sender As Object, e As EventArgs) Handles JP_T_T.Click
        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1
    End Sub

    Private Sub NO_TOP_JP_HV_Click(sender As Object, e As EventArgs) Handles NO_TOP_JP_HV.Click
        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1
        JP_HV_B.Hide()
        JP_HV_B1.Hide()
        JP_HV_BB1.Hide()
        JP_HV_50.Hide()
        JP_HV_T.Hide()
        JP_HV_E.Hide()
        JP_HV_E1.Hide()
        JP_HV_G.Hide()
        JP_HV_X.Hide()
        JP_HV_500.Hide()
        JP_HV_R100.Hide()
        JP_HV_B2.Hide()
        JP_HV_A.Hide()
        JP_HV_C.Hide()
        JP_HV_D.Hide()
        JP_HV_C1.Hide()
    End Sub

    Private Sub HV_DIM1_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.HV_STIFNER
    End Sub

    Private Sub HV_DIM2_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.HV_STIFNER
    End Sub

    Private Sub HV_DIM3_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.HV_STIFNER
    End Sub

    Private Sub LV_DIM1_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.LV_STIFNER
    End Sub

    Private Sub LV_DIM2_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.LV_STIFNER
    End Sub

    Private Sub LV_DIM3_Click(sender As Object, e As EventArgs)
        STIFNER_PIC_BOX.Image = My.Resources.LV_STIFNER
    End Sub

    Private Sub NO_OLTC_STNR_SelectedIndexChanged(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_2_STNR

        '    OLTC_DIM1.Enabled = True
        '    OLTC_DIM2.Enabled = True
        '    OLTC_DIM3.Enabled = False

        'ElseIf NO_OLTC_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_3_STNR

        '    OLTC_DIM1.Enabled = True
        '    OLTC_DIM2.Enabled = True
        '    OLTC_DIM3.Enabled = True

        'End If





    End Sub

    Private Sub NO_OLTC_STNR2_SelectedIndexChanged(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR2.SelectedItem = "1" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_1DIM

        '    OLTC_STNR2_DIM1.Enabled = True
        '    OLTC_STNR2_DIM2.Enabled = False
        '    OLTC_STNR2_DIM3.Enabled = False

        'ElseIf NO_OLTC_STNR2.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_2DIM1

        '    OLTC_STNR2_DIM1.Enabled = True
        '    OLTC_STNR2_DIM2.Enabled = True
        '    OLTC_STNR2_DIM3.Enabled = False

        'ElseIf NO_OLTC_STNR2.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_3DIM

        '    OLTC_STNR2_DIM1.Enabled = True
        '    OLTC_STNR2_DIM2.Enabled = True
        '    OLTC_STNR2_DIM3.Enabled = True

        'End If



    End Sub

    Private Sub NO_OLTC_END_STNR_SelectedIndexChanged(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        '    OLTC_END_DIM1.Enabled = True
        '    OLTC_END_DIM2.Enabled = True
        '    OLTC_END_DIM3.Enabled = True
        '    OLTC_END_DIM4.Enabled = False
        '    OLTC_END_DIM5.Enabled = False

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        '    OLTC_END_DIM1.Enabled = True
        '    OLTC_END_DIM2.Enabled = True
        '    OLTC_END_DIM3.Enabled = True
        '    OLTC_END_DIM4.Enabled = True
        '    OLTC_END_DIM5.Enabled = False

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        '    OLTC_END_DIM1.Enabled = True
        '    OLTC_END_DIM2.Enabled = True
        '    OLTC_END_DIM3.Enabled = True
        '    OLTC_END_DIM4.Enabled = True
        '    OLTC_END_DIM5.Enabled = True
        'End If


    End Sub

    Private Sub OLTC_DIM1_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_2_STNR

        'ElseIf NO_OLTC_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_3_STNR

        'End If
    End Sub

    Private Sub OLTC_DIM2_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_2_STNR

        'ElseIf NO_OLTC_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_3_STNR

        'End If
    End Sub

    Private Sub OLTC_DIM3_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_2_STNR

        'ElseIf NO_OLTC_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE_3_STNR

        'End If
    End Sub

    Private Sub OLTC_STNR2_DIM1_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR2.SelectedItem = "1" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_1DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_2DIM1

        'ElseIf NO_OLTC_STNR2.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_3DIM

        'End If

    End Sub

    Private Sub OLTC_STNR2_DIM2_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR2.SelectedItem = "1" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_1DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_2DIM1

        'ElseIf NO_OLTC_STNR2.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_3DIM

        'End If

    End Sub

    Private Sub OLTC_STNR2_DIM3_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_STNR2.SelectedItem = "1" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_1DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "2" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_2DIM1

        'ElseIf NO_OLTC_STNR2.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTCSIDE2_3DIM

        'End If

    End Sub

    Private Sub OLTC_END_DIM1_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        'End If

    End Sub

    Private Sub OLTC_END_DIM2_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        'End If

    End Sub

    Private Sub OLTC_END_DIM3_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        'End If

    End Sub

    Private Sub OLTC_END_DIM4_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        'End If

    End Sub

    Private Sub OLTC_END_DIM5_Click(sender As Object, e As EventArgs)
        'If NO_OLTC_END_STNR.SelectedItem = "3" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_3DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "4" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_4DIM

        'ElseIf NO_OLTC_STNR2.SelectedItem = "5" Then
        '    STIFNER_PIC_BOX.Image = My.Resources.OLTC_END_5DIM

        'End If

    End Sub

    Private Sub NO_OF_JP_HV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NO_OF_JP_HV.SelectedIndexChanged

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If NO_OF_JP_HV.SelectedItem = "2" Then

            HV_JPD1.Enabled = True
            HV_JPD2.Enabled = True
            HV_JPD3.Enabled = False
            HV_JPD4.Enabled = False
            JP_Y_DIST.Enabled = False

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then
            HV_JPD1.Enabled = True
            HV_JPD2.Enabled = True
            HV_JPD3.Enabled = True
            HV_JPD4.Enabled = False
            JP_Y_DIST.Enabled = False

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then
            HV_JPD1.Enabled = True
            HV_JPD2.Enabled = True
            HV_JPD3.Enabled = True
            HV_JPD4.Enabled = True
            JP_Y_DIST.Enabled = False

        End If

        If NO_OF_JP_HV.SelectedItem = "2" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_2

            jp_hv_x1.Location = New System.Drawing.Point(228, 452)
            jp_hv_x2.Location = New System.Drawing.Point(348, 452)
            jp_h_500.Location = New System.Drawing.Point(224, 511)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_3

            jp_hv_x1.Location = New System.Drawing.Point(238, 468)
            jp_hv_x2.Location = New System.Drawing.Point(357, 467)
            jp_hv_x3.Location = New System.Drawing.Point(304, 572)
            jp_h_500.Location = New System.Drawing.Point(232, 529)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Show()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_4

            jp_hv_x1.Location = New System.Drawing.Point(167, 449)
            jp_hv_x2.Location = New System.Drawing.Point(406, 449)
            jp_hv_x3.Location = New System.Drawing.Point(236, 481)
            jp_hv_x4.Location = New System.Drawing.Point(341, 481)
            jp_h_500.Location = New System.Drawing.Point(601, 542)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_hv_x3.Show()
            jp_hv_x4.Show()
            jp_h_500.Show()

            JP_Y_DIST.Enabled = False
        End If


    End Sub

    Private Sub NO_JP_LV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NO_JP_LV.SelectedIndexChanged

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If NO_JP_LV.SelectedItem = "2" Then

            LV_JP1.Enabled = True
            LV_JP2.Enabled = True
            LV_JP3.Enabled = False
            LV_JP4.Enabled = False
            JP_Y_DIST_LV.Enabled = False

        ElseIf NO_JP_LV.SelectedItem = "3" Then
            LV_JP1.Enabled = True
            LV_JP2.Enabled = True
            LV_JP3.Enabled = True
            LV_JP4.Enabled = False
            JP_Y_DIST_LV.Enabled = False

        ElseIf NO_JP_LV.SelectedItem = "4" Then
            LV_JP1.Enabled = True
            LV_JP2.Enabled = True
            LV_JP3.Enabled = True
            LV_JP4.Enabled = True
            JP_Y_DIST_LV.Enabled = False
        End If

        If NO_JP_LV.SelectedItem = "2" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_2

            jp_lv_x1.Location = New System.Drawing.Point(228, 452)
            jp_lv_x2.Location = New System.Drawing.Point(348, 452)
            jp_l_500.Location = New System.Drawing.Point(224, 511)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

        ElseIf NO_JP_LV.SelectedItem = "3" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_3

            jp_lv_x1.Location = New System.Drawing.Point(238, 468)
            jp_lv_x2.Location = New System.Drawing.Point(357, 467)
            jp_lv_x3.Location = New System.Drawing.Point(304, 572)
            jp_l_500.Location = New System.Drawing.Point(232, 529)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Show()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

        ElseIf NO_JP_LV.SelectedItem = "4" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_4

            jp_lv_x1.Location = New System.Drawing.Point(167, 449)
            jp_lv_x2.Location = New System.Drawing.Point(406, 449)
            jp_lv_x3.Location = New System.Drawing.Point(236, 481)
            jp_lv_x4.Location = New System.Drawing.Point(341, 481)
            jp_l_500.Location = New System.Drawing.Point(601, 542)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_lv_x3.Show()
            jp_lv_x4.Show()
            jp_l_500.Show()

            JP_Y_DIST_LV.Enabled = False
        End If


    End Sub

    Private Sub HV_STNR_TYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles HV_STNR_TYPE.SelectedIndexChanged




        'If HV_STNR_TYPE.SelectedItem = "Type 1" Then

        '    HV_PI_BOX_TOP.Image = My.Resources.ROLLER_GUSSET_T1L1


        '    C.Hide()
        '    D.Hide()
        '    HV_E.Hide()
        '    F.Hide()

        '    F_STNR_C.Hide()
        '    F_STNR_D.Hide()
        '    F_STNR_E.Hide()
        '    F_STNR_F.Hide()
        '    Label102.Show()
        '    F_STNR_R2.Show()


        '    Label150.Show()
        '    Label170.Show()
        '    Label152.Show()
        '    CURB_HOLECUT.Show()
        '    CURB_HOLE_W.Show()
        '    CURB_HOLE_H.Show()

        '    Label101.Text = "R1"
        'Else

        '    HV_PI_BOX_TOP.Image = My.Resources.HV_SIDETYPE2

        '    Label101.Text = "R"
        '    C.Show()
        '    D.Show()
        '    HV_E.Show()
        '    F.Show()

        '    F_STNR_C.Show()
        '    F_STNR_D.Show()
        '    F_STNR_E.Show()
        '    F_STNR_F.Show()
        '    Label102.Hide()
        '    F_STNR_R2.Hide()




        '    Label150.Hide()
        '    Label170.Hide()
        '    Label152.Hide()
        '    CURB_HOLECUT.Hide()
        '    CURB_HOLE_W.Hide()
        '    CURB_HOLE_H.Hide()

        'End If

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then





            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If


    End Sub

    Private Sub LV_STNR_TYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LV_STNR_TYPE.SelectedIndexChanged

        'If LV_STNR_TYPE.SelectedItem = "Type 1" Then

        '    LV_PI_BOX_TOP.Image = My.Resources.ROLLER_GUSSET_T1R1


        '    Label172.Text = "R1"

        '    LV_C.Hide()
        '    LV_D.Hide()
        '    LV_E.Hide()
        '    LV_F.Hide()

        '    LV_F_STNR_C.Hide()
        '    LV_F_STNR_D.Hide()
        '    LV_F_STNR_E.Hide()
        '    LV_F_STNR_F.Hide()
        '    Label171.Show()
        '    LV_F_STNR_R2.Show()

        '    Label87.Show()
        '    CURB_HOLECUT_LV.Show()
        '    Label80.Show()
        '    LV_CURB_HOLE_W.Show()
        '    Label81.Show()
        '    LV_CURB_HOLE_H.Show()


        'Else

        '    LV_PI_BOX_TOP.Image = My.Resources.HV_SIDETYPE2R

        '    Label172.Text = "R"

        '    LV_C.Show()
        '    LV_D.Show()
        '    LV_E.Show()
        '    LV_F.Show()

        '    LV_F_STNR_C.Show()
        '    LV_F_STNR_D.Show()
        '    LV_F_STNR_E.Show()
        '    LV_F_STNR_F.Show()
        '    Label171.Hide()
        '    LV_F_STNR_R2.Hide()



        '    Label87.Hide()
        '    CURB_HOLECUT_LV.Hide()
        '    Label80.Hide()
        '    LV_CURB_HOLE_W.Hide()
        '    Label81.Hide()
        '    LV_CURB_HOLE_H.Hide()


        'End If

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then




            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()



            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If

    End Sub





    Private Sub GAUDE_PIN_SIZE_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub GAUDE_PIN_POS_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub X1_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub Y1_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub X2_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub Y2_Click(sender As Object, e As EventArgs)
        OP1_PIC_BOX.Image = My.Resources.GUIDINGPIN
    End Sub

    Private Sub SAM_VAL_POS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SAM_VAL_POS.SelectedIndexChanged

        Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

        If SAM_VAL_POS.SelectedItem = "LV" Then
            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1


        ElseIf SAM_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1


        ElseIf SAM_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2


        ElseIf SAM_VAL_POS.SelectedItem = "OLTC" Then
            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE

        End If


    End Sub

    Private Sub FILTER_VAL_POS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FILTER_VAL_POS.SelectedIndexChanged

        Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

        If FILTER_VAL_POS.SelectedItem = "LV" Then
            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1

        ElseIf FILTER_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC" Then
            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE
        End If

    End Sub

    Private Sub DRAIN_VALVE_POS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DRAIN_VALVE_POS.SelectedIndexChanged

        Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

        If DRAIN_VALVE_POS.SelectedItem = "LV" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_LV

        ElseIf DRAIN_VALVE_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_HV

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_End

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_SIDE
        End If

    End Sub

    Private Sub DGA_VALVE_POS_SelectedIndexChanged(sender As Object, e As EventArgs)

        Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()



    End Sub

    Public Sub BT_Bottom_Tank()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Step 1: Check if wk_odr_No exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM Bottom_Tank_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    ' Common parameter list for both insert and update
                    Dim parameters As SqlParameterCollection = New SqlCommand().Parameters
                    Dim paramValues As Action(Of SqlCommand) = Sub(cmd)
                                                                   cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                                                   cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                                                                   cmd.Parameters.AddWithValue("@btm_tk_type", btm_tk_type.SelectedItem)
                                                                   cmd.Parameters.AddWithValue("@btm_tk_curb_pro", btm_tk_curb_pro.SelectedItem)
                                                                   cmd.Parameters.AddWithValue("@curb_size", curb_size.SelectedItem)
                                                                   cmd.Parameters.AddWithValue("@btm_plt_ht", btm_plt_ht.SelectedItem)
                                                                   cmd.Parameters.AddWithValue("@Cus_ht", Cus_ht.Text)
                                                                   cmd.Parameters.AddWithValue("@btm_tk_i_len", btm_tk_i_len.Text)
                                                                   cmd.Parameters.AddWithValue("@btm_plt_len", btm_plt_len.Text)
                                                                   cmd.Parameters.AddWithValue("@btm_plt_thk", btm_plt_thk.Text)
                                                                   cmd.Parameters.AddWithValue("@btm_plt_width", btm_plt_width.Text)
                                                                   cmd.Parameters.AddWithValue("@Side_Plt_thk", Side_Plt_thk.Text)
                                                                   cmd.Parameters.AddWithValue("@btm_tk_i_wdh", btm_tk_i_wdh.Text)
                                                                   cmd.Parameters.AddWithValue("@OLTC_pck_len", OLTC_pck_len.Text)
                                                                   cmd.Parameters.AddWithValue("@OLTC_pck_wdh", OLTC_pck_wdh.Text)
                                                                   cmd.Parameters.AddWithValue("@tk_curb_wdh", tk_curb_wdh.Text)
                                                                   cmd.Parameters.AddWithValue("@Curb_wdh_LV", Curb_wdh_LV.Text)
                                                                   cmd.Parameters.AddWithValue("@Curb_OLTC_End", Curb_OLTC_End.Text)
                                                                   cmd.Parameters.AddWithValue("@Curb_wdh_D", Curb_wdh_D.Text)
                                                                   cmd.Parameters.AddWithValue("@GAUDE_PIN_SIZE", GAUDE_PIN_SIZE.SelectedItem)
                                                                   cmd.Parameters.AddWithValue("@GPX1", GPX1.Text)
                                                                   cmd.Parameters.AddWithValue("@GPX2", GPX2.Text)
                                                                   cmd.Parameters.AddWithValue("@GPY1", GPY1.Text)
                                                                   cmd.Parameters.AddWithValue("@GPY2", GPY2.Text)

                                                               End Sub

                    If count > 0 Then
                        ' 🔁 Update existing record
                        Dim updateQuery As String = "UPDATE Bottom_Tank_BT SET 
                        Username=@Username,
                        btm_tk_type=@btm_tk_type,
                        btm_tk_curb_pro=@btm_tk_curb_pro,
                        curb_size=@curb_size,
                        btm_plt_ht=@btm_plt_ht,
                        Cus_ht=@Cus_ht,
                        btm_tk_i_len=@btm_tk_i_len,
                        btm_plt_len=@btm_plt_len,
                        btm_plt_thk=@btm_plt_thk,
                        btm_plt_width=@btm_plt_width,
                        Side_Plt_thk=@Side_Plt_thk,
                        btm_tk_i_wdh=@btm_tk_i_wdh,
                        OLTC_pck_len=@OLTC_pck_len,
                        OLTC_pck_wdh=@OLTC_pck_wdh,
                        tk_curb_wdh=@tk_curb_wdh,
                        Curb_wdh_LV=@Curb_wdh_LV,
                        Curb_OLTC_End=@Curb_OLTC_End,
                        Curb_wdh_D=@Curb_wdh_D,
                        GAUDE_PIN_SIZE=@GAUDE_PIN_SIZE,
                        GPX1=@GPX1,
                        GPX2=@GPX2,
                        GPY1=@GPY1,
                        GPY2=@GPY2
                        WHERE wk_odr_No=@wk_odr_No"

                        Using updateCmd As New SqlCommand(updateQuery, connection)
                            paramValues(updateCmd)
                            updateCmd.ExecuteNonQuery()
                            MsgBox("Data updated successfully!")
                        End Using


                    End If
                End Using
            End Using



        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try


        'Try

        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO Bottom_Tank_BT(Username,wk_odr_No,btm_tk_type,btm_tk_curb_pro,OLTC_pkt_position,curb_size,TT_curb_size,btm_plt_ht,Cus_ht,
        '    btm_tk_i_len,btm_plt_len,btm_plt_thk,btm_plt_width,Side_Plt_thk,btm_tk_i_wdh,bt_tk_wdh_2,OLTC_pck_len,OLTC_pck_wdh,tk_curb_wdh,Curb_wdh_LV,Curb_OLTC_End,
        '    Curb_wdh_D,BP_X1,BP_X2,BP_X3,BP_R1,GAUDE_PIN_SIZE,GPX1,GPX2,GPY1,GPY2,BTM_TK_Lifting_lugs) 
        '    values (@Username,@wk_odr_No,@btm_tk_type,@btm_tk_curb_pro,@OLTC_pkt_position,@curb_size,@TT_curb_size,@btm_plt_ht,@Cus_ht,
        '    @btm_tk_i_len,@btm_plt_len,@btm_plt_thk,@btm_plt_width,@Side_Plt_thk,@btm_tk_i_wdh,@bt_tk_wdh_2,@OLTC_pck_len,@OLTC_pck_wdh,@tk_curb_wdh,@Curb_wdh_LV,@Curb_OLTC_End,
        '    @Curb_wdh_D,@BP_X1,@BP_X2,@BP_X3,@BP_R1,@GAUDE_PIN_SIZE,@GPX1,@GPX2,@GPY1,@GPY2,@BTM_TK_Lifting_lugs)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@btm_tk_type", btm_tk_type.SelectedItem)
        '    cmd.Parameters.AddWithValue("@btm_tk_curb_pro", btm_tk_curb_pro.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_pkt_position", OLTC_pkt_position)
        '    cmd.Parameters.AddWithValue("@curb_size", curb_size.SelectedItem)
        '    cmd.Parameters.AddWithValue("@TT_curb_size", TT_curb_size.SelectedItem)
        '    cmd.Parameters.AddWithValue("@btm_plt_ht", btm_plt_ht.SelectedItem)

        '    cmd.Parameters.AddWithValue("@Cus_ht", Cus_ht.Text)
        '    cmd.Parameters.AddWithValue("@btm_tk_i_len", btm_tk_i_len.Text)
        '    cmd.Parameters.AddWithValue("@btm_plt_len", btm_plt_len.Text)
        '    cmd.Parameters.AddWithValue("@btm_plt_thk", btm_plt_thk.Text)
        '    cmd.Parameters.AddWithValue("@btm_plt_width", btm_plt_width.Text)
        '    cmd.Parameters.AddWithValue("@Side_Plt_thk", Side_Plt_thk.Text)
        '    cmd.Parameters.AddWithValue("@btm_tk_i_wdh", btm_tk_i_wdh.Text)
        '    cmd.Parameters.AddWithValue("@bt_tk_wdh_2", bt_tk_wdh_2.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_pck_len", OLTC_pck_len.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_pck_wdh", OLTC_pck_wdh.Text)
        '    cmd.Parameters.AddWithValue("@tk_curb_wdh", tk_curb_wdh.Text)
        '    cmd.Parameters.AddWithValue("@Curb_wdh_LV", Curb_wdh_LV.Text)
        '    cmd.Parameters.AddWithValue("@Curb_OLTC_End", Curb_OLTC_End.Text)
        '    cmd.Parameters.AddWithValue("@Curb_wdh_D", Curb_wdh_D.Text)
        '    cmd.Parameters.AddWithValue("@BP_X1", BP_X1.Text)
        '    cmd.Parameters.AddWithValue("@BP_X2", BP_X2.Text)
        '    cmd.Parameters.AddWithValue("@BP_X3", BP_X3.Text)
        '    cmd.Parameters.AddWithValue("@BP_R1", BP_R1.Text)

        '    cmd.Parameters.AddWithValue("@GAUDE_PIN_SIZE", GAUDE_PIN_SIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@GPX1", GPX1.Text)
        '    cmd.Parameters.AddWithValue("@GPX2", GPX2.Text)
        '    cmd.Parameters.AddWithValue("@GPY1", GPY1.Text)
        '    cmd.Parameters.AddWithValue("@GPY2", GPY2.Text)
        '    cmd.Parameters.AddWithValue("@BTM_TK_Lifting_lugs", BTM_TK_Lifting_lugs.SelectedItem)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        'BT_Bottom_Tank()
        BT_Bottom_Tank()
    End Sub

    Public Sub BT_Underbase()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Step 1: Check if wk_odr_No exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM UnderBase_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    ' Define parameter-filling action
                    Dim fillParams As Action(Of SqlCommand) = Sub(cmd)
                                                                  cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                                                  cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                                                                  cmd.Parameters.AddWithValue("@No_of_Rails", No_of_Rails.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@No_of_Rails1", No_of_Rails1.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@CC_LENGTH", CC_LENGTH.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@Roller_Custom", Roller_Custom.Text)
                                                                  cmd.Parameters.AddWithValue("@roller_gauge_cmb", roller_gauge_cmb.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@roller_orientation", roller_orientation.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@BS1_WDH_1", BS1_WDH_1.Text)
                                                                  cmd.Parameters.AddWithValue("@BS1_HT_WS", BS1_HT_WS.Text)
                                                                  cmd.Parameters.AddWithValue("@BS1_THK_1", BS1_THK_1.Text)
                                                                  cmd.Parameters.AddWithValue("@BS2_WDH_1", BS2_WDH_1.Text)
                                                                  cmd.Parameters.AddWithValue("@BS2_HT_WS", BS2_HT_WS.Text)
                                                                  cmd.Parameters.AddWithValue("@BS2_THK_1", BS2_THK_1.Text)
                                                                  cmd.Parameters.AddWithValue("@BS3_WDH_1", BS3_WDH_1.Text)
                                                                  cmd.Parameters.AddWithValue("@BS3_HT_WS", BS3_HT_WS.Text)
                                                                  cmd.Parameters.AddWithValue("@BS3_THK_1", BS3_THK_1.Text)
                                                                  cmd.Parameters.AddWithValue("@HS_CC_len", HS_CC_len.Text)
                                                                  cmd.Parameters.AddWithValue("@X", X.Text)
                                                                  cmd.Parameters.AddWithValue("@Distance_A", Distance_A.Text)
                                                                  cmd.Parameters.AddWithValue("@Distance_B", Distance_B.Text)
                                                                  cmd.Parameters.AddWithValue("@MP_len", MP_len.Text)
                                                                  cmd.Parameters.AddWithValue("@MP_Wth", MP_Wth.Text)
                                                                  cmd.Parameters.AddWithValue("@MP_Thk", MP_Thk.Text)
                                                                  cmd.Parameters.AddWithValue("@T2", T2.Text)
                                                                  cmd.Parameters.AddWithValue("@Dim_L1", Dim_L1.Text)
                                                                  cmd.Parameters.AddWithValue("@prob_support_cmb", prob_support_cmb.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@prob_no", prob_no.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@PROB_HT", PROB_HT.Text)
                                                                  cmd.Parameters.AddWithValue("@Prob_X1axis", Prob_X1axis.Text)
                                                                  cmd.Parameters.AddWithValue("@Prob_X2axis", Prob_X2axis.Text)
                                                                  cmd.Parameters.AddWithValue("@Prob_Y1axis", Prob_Y1axis.Text)
                                                                  cmd.Parameters.AddWithValue("@Prob_Y2axis", Prob_Y2axis.Text)
                                                                  cmd.Parameters.AddWithValue("@BS_L", BS_L.Text)
                                                                  cmd.Parameters.AddWithValue("@BS_R", BS_R.Text)
                                                              End Sub

                    If count > 0 Then
                        ' 🔁 Update existing row
                        Dim updateQuery As String = "UPDATE UnderBase_BT SET
                    Username=@Username,
                    No_of_Rails=@No_of_Rails,
                    No_of_Rails1=@No_of_Rails1,
                    CC_LENGTH=@CC_LENGTH,
                    Roller_Custom=@Roller_Custom,
                    roller_gauge_cmb=@roller_gauge_cmb,
                    roller_orientation=@roller_orientation,
                    BS1_WDH_1=@BS1_WDH_1,
                    BS1_HT_WS=@BS1_HT_WS,
                    BS1_THK_1=@BS1_THK_1,
                    BS2_WDH_1=@BS2_WDH_1,
                    BS2_HT_WS=@BS2_HT_WS,
                    BS2_THK_1=@BS2_THK_1,
                    BS3_WDH_1=@BS3_WDH_1,
                    BS3_HT_WS=@BS3_HT_WS,
                    BS3_THK_1=@BS3_THK_1,
                    HS_CC_len=@HS_CC_len,
                    X=@X,
                    Distance_A=@Distance_A,
                    Distance_B=@Distance_B,
                    MP_len=@MP_len,
                    MP_Wth=@MP_Wth,
                    MP_Thk=@MP_Thk,
                    T2=@T2,
                    Dim_L1=@Dim_L1,
                    prob_support_cmb=@prob_support_cmb,
                    prob_no=@prob_no,
                    PROB_HT=@PROB_HT,
                    Prob_X1axis=@Prob_X1axis,
                    Prob_X2axis=@Prob_X2axis,
                    Prob_Y1axis=@Prob_Y1axis,
                    Prob_Y2axis=@Prob_Y2axis,
                    BS_L=@BS_L,
                    BS_R=@BS_R
                WHERE wk_odr_No=@wk_odr_No"

                        Using updateCmd As New SqlCommand(updateQuery, connection)
                            fillParams(updateCmd)
                            updateCmd.ExecuteNonQuery()
                            MsgBox("Data updated successfully!")
                        End Using
                    Else
                        '        ' ➕ Insert new row
                        '        Dim insertQuery As String = "INSERT INTO UnderBase_BT (
                        '    Username, wk_odr_No, No_of_Rails, No_of_Rails1, CC_LENGTH, Roller_Custom, roller_gauge_cmb, roller_orientation,
                        '    BS1_WDH_1, BS1_HT_WS, BS1_THK_1, BS2_WDH_1, BS2_HT_WS, BS2_THK_1, BS3_WDH_1, BS3_HT_WS, BS3_THK_1, HS_CC_len, X,
                        '    Distance_A, Distance_B, MP_len, MP_Wth, MP_Thk, T2, Dim_L1, prob_support_cmb, prob_no, PROB_HT,
                        '    Prob_X1axis, Prob_X2axis, Prob_Y1axis, Prob_Y2axis, BS_L, BS_R)
                        'VALUES (
                        '    @Username, @wk_odr_No, @No_of_Rails, @No_of_Rails1, @CC_LENGTH, @Roller_Custom, @roller_gauge_cmb, @roller_orientation,
                        '    @BS1_WDH_1, @BS1_HT_WS, @BS1_THK_1, @BS2_WDH_1, @BS2_HT_WS, @BS2_THK_1, @BS3_WDH_1, @BS3_HT_WS, @BS3_THK_1, @HS_CC_len, @X,
                        '    @Distance_A, @Distance_B, @MP_len, @MP_Wth, @MP_Thk, @T2, @Dim_L1, @prob_support_cmb, @prob_no, @PROB_HT,
                        '    @Prob_X1axis, @Prob_X2axis, @Prob_Y1axis, @Prob_Y2axis, @BS_L, @BS_R)"

                        '        Using insertCmd As New SqlCommand(insertQuery, connection)
                        '            fillParams(insertCmd)
                        '            insertCmd.ExecuteNonQuery()
                        '            MsgBox("Data inserted successfully!")
                        '        End Using
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try


        'Try
        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO UnderBase_BT(Username,wk_odr_No,No_of_Rails,No_of_Rails1,CC_LENGTH,Roller_Custom,roller_gauge_cmb,roller_orientation,
        '    BS1_WDH_1,BS1_HT_WS,BS1_THK_1,BS2_WDH_1,BS2_HT_WS,BS2_THK_1,BS3_WDH_1,BS3_HT_WS,BS3_THK_1,HS_CC_len,X,Distance_A,Distance_B,MP_len,MP_Wth,MP_Thk,T2,
        '    Dim_L1,prob_support_cmb,prob_no,PROB_HT,Prob_X1axis,Prob_X2axis,Prob_Y1axis,Prob_Y2axis,BS_L,BS_R) 
        '    values (@Username,@wk_odr_No,@No_of_Rails,@No_of_Rails1,@CC_LENGTH,@Roller_Custom,@roller_gauge_cmb,@roller_orientation,
        '    @BS1_WDH_1,@BS1_HT_WS,@BS1_THK_1,@BS2_WDH_1,@BS2_HT_WS,@BS2_THK_1,@BS3_WDH_1,@BS3_HT_WS,@BS3_THK_1,@HS_CC_len,@X,@Distance_A,@Distance_B,@MP_len,@MP_Wth,@MP_Thk,@T2,
        '    @Dim_L1,@prob_support_cmb,@prob_no,@PROB_HT,@Prob_X1axis,@Prob_X2axis,@Prob_Y1axis,@Prob_Y2axis,@BS_L,@BS_R)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@No_of_Rails", No_of_Rails.SelectedItem)
        '    cmd.Parameters.AddWithValue("@No_of_Rails1", No_of_Rails1.SelectedItem)
        '    cmd.Parameters.AddWithValue("@CC_LENGTH", CC_LENGTH.SelectedItem)
        '    cmd.Parameters.AddWithValue("@Roller_Custom", Roller_Custom.Text)
        '    cmd.Parameters.AddWithValue("@roller_gauge_cmb", roller_gauge_cmb.SelectedItem)
        '    cmd.Parameters.AddWithValue("@roller_orientation", roller_orientation.SelectedItem)

        '    cmd.Parameters.AddWithValue("@BS1_WDH_1", BS1_WDH_1.Text)
        '    cmd.Parameters.AddWithValue("@BS1_HT_WS", BS1_HT_WS.Text)
        '    cmd.Parameters.AddWithValue("@BS1_THK_1", BS1_THK_1.Text)
        '    cmd.Parameters.AddWithValue("@BS2_WDH_1", BS2_WDH_1.Text)
        '    cmd.Parameters.AddWithValue("@BS2_HT_WS", BS2_HT_WS.Text)
        '    cmd.Parameters.AddWithValue("@BS2_THK_1", BS2_THK_1.Text)
        '    cmd.Parameters.AddWithValue("@BS3_WDH_1", BS3_WDH_1.Text)
        '    cmd.Parameters.AddWithValue("@BS3_HT_WS", BS3_HT_WS.Text)
        '    cmd.Parameters.AddWithValue("@BS3_THK_1", BS3_THK_1.Text)

        '    cmd.Parameters.AddWithValue("@HS_CC_len", HS_CC_len.Text)
        '    cmd.Parameters.AddWithValue("@X", X.Text)
        '    cmd.Parameters.AddWithValue("@Distance_A", Distance_A.Text)
        '    cmd.Parameters.AddWithValue("@Distance_B", Distance_B.Text)

        '    cmd.Parameters.AddWithValue("@MP_len", MP_len.Text)
        '    cmd.Parameters.AddWithValue("@MP_Wth", MP_Wth.Text)
        '    cmd.Parameters.AddWithValue("@MP_Thk", MP_Thk.Text)
        '    cmd.Parameters.AddWithValue("@T2", T2.Text)
        '    cmd.Parameters.AddWithValue("@Dim_L1", Dim_L1.Text)

        '    cmd.Parameters.AddWithValue("@prob_support_cmb", prob_support_cmb.SelectedItem)
        '    cmd.Parameters.AddWithValue("@prob_no", prob_no.SelectedItem)
        '    cmd.Parameters.AddWithValue("@PROB_HT", PROB_HT.Text)
        '    cmd.Parameters.AddWithValue("@Prob_X1axis", Prob_X1axis.Text)
        '    cmd.Parameters.AddWithValue("@Prob_X2axis", Prob_X2axis.Text)
        '    cmd.Parameters.AddWithValue("@Prob_Y1axis", Prob_Y1axis.Text)
        '    cmd.Parameters.AddWithValue("@Prob_Y2axis", Prob_Y2axis.Text)

        '    cmd.Parameters.AddWithValue("@BS_L", BS_L.Text)
        '    cmd.Parameters.AddWithValue("@BS_R", BS_R.Text)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub
    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click

        BT_Underbase()
    End Sub

    Public Sub BT_Jacking_HV()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Step 1: Check if wk_odr_No exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM JackingPad_HV_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    ' Step 2: Create shared parameter filler
                    Dim fillParams As Action(Of SqlCommand) = Sub(cmd)
                                                                  cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                                                  cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                                                                  cmd.Parameters.AddWithValue("@jack_pad_type", jack_pad_type.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@JP_A", JP_A.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_B", JP_B.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_T", JP_T.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_G", JP_G.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_E", JP_E.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_D", JP_D.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_C", JP_C.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_B1", JP_B1.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_BB1", JP_BB1.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_B2", JP_B2.Text)
                                                                  cmd.Parameters.AddWithValue("@NO_TOP_JP_HV", NO_TOP_JP_HV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@JP_T_L", JP_T_L.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_T_H", JP_T_H.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_T_T", JP_T_T.Text)
                                                                  cmd.Parameters.AddWithValue("@NO_OF_JP_HV", NO_OF_JP_HV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@HV_JPD1", HV_JPD1.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_JPD2", HV_JPD2.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_JPD3", HV_JPD3.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_JPD4", HV_JPD4.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_Y_DIST", JP_Y_DIST.Text)
                                                              End Sub

                    If count > 0 Then
                        ' 🔁 Update existing record
                        Dim updateQuery As String = "UPDATE JackingPad_HV_BT SET
                    Username=@Username,
                    jack_pad_type=@jack_pad_type,
                    JP_A=@JP_A,
                    JP_B=@JP_B,
                    JP_T=@JP_T,
                    JP_G=@JP_G,
                    JP_E=@JP_E,
                    JP_D=@JP_D,
                    JP_C=@JP_C,
                    JP_B1=@JP_B1,
                    JP_BB1=@JP_BB1,
                    JP_B2=@JP_B2,
                    NO_TOP_JP_HV=@NO_TOP_JP_HV,
                    JP_T_L=@JP_T_L,
                    JP_T_H=@JP_T_H,
                    JP_T_T=@JP_T_T,
                    NO_OF_JP_HV=@NO_OF_JP_HV,
                    HV_JPD1=@HV_JPD1,
                    HV_JPD2=@HV_JPD2,
                    HV_JPD3=@HV_JPD3,
                    HV_JPD4=@HV_JPD4,
                    JP_Y_DIST=@JP_Y_DIST
                WHERE wk_odr_No=@wk_odr_No"

                        Using updateCmd As New SqlCommand(updateQuery, connection)
                            fillParams(updateCmd)
                            updateCmd.ExecuteNonQuery()
                            MsgBox("Data updated successfully!")
                        End Using
                    Else
                        '        ' ➕ Insert new row
                        '        Dim insertQuery As String = "INSERT INTO JackingPad_HV_BT (
                        '    Username, wk_odr_No, jack_pad_type, JP_A, JP_B, JP_T, JP_G, JP_E, JP_D, JP_X, JP_C, JP_B1, JP_BB1, JP_B2,
                        '    NO_TOP_JP_HV, JP_T_L, JP_T_H, JP_T_T, NO_OF_JP_HV, HV_JPD1, HV_JPD2, HV_JPD3, HV_JPD4, JP_Y_DIST
                        ') VALUES (
                        '    @Username, @wk_odr_No, @jack_pad_type, @JP_A, @JP_B, @JP_T, @JP_G, @JP_E, @JP_D, @JP_X, @JP_C, @JP_B1, @JP_BB1, @JP_B2,
                        '    @NO_TOP_JP_HV, @JP_T_L, @JP_T_H, @JP_T_T, @NO_OF_JP_HV, @HV_JPD1, @HV_JPD2, @HV_JPD3, @HV_JPD4, @JP_Y_DIST
                        ')"

                        '        Using insertCmd As New SqlCommand(insertQuery, connection)
                        '            fillParams(insertCmd)
                        '            insertCmd.ExecuteNonQuery()
                        '            MsgBox("Data inserted successfully!")
                        '        End Using
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try


        'Try
        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO JackingPad_HV_BT(Username,wk_odr_No,jack_pad_type,JP_A,JP_B,JP_T,JP_G,JP_E,JP_D,JP_X,JP_C,JP_B1,JP_BB1,JP_B2,
        '    NO_TOP_JP_HV,JP_T_L,JP_T_H,JP_T_T,NO_OF_JP_HV,HV_JPD1,HV_JPD2,HV_JPD3,HV_JPD4,JP_Y_DIST) 
        '    values (@Username,@wk_odr_No,@jack_pad_type,@JP_A,@JP_B,@JP_T,@JP_G,@JP_E,@JP_D,@JP_X,@JP_C,@JP_B1,@JP_BB1,@JP_B2,
        '    @NO_TOP_JP_HV,@JP_T_L,@JP_T_H,@JP_T_T,@NO_OF_JP_HV,@HV_JPD1,@HV_JPD2,@HV_JPD3,@HV_JPD4,@JP_Y_DIST)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@jack_pad_type", jack_pad_type.SelectedItem)
        '    cmd.Parameters.AddWithValue("@JP_A", JP_A.Text)
        '    cmd.Parameters.AddWithValue("@JP_B", JP_B.Text)
        '    cmd.Parameters.AddWithValue("@JP_T", JP_T.Text)
        '    cmd.Parameters.AddWithValue("@JP_G", JP_G.Text)
        '    cmd.Parameters.AddWithValue("@JP_E", JP_E.Text)
        '    cmd.Parameters.AddWithValue("@JP_D", JP_D.Text)
        '    cmd.Parameters.AddWithValue("@JP_X", JP_X.Text)
        '    cmd.Parameters.AddWithValue("@JP_C", JP_C.Text)
        '    cmd.Parameters.AddWithValue("@JP_B1", JP_B1.Text)
        '    cmd.Parameters.AddWithValue("@JP_BB1", JP_BB1.Text)
        '    cmd.Parameters.AddWithValue("@JP_B2", JP_B2.Text)

        '    cmd.Parameters.AddWithValue("@NO_TOP_JP_HV", NO_TOP_JP_HV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@JP_T_L", JP_T_L.Text)
        '    cmd.Parameters.AddWithValue("@JP_T_H", JP_T_H.Text)
        '    cmd.Parameters.AddWithValue("@JP_T_T", JP_T_T.Text)

        '    cmd.Parameters.AddWithValue("@NO_OF_JP_HV", NO_OF_JP_HV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@HV_JPD1", HV_JPD1.Text)
        '    cmd.Parameters.AddWithValue("@HV_JPD2", HV_JPD2.Text)
        '    cmd.Parameters.AddWithValue("@HV_JPD3", HV_JPD3.Text)
        '    cmd.Parameters.AddWithValue("@HV_JPD4", HV_JPD4.Text)
        '    cmd.Parameters.AddWithValue("@JP_Y_DIST", JP_Y_DIST.Text)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click

        BT_Jacking_HV()

    End Sub

    Public Sub BT_Jacking_LV()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Step 1: Check if wk_odr_No exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM JakingPad_LV_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    ' Step 2: Shared parameter filler
                    Dim fillParams As Action(Of SqlCommand) = Sub(cmd)
                                                                  cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                                                  cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                                                                  cmd.Parameters.AddWithValue("@jack_pad_type1", jack_pad_type1.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_A", JP_LV_A.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_B", JP_LV_B.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_C", JP_LV_C.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_D", JP_LV_D.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_E", JP_LV_E.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_G", JP_LV_G.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_T", JP_LV_T.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_B1", JP_LV_B1.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_BB1", JP_LV_BB1.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_LV_B2", JP_LV_B2.Text)
                                                                  cmd.Parameters.AddWithValue("@NO_TOP_JP_LV", NO_TOP_JP_LV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@JP_T_L_LV", JP_T_L_LV.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_T_T_LV", JP_T_T_LV.Text)
                                                                  cmd.Parameters.AddWithValue("@JP_T_H_LV", JP_T_H_LV.Text)
                                                                  cmd.Parameters.AddWithValue("@NO_JP_LV", NO_JP_LV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@JP_Y_DIST_LV", JP_Y_DIST_LV.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_JP1", LV_JP1.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_JP2", LV_JP2.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_JP3", LV_JP3.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_JP4", LV_JP4.Text)
                                                              End Sub

                    If count > 0 Then
                        ' 🔁 Update existing row
                        Dim updateQuery As String = "UPDATE JakingPad_LV_BT SET
                    Username=@Username,
                    jack_pad_type1=@jack_pad_type1,
                    JP_LV_A=@JP_LV_A,
                    JP_LV_B=@JP_LV_B,
                    JP_LV_C=@JP_LV_C,
                    JP_LV_D=@JP_LV_D,
                    JP_LV_E=@JP_LV_E,
                    JP_LV_G=@JP_LV_G,
                    JP_LV_T=@JP_LV_T,
                    JP_LV_B1=@JP_LV_B1,
                    JP_LV_BB1=@JP_LV_BB1,
                    JP_LV_B2=@JP_LV_B2,
                    NO_TOP_JP_LV=@NO_TOP_JP_LV,
                    JP_T_L_LV=@JP_T_L_LV,
                    JP_T_T_LV=@JP_T_T_LV,
                    JP_T_H_LV=@JP_T_H_LV,
                    NO_JP_LV=@NO_JP_LV,
                    JP_Y_DIST_LV=@JP_Y_DIST_LV,
                    LV_JP1=@LV_JP1,
                    LV_JP2=@LV_JP2,
                    LV_JP3=@LV_JP3,
                    LV_JP4=@LV_JP4
                WHERE wk_odr_No=@wk_odr_No"

                        Using updateCmd As New SqlCommand(updateQuery, connection)
                            fillParams(updateCmd)
                            updateCmd.ExecuteNonQuery()
                            MsgBox("Data updated successfully!")
                        End Using
                    Else
                        '        ' ➕ Insert new row
                        '        Dim insertQuery As String = "INSERT INTO JakingPad_LV_BT (
                        '    Username, wk_odr_No, jack_pad_type1, JP_LV_A, JP_LV_B, JP_LV_C, JP_LV_D, JP_LV_E, JP_LV_G, JP_LV_X, JP_LV_T,
                        '    JP_LV_B1, JP_LV_BB1, JP_LV_B2, NO_TOP_JP_LV, JP_T_L_LV, JP_T_T_LV, JP_T_H_LV, NO_JP_LV, JP_Y_DIST_LV,
                        '    LV_JP1, LV_JP2, LV_JP3, LV_JP4
                        ') VALUES (
                        '    @Username, @wk_odr_No, @jack_pad_type1, @JP_LV_A, @JP_LV_B, @JP_LV_C, @JP_LV_D, @JP_LV_E, @JP_LV_G, @JP_LV_X, @JP_LV_T,
                        '    @JP_LV_B1, @JP_LV_BB1, @JP_LV_B2, @NO_TOP_JP_LV, @JP_T_L_LV, @JP_T_T_LV, @JP_T_H_LV, @NO_JP_LV, @JP_Y_DIST_LV,
                        '    @LV_JP1, @LV_JP2, @LV_JP3, @LV_JP4
                        ')"

                        '        Using insertCmd As New SqlCommand(insertQuery, connection)
                        '            fillParams(insertCmd)
                        '            insertCmd.ExecuteNonQuery()
                        '            MsgBox("Data inserted successfully!")
                        '        End Using
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try

        'Try
        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO JakingPad_LV_BT(Username,wk_odr_No,jack_pad_type1,JP_LV_A,JP_LV_B,JP_LV_C,JP_LV_D,JP_LV_E,JP_LV_G,JP_LV_X,JP_LV_T,
        '    JP_LV_B1,JP_LV_BB1,JP_LV_B2,NO_TOP_JP_LV,JP_T_L_LV,JP_T_T_LV,JP_T_H_LV,NO_JP_LV,JP_Y_DIST_LV,LV_JP1,LV_JP2,LV_JP3,LV_JP4) 
        '    values (@Username,@wk_odr_No,@jack_pad_type1,@JP_LV_A,@JP_LV_B,@JP_LV_C,@JP_LV_D,@JP_LV_E,@JP_LV_G,@JP_LV_X,@JP_LV_T,
        '    @JP_LV_B1,@JP_LV_BB1,@JP_LV_B2,@NO_TOP_JP_LV,@JP_T_L_LV,@JP_T_T_LV,@JP_T_H_LV,@NO_JP_LV,@JP_Y_DIST_LV,@LV_JP1,@LV_JP2,@LV_JP3,@LV_JP4)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@jack_pad_type1", jack_pad_type1.SelectedItem)
        '    cmd.Parameters.AddWithValue("@JP_LV_A", JP_LV_A.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_B", JP_LV_B.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_C", JP_LV_C.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_D", JP_LV_D.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_E", JP_LV_E.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_G", JP_LV_G.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_X", JP_LV_X.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_T", JP_LV_T.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_B1", JP_LV_B1.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_BB1", JP_LV_BB1.Text)
        '    cmd.Parameters.AddWithValue("@JP_LV_B2", JP_LV_B2.Text)

        '    cmd.Parameters.AddWithValue("@NO_TOP_JP_LV", NO_TOP_JP_LV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@JP_T_L_LV", JP_T_L_LV.Text)
        '    cmd.Parameters.AddWithValue("@JP_T_T_LV", JP_T_T_LV.Text)
        '    cmd.Parameters.AddWithValue("@JP_T_H_LV", JP_T_H_LV.Text)

        '    cmd.Parameters.AddWithValue("@NO_JP_LV", NO_JP_LV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@JP_Y_DIST_LV", JP_Y_DIST_LV.Text)
        '    cmd.Parameters.AddWithValue("@LV_JP1", LV_JP1.Text)
        '    cmd.Parameters.AddWithValue("@LV_JP2", LV_JP2.Text)
        '    cmd.Parameters.AddWithValue("@LV_JP3", LV_JP3.Text)
        '    cmd.Parameters.AddWithValue("@LV_JP4", LV_JP4.Text)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        BT_Jacking_LV()
    End Sub


    Public Sub BT_Flat_Stiffner()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Step 1: Check if wk_odr_No exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM FlatStiffner_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    ' Step 2: Define shared parameters
                    Dim fillParams As Action(Of SqlCommand) = Sub(cmd)
                                                                  cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                                                  cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

                                                                  cmd.Parameters.AddWithValue("@HV_STNR_TYPE", HV_STNR_TYPE.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@HV_FS_A", HV_FS_A.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_FS_A1", HV_FS_A1.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_FS_B", HV_FS_B.Text)
                                                                  cmd.Parameters.AddWithValue("@HV_FS_B1", HV_FS_B1.Text)

                                                                  cmd.Parameters.AddWithValue("@LV_STNR_TYPE", LV_STNR_TYPE.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@LV_FS_A", LV_FS_A.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_FS_A1", LV_FS_A1.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_FS_B", LV_FS_B.Text)
                                                                  cmd.Parameters.AddWithValue("@LV_FS_B1", LV_FS_B1.Text)

                                                                  cmd.Parameters.AddWithValue("@OLTC_STNR_TYPE", OLTC_STNR_TYPE.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S1_A", OLTC_S1_A.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S1_A1", OLTC_S1_A1.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S1_B", OLTC_S1_B.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S1_B1", OLTC_S1_B1.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S1_T", OLTC_S1_T.Text)
                                                                  cmd.Parameters.AddWithValue("@OS1D1", OS1D1.Text)
                                                                  cmd.Parameters.AddWithValue("@OS1D2", OS1D2.Text)

                                                                  cmd.Parameters.AddWithValue("@OLTC_STNR2_TYPE", OLTC_STNR2_TYPE.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S2_A", OLTC_S2_A.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S2_B", OLTC_S2_B.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_S2_A1", OLTC_S2_A1.Text)
                                                                  cmd.Parameters.AddWithValue("@OS2D1", OS2D1.Text)
                                                                  cmd.Parameters.AddWithValue("@OS2D2", OS2D2.Text)

                                                                  cmd.Parameters.AddWithValue("@OLTC_END_STNR_TYP", OLTC_END_STNR_TYP.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@OLTC_END_A", OLTC_END_A.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_END_B", OLTC_END_B.Text)
                                                                  cmd.Parameters.AddWithValue("@OLTC_END_A1", OLTC_END_A1.Text)
                                                                  cmd.Parameters.AddWithValue("@OED1", OED1.Text)
                                                                  cmd.Parameters.AddWithValue("@OED2", OED2.Text)
                                                                  cmd.Parameters.AddWithValue("@OED3", OED3.Text)
                                                                  cmd.Parameters.AddWithValue("@OED4", OED4.Text)
                                                                  cmd.Parameters.AddWithValue("@Stiffner_count_BT", Stiffner_count_BT.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@oltc_stiffner1_count", oltc_stiffner1_count.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@oltc_stiffner2_count", oltc_stiffner2_count.SelectedItem)

                                                                  cmd.Parameters.AddWithValue("@stiffner_count_HV", stiffner_count_HV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@HVD1", HVD1.Text)
                                                                  cmd.Parameters.AddWithValue("@HVD2", HVD2.Text)
                                                                  cmd.Parameters.AddWithValue("@HVD3", HVD3.Text)
                                                                  cmd.Parameters.AddWithValue("@HVD4", HVD4.Text)

                                                                  cmd.Parameters.AddWithValue("@stiffner_count_LV", stiffner_count_LV.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@LVD1", LVD1.Text)
                                                                  cmd.Parameters.AddWithValue("@LVD2", LVD2.Text)
                                                                  cmd.Parameters.AddWithValue("@LVD3", LVD3.Text)
                                                                  cmd.Parameters.AddWithValue("@LVD4", LVD4.Text)

                                                                  cmd.Parameters.AddWithValue("@FS_HV_REQ", FS_HV_REQ.SelectedItem)
                                                                  cmd.Parameters.AddWithValue("@FS_LV_REQ", FS_LV_REQ.SelectedItem)


                                                              End Sub

                    If count > 0 Then
                        ' 🔁 Update existing row
                        Dim updateQuery As String = "UPDATE FlatStiffner_BT SET
                    Username=@Username,
                    HV_STNR_TYPE=@HV_STNR_TYPE, HV_FS_A=@HV_FS_A, HV_FS_A1=@HV_FS_A1, HV_FS_B=@HV_FS_B, HV_FS_B1=@HV_FS_B1,
                    LV_STNR_TYPE=@LV_STNR_TYPE, LV_FS_A=@LV_FS_A, LV_FS_A1=@LV_FS_A1, LV_FS_B=@LV_FS_B, LV_FS_B1=@LV_FS_B1,
                    OLTC_STNR_TYPE=@OLTC_STNR_TYPE, OLTC_S1_A=@OLTC_S1_A, OLTC_S1_A1=@OLTC_S1_A1, OLTC_S1_B=@OLTC_S1_B, OLTC_S1_B1=@OLTC_S1_B1,
                    OLTC_S1_T=@OLTC_S1_T, OS1D1=@OS1D1, OS1D2=@OS1D2,
                    OLTC_STNR2_TYPE=@OLTC_STNR2_TYPE, OLTC_S2_A=@OLTC_S2_A, OLTC_S2_B=@OLTC_S2_B, OLTC_S2_A1=@OLTC_S2_A1, OS2D1=@OS2D1, OS2D2=@OS2D2,
                    OLTC_END_STNR_TYP=@OLTC_END_STNR_TYP, OLTC_END_A=@OLTC_END_A, OLTC_END_B=@OLTC_END_B, OLTC_END_A1=@OLTC_END_A1,
                    OED1=@OED1, OED2=@OED2, OED3=@OED3, OED4=@OED4,Stiffner_count_BT=@Stiffner_count_BT,oltc_stiffner1_count=@oltc_stiffner1_count,oltc_stiffner2_count=@oltc_stiffner2_count,
                    stiffner_count_HV=@stiffner_count_HV,HVD1=@HVD1,HVD2=@HVD2,HVD3=@HVD3,HVD4=@HVD4,
                    stiffner_count_LV=@stiffner_count_LV,LVD1=@LVD1,LVD2=@LVD2,LVD3=@LVD3,LVD4=@LVD4,FS_HV_REQ=@FS_HV_REQ,FS_LV_REQ=@FS_LV_REQ
                WHERE wk_odr_No=@wk_odr_No"

                        Using updateCmd As New SqlCommand(updateQuery, connection)
                            fillParams(updateCmd)
                            updateCmd.ExecuteNonQuery()
                            MsgBox("Data updated successfully!")
                        End Using
                    Else
                        ' ➕ Insert new row
                        Dim insertQuery As String = "INSERT INTO FlatStiffner_BT (
                    Username, wk_odr_No, HV_STNR_TYPE, HV_FS_A, HV_FS_A1, HV_FS_B, HV_FS_B1,
                    LV_STNR_TYPE, LV_FS_A, LV_FS_A1, LV_FS_B, LV_FS_B1,
                    OLTC_STNR_TYPE, OLTC_S1_A, OLTC_S1_A1, OLTC_S1_B, OLTC_S1_B1, OLTC_S1_T, OS1D1, OS1D2,
                    OLTC_STNR2_TYPE, OLTC_S2_A, OLTC_S2_B, OLTC_S2_A1, OS2D1, OS2D2,
                    OLTC_END_STNR_TYP, OLTC_END_A, OLTC_END_B, OLTC_END_A1, OED1, OED2, OED3, OED4
                ) VALUES (
                    @Username, @wk_odr_No, @HV_STNR_TYPE, @HV_FS_A, @HV_FS_A1, @HV_FS_B, @HV_FS_B1,
                    @LV_STNR_TYPE, @LV_FS_A, @LV_FS_A1, @LV_FS_B, @LV_FS_B1,
                    @OLTC_STNR_TYPE, @OLTC_S1_A, @OLTC_S1_A1, @OLTC_S1_B, @OLTC_S1_B1, @OLTC_S1_T, @OS1D1, @OS1D2,
                    @OLTC_STNR2_TYPE, @OLTC_S2_A, @OLTC_S2_B, @OLTC_S2_A1, @OS2D1, @OS2D2,
                    @OLTC_END_STNR_TYP, @OLTC_END_A, @OLTC_END_B, @OLTC_END_A1, @OED1, @OED2, @OED3, @OED4
                )"

                        Using insertCmd As New SqlCommand(insertQuery, connection)
                            fillParams(insertCmd)
                            insertCmd.ExecuteNonQuery()
                            MsgBox("Data inserted successfully!")
                        End Using
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try


        'Try
        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO FlatStiffner_BT(Username,wk_odr_No,HV_STNR_TYPE,HV_FS_A,HV_FS_A1,HV_FS_B,HV_FS_B1,LV_STNR_TYPE,LV_FS_A,LV_FS_A1,LV_FS_B,LV_FS_B1,
        '    OLTC_STNR_TYPE,OLTC_S1_A,OLTC_S1_A1,OLTC_S1_B,OLTC_S1_B1,OLTC_S1_T,OS1D1,OS1D2,OLTC_STNR2_TYPE,OLTC_S2_A,OLTC_S2_B,OLTC_S2_A1,OS2D1,OS2D2,
        '    OLTC_END_STNR_TYP,OLTC_END_A,OLTC_END_B,OLTC_END_A1,OED1,OED2,OED3,OED4) 
        '    values (@Username,@wk_odr_No,@HV_STNR_TYPE,@HV_FS_A,@HV_FS_A1,@HV_FS_B,@HV_FS_B1,@LV_STNR_TYPE,@LV_FS_A,@LV_FS_A1,@LV_FS_B,@LV_FS_B1,
        '    @OLTC_STNR_TYPE,@OLTC_S1_A,@OLTC_S1_A1,@OLTC_S1_B,@OLTC_S1_B1,@OLTC_S1_T,@OS1D1,@OS1D2,@OLTC_STNR2_TYPE,@OLTC_S2_A,@OLTC_S2_B,@OLTC_S2_A1,@OS2D1,@OS2D2,
        '    @OLTC_END_STNR_TYP,@OLTC_END_A,@OLTC_END_B,@OLTC_END_A1,@OED1,@OED2,@OED3,@OED4)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@HV_STNR_TYPE", HV_STNR_TYPE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@HV_FS_A", HV_FS_A.Text)
        '    cmd.Parameters.AddWithValue("@HV_FS_A1", HV_FS_A1.Text)
        '    cmd.Parameters.AddWithValue("@HV_FS_B", HV_FS_B.Text)
        '    cmd.Parameters.AddWithValue("@HV_FS_B1", HV_FS_B1.Text)

        '    cmd.Parameters.AddWithValue("@LV_STNR_TYPE", LV_STNR_TYPE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@LV_FS_A", LV_FS_A.Text)
        '    cmd.Parameters.AddWithValue("@LV_FS_A1", LV_FS_A1.Text)
        '    cmd.Parameters.AddWithValue("@LV_FS_B", LV_FS_B.Text)
        '    cmd.Parameters.AddWithValue("@LV_FS_B1", LV_FS_B1.Text)

        '    cmd.Parameters.AddWithValue("@OLTC_STNR_TYPE", OLTC_STNR_TYPE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_S1_A", OLTC_S1_A.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S1_A1", OLTC_S1_A1.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S1_B", OLTC_S1_B.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S1_B1", OLTC_S1_B1.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S1_T", OLTC_S1_T.Text)
        '    cmd.Parameters.AddWithValue("@OS1D1", OS1D1.Text)
        '    cmd.Parameters.AddWithValue("@OS1D2", OS1D2.Text)

        '    cmd.Parameters.AddWithValue("@OLTC_STNR2_TYPE", OLTC_STNR2_TYPE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_S2_A", OLTC_S2_A.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S2_B", OLTC_S2_B.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_S2_A1", OLTC_S2_A1.Text)
        '    cmd.Parameters.AddWithValue("@OS2D1", OS2D1.Text)
        '    cmd.Parameters.AddWithValue("@OS2D2", OS2D2.Text)

        '    cmd.Parameters.AddWithValue("@OLTC_END_STNR_TYP", OLTC_END_STNR_TYP.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_END_A", OLTC_END_A.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_END_B", OLTC_END_B.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_END_A1", OLTC_END_A1.Text)
        '    cmd.Parameters.AddWithValue("@OED1", OED1.Text)
        '    cmd.Parameters.AddWithValue("@OED2", OED2.Text)
        '    cmd.Parameters.AddWithValue("@OED3", OED3.Text)
        '    cmd.Parameters.AddWithValue("@OED4", OED4.Text)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        BT_Flat_Stiffner()
    End Sub

    Public Sub BT_Accessories()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                ' Check if the wk_odr_No already exists
                Dim checkQuery As String = "SELECT COUNT(*) FROM Accessories_BT WHERE wk_odr_No = @wk_odr_No"
                Using checkCmd As New SqlCommand(checkQuery, connection)
                    checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                    Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                    Dim query As String

                    If exists > 0 Then
                        ' Update existing record
                        query = "UPDATE Accessories_BT SET 
                    Username=@Username, SAMP_PIPE_SIZE=@SAMP_PIPE_SIZE, SAM_VAL_POS=@SAM_VAL_POS, SV_XAXIS=@SV_XAXIS,
                    FILTER_VALUVE_SIZE=@FILTER_VALUVE_SIZE, FILTER_VAL_POS=@FILTER_VAL_POS, FV_XAXIS=@FV_XAXIS,
                    DRAIN_VALVE_SIZE=@DRAIN_VALVE_SIZE, DRAIN_VALVE_POS=@DRAIN_VALVE_POS, DV_XAXIS=@DV_XAXIS, DV_YAXIS=@DV_YAXIS,
                    NIFPS_REQ=@NIFPS_REQ, NIFPS_VALSIZE=@NIFPS_VALSIZE, N2_filling_req_HV=@N2_filling_req_HV, N2_qty_HV=@N2_qty_HV,
                    HV_X1_AXIS=@HV_X1_AXIS, HV_X2_AXIS=@HV_X2_AXIS, HV_X3_AXIS=@HV_X3_AXIS,
                    N2_filling_req_LV=@N2_filling_req_LV, N2_qty_LV=@N2_qty_LV, LV_X1_AXIS=@LV_X1_AXIS, LV_X2_AXIS=@LV_X2_AXIS, LV_X3_AXIS=@LV_X3_AXIS,
                    N2_filling_req_OLTC_opp=@N2_filling_req_OLTC_opp, N2_qty_OLTC_opp=@N2_qty_OLTC_opp, OLTC_opp_X1_AXIS=@OLTC_opp_X1_AXIS, OLTC_opp_X2_AXIS=@OLTC_opp_X2_AXIS,
                    OLTC_opp_X3_AXIS=@OLTC_opp_X3_AXIS,  N2_filling_req_OLTC=@N2_filling_req_OLTC, N2_qty_OLTC=@N2_qty_OLTC,
                    OLTC_X1_AXIS=@OLTC_X1_AXIS, OLTC_X2_AXIS=@OLTC_X2_AXIS, OLTC_X3_AXIS=@OLTC_X3_AXIS,
                    EARTHINGPAD_HV=@EARTHINGPAD_HV, HV_XAXIS=@HV_XAXIS, HV_YAXIS=@HV_YAXIS, EARTHINGPAD_POS=@EARTHINGPAD_POS, LV_XAXIS=@LV_XAXIS, LV_YAXIS=@LV_YAXIS,
                    earthing_pad_HV_type=@earthing_pad_HV_type,earthing_pad_LV_type=@earthing_pad_LV_type,NO_OF_PORT_BOT=@NO_OF_PORT_BOT,
                    NO_OF_PORT_BOT_X1=@NO_OF_PORT_BOT_X1,NO_OF_PORT_BOT_X2=@NO_OF_PORT_BOT_X2,NO_OF_PORT_BOT_X3=@NO_OF_PORT_BOT_X3,
                    NO_OF_PORT_BOT_Y1=@NO_OF_PORT_BOT_Y1,NO_OF_PORT_BOT_LV=@NO_OF_PORT_BOT_LV,NO_OF_PORT_BOT_X1_LV=@NO_OF_PORT_BOT_X1_LV,NO_OF_PORT_BOT_X2_LV=@NO_OF_PORT_BOT_X2_LV,NO_OF_PORT_BOT_X3_LV=@NO_OF_PORT_BOT_X3_LV
                    WHERE wk_odr_No=@wk_odr_No"
                    Else
                        '    ' Insert new record
                        '    query = "INSERT INTO Accessories_BT(Username,wk_odr_No,SAMP_PIPE_SIZE,SAM_VAL_POS,SV_XAXIS,SV_YAXIS,FILTER_VALUVE_SIZE,FILTER_VAL_POS,
                        'FV_XAXIS,FV_YAXIS,DRAIN_VALVE_SIZE,DRAIN_VALVE_POS,DV_XAXIS,DV_YAXIS,DGA_valve_req,DGA_VALVE_SIZE,DGA_VALVE_POS,DGA_XAXIS,DGA_YAXIS,NIFPS_REQ,NIFPS_VALSIZE,
                        'NIFPS_MAKE,N2_filling_req_HV,N2_qty_HV,HV_X1_AXIS,HV_X2_AXIS,HV_X3_AXIS,HV_Y1_AXIS,N2_filling_req_LV,N2_qty_LV,LV_X1_AXIS,LV_X2_AXIS,LV_X3_AXIS,LV_Y1_AXIS,
                        'N2_filling_req_OLTC_opp,N2_qty_OLTC_opp,OLTC_opp_X1_AXIS,OLTC_opp_X2_AXIS,OLTC_opp_X3_AXIS,OLTC_opp_Y1_AXIS,N2_filling_req_OLTC,N2_qty_OLTC,OLTC_X1_AXIS,
                        'OLTC_X2_AXIS,OLTC_X3_AXIS,OLTC_Y1_AXIS,EARTHINGPAD_HV,HV_XAXIS,HV_YAXIS,EARTHINGPAD_POS,LV_XAXIS,LV_YAXIS)
                        'VALUES (@Username,@wk_odr_No,@SAMP_PIPE_SIZE,@SAM_VAL_POS,@SV_XAXIS,@SV_YAXIS,@FILTER_VALUVE_SIZE,@FILTER_VAL_POS,
                        '@FV_XAXIS,@FV_YAXIS,@DRAIN_VALVE_SIZE,@DRAIN_VALVE_POS,@DV_XAXIS,@DV_YAXIS,@DGA_valve_req,@DGA_VALVE_SIZE,@DGA_VALVE_POS,@DGA_XAXIS,@DGA_YAXIS,
                        '@NIFPS_REQ,@NIFPS_VALSIZE,@NIFPS_MAKE,@N2_filling_req_HV,@N2_qty_HV,@HV_X1_AXIS,@HV_X2_AXIS,@HV_X3_AXIS,@HV_Y1_AXIS,
                        '@N2_filling_req_LV,@N2_qty_LV,@LV_X1_AXIS,@LV_X2_AXIS,@LV_X3_AXIS,@LV_Y1_AXIS,@N2_filling_req_OLTC_opp,@N2_qty_OLTC_opp,
                        '@OLTC_opp_X1_AXIS,@OLTC_opp_X2_AXIS,@OLTC_opp_X3_AXIS,@OLTC_opp_Y1_AXIS,@N2_filling_req_OLTC,@N2_qty_OLTC,@OLTC_X1_AXIS,@OLTC_X2_AXIS,@OLTC_X3_AXIS,
                        '@OLTC_Y1_AXIS,@EARTHINGPAD_HV,@HV_XAXIS,@HV_YAXIS,@EARTHINGPAD_POS,@LV_XAXIS,@LV_YAXIS)"
                    End If

                    Using cmd As New SqlCommand(query, connection)
                        ' Add all parameters same as before
                        cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                        cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                        ' (Add rest of your parameters here as before...)

                        cmd.Parameters.AddWithValue("@SAMP_PIPE_SIZE", SAMP_PIPE_SIZE.SelectedItem)
                        cmd.Parameters.AddWithValue("@SAM_VAL_POS", SAM_VAL_POS.SelectedItem)
                        cmd.Parameters.AddWithValue("@SV_XAXIS", SV_XAXIS.Text)

                        cmd.Parameters.AddWithValue("@FILTER_VALUVE_SIZE", FILTER_VALUVE_SIZE.SelectedItem)
                        cmd.Parameters.AddWithValue("@FILTER_VAL_POS", FILTER_VAL_POS.SelectedItem)
                        cmd.Parameters.AddWithValue("@FV_XAXIS", FV_XAXIS.Text)

                        cmd.Parameters.AddWithValue("@DRAIN_VALVE_SIZE", DRAIN_VALVE_SIZE.SelectedItem)
                        cmd.Parameters.AddWithValue("@DRAIN_VALVE_POS", DRAIN_VALVE_POS.SelectedItem)
                        cmd.Parameters.AddWithValue("@DV_XAXIS", DV_XAXIS.Text)
                        cmd.Parameters.AddWithValue("@DV_YAXIS", DV_YAXIS.Text)

                        cmd.Parameters.AddWithValue("@NIFPS_REQ", NIFPS_REQ.SelectedItem)
                        cmd.Parameters.AddWithValue("@NIFPS_VALSIZE", NIFPS_VALSIZE.SelectedItem)

                        cmd.Parameters.AddWithValue("@N2_filling_req_HV", N2_filling_req_HV.SelectedItem)
                        cmd.Parameters.AddWithValue("@N2_qty_HV", N2_qty_HV.SelectedItem)
                        cmd.Parameters.AddWithValue("@HV_X1_AXIS", HV_X1_AXIS.Text)
                        cmd.Parameters.AddWithValue("@HV_X2_AXIS", HV_X2_AXIS.Text)
                        cmd.Parameters.AddWithValue("@HV_X3_AXIS", HV_X3_AXIS.Text)

                        cmd.Parameters.AddWithValue("@N2_filling_req_LV", N2_filling_req_LV.SelectedItem)
                        cmd.Parameters.AddWithValue("@N2_qty_LV", N2_qty_LV.SelectedItem)
                        cmd.Parameters.AddWithValue("@LV_X1_AXIS", LV_X1_AXIS.Text)
                        cmd.Parameters.AddWithValue("@LV_X2_AXIS", LV_X2_AXIS.Text)
                        cmd.Parameters.AddWithValue("@LV_X3_AXIS", LV_X3_AXIS.Text)

                        cmd.Parameters.AddWithValue("@N2_filling_req_OLTC_opp", N2_filling_req_OLTC_opp.SelectedItem)
                        cmd.Parameters.AddWithValue("@N2_qty_OLTC_opp", N2_qty_OLTC_opp.SelectedItem)
                        cmd.Parameters.AddWithValue("@OLTC_opp_X1_AXIS", OLTC_opp_X1_AXIS.Text)
                        cmd.Parameters.AddWithValue("@OLTC_opp_X2_AXIS", OLTC_opp_X2_AXIS.Text)
                        cmd.Parameters.AddWithValue("@OLTC_opp_X3_AXIS", OLTC_opp_X3_AXIS.Text)

                        cmd.Parameters.AddWithValue("@N2_filling_req_OLTC", N2_filling_req_OLTC.SelectedItem)
                        cmd.Parameters.AddWithValue("@N2_qty_OLTC", N2_qty_OLTC.SelectedItem)
                        cmd.Parameters.AddWithValue("@OLTC_X1_AXIS", OLTC_X1_AXIS.Text)
                        cmd.Parameters.AddWithValue("@OLTC_X2_AXIS", OLTC_X2_AXIS.Text)
                        cmd.Parameters.AddWithValue("@OLTC_X3_AXIS", OLTC_X3_AXIS.Text)

                        cmd.Parameters.AddWithValue("@EARTHINGPAD_HV", EARTHINGPAD_HV.SelectedItem)
                        cmd.Parameters.AddWithValue("@HV_XAXIS", HV_XAXIS.Text)
                        cmd.Parameters.AddWithValue("@HV_YAXIS", HV_YAXIS.Text)

                        cmd.Parameters.AddWithValue("@EARTHINGPAD_POS", EARTHINGPAD_POS.SelectedItem)
                        cmd.Parameters.AddWithValue("@LV_XAXIS", LV_XAXIS.Text)
                        cmd.Parameters.AddWithValue("@LV_YAXIS", LV_YAXIS.Text)

                        cmd.Parameters.AddWithValue("@earthing_pad_HV_type", earthing_pad_HV_type.SelectedItem)
                        cmd.Parameters.AddWithValue("@earthing_pad_LV_type", earthing_pad_LV_type.SelectedItem)

                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT", NO_OF_PORT_BOT.SelectedItem)

                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X1", NO_OF_PORT_BOT_X1.Text)
                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X2", NO_OF_PORT_BOT_X2.Text)
                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X3", NO_OF_PORT_BOT_X3.Text)

                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_Y1", NO_OF_PORT_BOT_Y1.Text)

                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_LV", NO_OF_PORT_BOT_LV.SelectedItem)

                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X1_LV", NO_OF_PORT_BOT_X1_LV.Text)
                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X2_LV", NO_OF_PORT_BOT_X2_LV.Text)
                        cmd.Parameters.AddWithValue("@NO_OF_PORT_BOT_X3_LV", NO_OF_PORT_BOT_X3_LV.Text)


                        Dim result As Integer = cmd.ExecuteNonQuery()
                        If result > 0 Then
                            MsgBox("Data saved successfully!")
                        Else
                            MsgBox("Data not saved!")
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Unable to save record: " & ex.Message)
        End Try


        'Try
        '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '    Dim connection As New SqlConnection(connectionString)

        '    connection.Open()
        '    Dim query As String = "INSERT INTO Accessories_BT(Username,wk_odr_No,SAMP_PIPE_SIZE,SAM_VAL_POS,SV_XAXIS,SV_YAXIS,FILTER_VALUVE_SIZE,FILTER_VAL_POS,
        '    FV_XAXIS,FV_YAXIS,DRAIN_VALVE_SIZE,DRAIN_VALVE_POS,DV_XAXIS,DV_YAXIS,DGA_valve_req,DGA_VALVE_SIZE,DGA_VALVE_POS,DGA_XAXIS,DGA_YAXIS,NIFPS_REQ,NIFPS_VALSIZE,
        '    NIFPS_MAKE,N2_filling_req_HV,N2_qty_HV,HV_X1_AXIS,HV_X2_AXIS,HV_X3_AXIS,HV_Y1_AXIS,N2_filling_req_LV,N2_qty_LV,LV_X1_AXIS,LV_X2_AXIS,LV_X3_AXIS,LV_Y1_AXIS,
        '    N2_filling_req_OLTC_opp,N2_qty_OLTC_opp,OLTC_opp_X1_AXIS,OLTC_opp_X2_AXIS,OLTC_opp_X3_AXIS,OLTC_opp_Y1_AXIS,N2_filling_req_OLTC,N2_qty_OLTC,OLTC_X1_AXIS,
        '    OLTC_X2_AXIS,OLTC_X3_AXIS,OLTC_Y1_AXIS,EARTHINGPAD_HV,HV_XAXIS,HV_YAXIS,EARTHINGPAD_POS,LV_XAXIS,LV_YAXIS) 
        '    values (@Username,@wk_odr_No,@SAMP_PIPE_SIZE,@SAM_VAL_POS,@SV_XAXIS,@SV_YAXIS,@FILTER_VALUVE_SIZE,@FILTER_VAL_POS,
        '    @FV_XAXIS,@FV_YAXIS,@DRAIN_VALVE_SIZE,@DRAIN_VALVE_POS,@DV_XAXIS,@DV_YAXIS,@DGA_valve_req,@DGA_VALVE_SIZE,@DGA_VALVE_POS,@DGA_XAXIS,@DGA_YAXIS,@NIFPS_REQ,@NIFPS_VALSIZE,
        '    @NIFPS_MAKE,@N2_filling_req_HV,@N2_qty_HV,@HV_X1_AXIS,@HV_X2_AXIS,@HV_X3_AXIS,@HV_Y1_AXIS,@N2_filling_req_LV,@N2_qty_LV,@LV_X1_AXIS,@LV_X2_AXIS,@LV_X3_AXIS,@LV_Y1_AXIS,
        '    @N2_filling_req_OLTC_opp,@N2_qty_OLTC_opp,@OLTC_opp_X1_AXIS,@OLTC_opp_X2_AXIS,@OLTC_opp_X3_AXIS,@OLTC_opp_Y1_AXIS,@N2_filling_req_OLTC,@N2_qty_OLTC,@OLTC_X1_AXIS,
        '    @OLTC_X2_AXIS,@OLTC_X3_AXIS,@OLTC_Y1_AXIS,@EARTHINGPAD_HV,@HV_XAXIS,@HV_YAXIS,@EARTHINGPAD_POS,@LV_XAXIS,@LV_YAXIS)"

        '    Dim cmd As New SqlCommand(query, connection)
        '    cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '    cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '    cmd.Parameters.AddWithValue("@SAMP_PIPE_SIZE", SAMP_PIPE_SIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@SAM_VAL_POS", SAM_VAL_POS.SelectedItem)
        '    cmd.Parameters.AddWithValue("@SV_XAXIS", SV_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@SV_YAXIS", SV_YAXIS.Text)

        '    cmd.Parameters.AddWithValue("@FILTER_VALUVE_SIZE", FILTER_VALUVE_SIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@FILTER_VAL_POS", FILTER_VAL_POS.SelectedItem)
        '    cmd.Parameters.AddWithValue("@FV_XAXIS", FV_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@FV_YAXIS", FV_YAXIS.Text)

        '    cmd.Parameters.AddWithValue("@DRAIN_VALVE_SIZE", DRAIN_VALVE_SIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@DRAIN_VALVE_POS", DRAIN_VALVE_POS.SelectedItem)
        '    cmd.Parameters.AddWithValue("@DV_XAXIS", DV_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@DV_YAXIS", DV_YAXIS.Text)

        '    cmd.Parameters.AddWithValue("@DGA_valve_req", DGA_valve_req.SelectedItem)
        '    cmd.Parameters.AddWithValue("@DGA_VALVE_SIZE", DGA_VALVE_SIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@DGA_VALVE_POS", DGA_VALVE_POS.SelectedItem)
        '    cmd.Parameters.AddWithValue("@DGA_XAXIS", DGA_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@DGA_YAXIS", DGA_YAXIS.Text)

        '    cmd.Parameters.AddWithValue("@NIFPS_REQ", NIFPS_REQ.SelectedItem)
        '    cmd.Parameters.AddWithValue("@NIFPS_VALSIZE", NIFPS_VALSIZE.SelectedItem)
        '    cmd.Parameters.AddWithValue("@NIFPS_MAKE", NIFPS_MAKE.SelectedItem)

        '    cmd.Parameters.AddWithValue("@N2_filling_req_HV", N2_filling_req_HV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@N2_qty_HV", N2_qty_HV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@HV_X1_AXIS", HV_X1_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@HV_X2_AXIS", HV_X2_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@HV_X3_AXIS", HV_X3_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@HV_Y1_AXIS", HV_Y1_AXIS.Text)

        '    cmd.Parameters.AddWithValue("@N2_filling_req_LV", N2_filling_req_LV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@N2_qty_LV", N2_qty_LV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@LV_X1_AXIS", LV_X1_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@LV_X2_AXIS", LV_X2_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@LV_X3_AXIS", LV_X3_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@LV_Y1_AXIS", LV_Y1_AXIS.Text)

        '    cmd.Parameters.AddWithValue("@N2_filling_req_OLTC_opp", N2_filling_req_OLTC_opp.SelectedItem)
        '    cmd.Parameters.AddWithValue("@N2_qty_OLTC_opp", N2_qty_OLTC_opp.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_opp_X1_AXIS", OLTC_opp_X1_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_opp_X2_AXIS", OLTC_opp_X2_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_opp_X3_AXIS", OLTC_opp_X3_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_opp_Y1_AXIS", OLTC_opp_Y1_AXIS.Text)

        '    cmd.Parameters.AddWithValue("@N2_filling_req_OLTC", N2_filling_req_OLTC.SelectedItem)
        '    cmd.Parameters.AddWithValue("@N2_qty_OLTC", N2_qty_OLTC.SelectedItem)
        '    cmd.Parameters.AddWithValue("@OLTC_X1_AXIS", OLTC_X1_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_X2_AXIS", OLTC_X2_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_X3_AXIS", OLTC_X3_AXIS.Text)
        '    cmd.Parameters.AddWithValue("@OLTC_Y1_AXIS", OLTC_Y1_AXIS.Text)

        '    cmd.Parameters.AddWithValue("@EARTHINGPAD_HV", EARTHINGPAD_HV.SelectedItem)
        '    cmd.Parameters.AddWithValue("@HV_XAXIS", HV_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@HV_YAXIS", HV_YAXIS.Text)

        '    cmd.Parameters.AddWithValue("@EARTHINGPAD_POS", EARTHINGPAD_POS.SelectedItem)
        '    cmd.Parameters.AddWithValue("@LV_XAXIS", LV_XAXIS.Text)
        '    cmd.Parameters.AddWithValue("@LV_YAXIS", LV_YAXIS.Text)

        '    Dim result As Integer = cmd.ExecuteNonQuery()
        '    connection.Close()

        '    If result > 0 Then
        '        MsgBox("Data inserted successfully!!!")
        '    Else
        '        MsgBox("Data not inserted!!!")
        '    End If

        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try

    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        BT_Accessories()
    End Sub



    Private Sub Button17_Click(sender As Object, e As EventArgs)
        TabControl1.SelectTab("TabPage8")
    End Sub

    Private Sub Button16_Click_1(sender As Object, e As EventArgs)



        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()


        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If


        TabControl1.SelectTab("TabPage3")

    End Sub

    Private Sub btm_tk_i_len_TextChanged(sender As Object, e As EventArgs) Handles btm_tk_i_len.TextChanged

        btm_plt_len.Text = Val(btm_tk_i_len.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

    End Sub

    Private Sub btm_tk_i_wdh_TextChanged(sender As Object, e As EventArgs) Handles btm_tk_i_wdh.TextChanged
        If btm_tk_type.SelectedItem = "Type 1" Then
            btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

            'bt_tk_wdh_2.Text = Main_Panel.bt_tk_wdh_2V
        Else
            ' bt_tk_wdh_2.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

            'btm_plt_width.Text = Main_Panel.btm_plt_widthV
        End If


    End Sub



    Private Sub Side_Plt_thk_TextChanged(sender As Object, e As EventArgs) Handles Side_Plt_thk.TextChanged
        btm_plt_width.Text = Val(btm_tk_i_wdh.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30
        btm_plt_len.Text = Val(btm_tk_i_len.Text) + Val(Side_Plt_thk.Text) + Val(Side_Plt_thk.Text) + 30

    End Sub




    Private Sub prob_support_cmb_SelectedIndexChanged(sender As Object, e As EventArgs) Handles prob_support_cmb.SelectedIndexChanged

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If prob_support_cmb.SelectedItem = "Yes" Then
            If No_of_Rails.SelectedItem = "2" Then


                If No_of_Rails1.SelectedItem = "Option 1" Then



                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    ' prob_no.Enabled = True


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False




                ElseIf No_of_Rails1.SelectedItem = "Option 2" Then



                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    'prob_no.Enabled = True


                    Prob_Y1axis.Enabled = True
                    Prob_Y2axis.Enabled = True




                End If





            ElseIf No_of_Rails.SelectedItem = "3" Then

                prob_support_cmb.SelectedItem = "No"
                prob_support_cmb.Enabled = False
                prob_no.Enabled = False
                PROB_HT.Enabled = False
                Prob_X1axis.Enabled = False
                Prob_X2axis.Enabled = False
                Prob_Y1axis.Enabled = False
                Prob_Y2axis.Enabled = False

            End If


        ElseIf prob_support_cmb.SelectedItem = "No" Then
            prob_no.Enabled = False
            PROB_HT.Enabled = False
            Prob_X1axis.Enabled = False
            Prob_X2axis.Enabled = False
            Prob_Y1axis.Enabled = False
            Prob_Y2axis.Enabled = False

        End If


    End Sub



    Private Sub N2_filling_req_HV_SelectedIndexChanged(sender As Object, e As EventArgs)
        'otherparts 1
        If N2_filling_req_HV.SelectedItem = "Yes" Then

            N2_qty_HV.Enabled = True

            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

            If N2_qty_HV.SelectedItem = "1" Then


                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = False
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "2" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "3" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_HV.SelectedItem = "No" Then
            N2_qty_HV.Enabled = False
            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False


        End If





    End Sub

    Private Sub N2_filling_req_LV_SelectedIndexChanged(sender As Object, e As EventArgs)

        If N2_filling_req_LV.SelectedItem = "Yes" Then

            N2_qty_LV.Enabled = True
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

            If N2_qty_LV.SelectedItem = "1" Then


                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = False
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "2" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "3" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_LV.SelectedItem = "No" Then
            N2_qty_LV.Enabled = False
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

        End If

    End Sub

    Private Sub N2_filling_req_OLTC_opp_SelectedIndexChanged(sender As Object, e As EventArgs)
        If N2_filling_req_OLTC_opp.SelectedItem = "Yes" Then

            N2_qty_OLTC_opp.Enabled = True
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False

            If N2_qty_OLTC_opp.SelectedItem = "1" Then


                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = False
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "2" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "3" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = True

            End If


        ElseIf N2_filling_req_OLTC_opp.SelectedItem = "No" Then
            N2_qty_OLTC_opp.Enabled = False
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False


        End If
    End Sub

    Private Sub N2_filling_req_OLTC_SelectedIndexChanged(sender As Object, e As EventArgs)

        If N2_filling_req_OLTC.SelectedItem = "Yes" Then

            N2_qty_OLTC.Enabled = True

            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

            If N2_qty_OLTC.SelectedItem = "1" Then


                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = False
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "2" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "3" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = True

            End If

        ElseIf N2_filling_req_OLTC.SelectedItem = "No" Then
            N2_qty_OLTC.Enabled = False
            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

        End If

    End Sub






    Private Sub curb_size_SelectedIndexChanged(sender As Object, e As EventArgs) Handles curb_size.SelectedIndexChanged


        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles btm_len_chkbx.CheckedChanged

        If btm_len_chkbx.Checked Then
            btm_plt_len.Enabled = True
        Else
            btm_plt_len.Enabled = False
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles btm_plt_wdt_ck_bx.CheckedChanged

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()


        If btm_tk_type.SelectedItem = "Type 1" Then
            Btm_TK_pic_box.Image = My.Resources.BT_T1

            If btm_plt_wdt_ck_bx.Checked Then
                btm_plt_width.Enabled = True

            Else
                btm_plt_width.Enabled = False

            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then
            Btm_TK_pic_box.Image = My.Resources.BT_T2

            If btm_plt_wdt_ck_bx.Checked Then
                btm_plt_width.Enabled = True

            Else

                btm_plt_width.Enabled = False
            End If


        End If


    End Sub

    Private Sub N2_qty_HV_SelectedIndexChanged(sender As Object, e As EventArgs)
        If N2_qty_HV.SelectedItem = "1" Then


            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = True
            HV_X3_AXIS.Enabled = False

        ElseIf N2_qty_HV.SelectedItem = "3" Then

            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = True
            HV_X3_AXIS.Enabled = True

        End If
    End Sub

    Private Sub N2_qty_LV_SelectedIndexChanged(sender As Object, e As EventArgs)

        If N2_qty_LV.SelectedItem = "1" Then


            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = True
            LV_X3_AXIS.Enabled = False

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = True
            LV_X3_AXIS.Enabled = True

        End If



    End Sub

    Private Sub N2_qty_OLTC_opp_SelectedIndexChanged(sender As Object, e As EventArgs)

        If N2_qty_OLTC_opp.SelectedItem = "1" Then


            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC_opp.SelectedItem = "2" Then

            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = True
            OLTC_opp_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC_opp.SelectedItem = "3" Then

            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = True
            OLTC_opp_X3_AXIS.Enabled = True

        End If


    End Sub

    Private Sub N2_qty_OLTC_SelectedIndexChanged(sender As Object, e As EventArgs)

        If N2_qty_OLTC.SelectedItem = "1" Then


            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC.SelectedItem = "2" Then

            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = True
            OLTC_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC.SelectedItem = "3" Then

            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = True
            OLTC_X3_AXIS.Enabled = True

        End If


    End Sub



    Private Sub btm_plt_ht_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_plt_ht.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tank_ht1

                bt_tk_l.Location = New System.Drawing.Point(205, 171)
                bt_tk_ht.Location = New System.Drawing.Point(40, 154)
            Else

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tank_ht

                bt_tk_l.Location = New System.Drawing.Point(205, 171)
                bt_tk_ht.Location = New System.Drawing.Point(568, 154)
            Else

            End If

        End If
        'Btm_TK_pic_box.Image = My.Resources.bt_tank_ht


        bt_tk_l.Show()
        bt_tk_ht.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()
        W1.Hide()
        W2.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

        bt_tk_l.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_tk_ht.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        'bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub btm_plt_ht_SelectedIndexChanged(sender As Object, e As EventArgs) Handles btm_plt_ht.SelectedIndexChanged

        If btm_plt_ht.SelectedItem = "Custom" Then

            Cus_ht.Enabled = True
        Else

            Cus_ht.Enabled = False

        End If


    End Sub

    Private Sub btm_plt_len_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_plt_len.MouseClick

        Load_OLTC_Pkt_Position()

        'Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

        'Btm_TK_pic_box.Image = My.Resources.bt_tank_ht
        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tank_ht1

                bt_tk_l.Location = New System.Drawing.Point(205, 171)
                bt_tk_ht.Location = New System.Drawing.Point(40, 154)
            Else

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tank_ht

                bt_tk_l.Location = New System.Drawing.Point(205, 171)
                bt_tk_ht.Location = New System.Drawing.Point(568, 154)
            Else

            End If

        End If

        bt_tk_l.Show()
        bt_tk_ht.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        bt_tk_l.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bt_tk_ht.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        'bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub btm_plt_thk_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_plt_thk.MouseClick

        Btm_TK_pic_box.Image = My.Resources.bt_tk_wth_T1

        bt_thickness.Location = New System.Drawing.Point(373, 123)
        bt_thickness1.Location = New System.Drawing.Point(388, 133)

        'If OLTC_pkt_position = "Left" Then

        '    If btm_tk_curb_pro.SelectedItem = "Type 1" Then
        '        Btm_TK_pic_box.Image = My.Resources.BT_TK_RWT

        '        bt_tk_lin.Location = New System.Drawing.Point(234, 103)
        '        BT_tk_win.Location = New System.Drawing.Point(480, 121)
        '        tk_thk.Location = New System.Drawing.Point(615, 142)
        '    Else
        '        Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT1

        '        bt_tk_lin.Location = New System.Drawing.Point(234, 140)
        '        BT_tk_win.Location = New System.Drawing.Point(460, 155)
        '        tk_thk.Location = New System.Drawing.Point(600, 100)

        '    End If

        'ElseIf OLTC_pkt_position = "Right" Then

        '    If btm_tk_curb_pro.SelectedItem = "Type 1" Then
        '        Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

        '        bt_tk_lin.Location = New System.Drawing.Point(234, 103)
        '        BT_tk_win.Location = New System.Drawing.Point(125, 121)
        '        tk_thk.Location = New System.Drawing.Point(10, 142)

        '    Else
        '        Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT

        '        bt_tk_lin.Location = New System.Drawing.Point(234, 140)
        '        BT_tk_win.Location = New System.Drawing.Point(150, 155)
        '        tk_thk.Location = New System.Drawing.Point(10, 100)

        '    End If

        'End If

        'Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT
        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        bt_thickness1.Show()
        bt_thickness.Show()

        bt_tk_l.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_tk_ht.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        tk_thk.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BT_tk_win.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_thickness.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bt_thickness1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        W1.Hide()
        W2.Hide()

    End Sub

    Private Sub btm_tk_i_wdh_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_tk_i_wdh.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_RWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(480, 121)
                tk_thk.Location = New System.Drawing.Point(615, 142)
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT1

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(460, 155)
                tk_thk.Location = New System.Drawing.Point(600, 100)

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(125, 121)
                tk_thk.Location = New System.Drawing.Point(10, 142)

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(150, 155)
                tk_thk.Location = New System.Drawing.Point(10, 100)

            End If

        End If

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        tk_thk.Show()
        BT_tk_win.Show()
        bt_tk_lin.Show()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BT_tk_win.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        W1.Hide()
        W2.Hide()
        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub Side_Plt_thk_MouseClick(sender As Object, e As MouseEventArgs) Handles Side_Plt_thk.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_RWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(480, 121)
                tk_thk.Location = New System.Drawing.Point(615, 142)
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT1

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(460, 155)
                tk_thk.Location = New System.Drawing.Point(600, 100)

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(125, 121)
                tk_thk.Location = New System.Drawing.Point(10, 142)

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(150, 155)
                tk_thk.Location = New System.Drawing.Point(10, 100)

            End If

        End If

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        tk_thk.Show()
        BT_tk_win.Show()
        bt_tk_lin.Show()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        BT_tk_win.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub OLTC_pck_len_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_pck_len.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If




        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()



        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub







    Private Sub btm_tk_i_len_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_tk_i_len.MouseClick



        'If btm_tk_curb_pro.SelectedItem = "Type 1" Then

        '    Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

        'Else
        '    Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2LHV

        'End If

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_RWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(480, 121)
                tk_thk.Location = New System.Drawing.Point(615, 142)
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT1

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(460, 155)
                tk_thk.Location = New System.Drawing.Point(600, 100)

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BT_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 103)
                BT_tk_win.Location = New System.Drawing.Point(125, 121)
                tk_thk.Location = New System.Drawing.Point(10, 142)

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_LWT

                bt_tk_lin.Location = New System.Drawing.Point(234, 140)
                BT_tk_win.Location = New System.Drawing.Point(150, 155)
                tk_thk.Location = New System.Drawing.Point(10, 100)

            End If

        End If



        bt_tk_lin.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        BT_tk_win.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        tk_thk.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        tk_thk.Show()
        BT_tk_win.Show()
        bt_tk_lin.Show()

        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub



    Private Sub OLTC_pck_wdh_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_pck_wdh.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If


        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()
        bt_thickness1.Hide()
        bt_thickness.Hide()


    End Sub

    Private Sub tk_curb_wdh_MouseClick(sender As Object, e As MouseEventArgs) Handles tk_curb_wdh.MouseClick


        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If


        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub Curb_wdh_LV_MouseClick(sender As Object, e As MouseEventArgs) Handles Curb_wdh_LV.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub Curb_OLTC_End_MouseClick(sender As Object, e As MouseEventArgs) Handles Curb_OLTC_End.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If


        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub TextBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles Curb_wdh_D.MouseClick
        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If



        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub TextBox2_MouseClick(sender As Object, e As MouseEventArgs) Handles GPX1.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()
    End Sub

    Private Sub GPX2_MouseClick(sender As Object, e As MouseEventArgs) Handles GPX2.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub GPY1_MouseClick(sender As Object, e As MouseEventArgs) Handles GPY1.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub TextBox5_MouseClick(sender As Object, e As MouseEventArgs) Handles GPY2.MouseClick

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        'Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub btm_plt_width_MouseClick(sender As Object, e As MouseEventArgs) Handles btm_plt_width.MouseClick

        ' Btm_TK_pic_box.Image = My.Resources.bt_tk_wth_T1

        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        W1.Hide()
        W2.Hide()

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then
            Btm_TK_pic_box.Image = My.Resources.bt_tk_wth_T1

            W1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            W2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            W1.Location = New System.Drawing.Point(306, 188)
            W2.Location = New System.Drawing.Point(299, 56)

            W1.Show()
            W2.Hide()

        Else btm_tk_type.SelectedItem = "Type 2"
            Btm_TK_pic_box.Image = My.Resources.T2_BTW1_W2_


            W1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            W2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            W1.Location = New System.Drawing.Point(306, 235)
            W2.Location = New System.Drawing.Point(299, 56)

            W1.Show()
            W2.Show()
        End If




    End Sub

    Private Sub bt_tk_wdh_2_MouseClick(sender As Object, e As MouseEventArgs)



        Btm_TK_pic_box.Image = My.Resources.T2_BTW1_W2_

        W2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        W1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        bush_pkt_X1_axis.Hide()
        bush_pkt_X2_axis.Hide()
        bush_pkt_Y1_axis.Hide()
        R.Hide()


        Curb_A.Hide()
        Curb_B.Hide()
        Curb_C.Hide()
        Curb_D.Hide()

        GP_Yaxis_1.Hide()
        GP_Yaxis_2.Hide()
        GP_Xaxis_1.Hide()
        GP_Xaxis_2.Hide()

        OLTC_PKT_WDH.Hide()
        OLTC_PKT_LEN.Hide()


        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()

        W1.Show()
        W2.Show()

        W1.Location = New System.Drawing.Point(306, 235)
        W2.Location = New System.Drawing.Point(299, 56)

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub BP_X1_MouseClick(sender As Object, e As MouseEventArgs)

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If



        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()

        bush_pkt_X1_axis.Show()
        bush_pkt_X2_axis.Show()
        bush_pkt_Y1_axis.Show()
        R.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub BP_X2_MouseClick(sender As Object, e As MouseEventArgs)

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()

        bush_pkt_X1_axis.Show()
        bush_pkt_X2_axis.Show()
        bush_pkt_Y1_axis.Show()
        R.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub BP_X3_MouseClick(sender As Object, e As MouseEventArgs)

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If

        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()

        bush_pkt_X1_axis.Show()
        bush_pkt_X2_axis.Show()
        bush_pkt_Y1_axis.Show()
        R.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub

    Private Sub BP_R1_MouseClick(sender As Object, e As MouseEventArgs)

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in1

                Curb_A.Location = New System.Drawing.Point(490, 284)
                Curb_B.Location = New System.Drawing.Point(8, 200)
                Curb_C.Location = New System.Drawing.Point(515, 25)
                Curb_D.Location = New System.Drawing.Point(595, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(125, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(25, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()

            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB1

                Curb_A.Location = New System.Drawing.Point(440, 244)
                Curb_B.Location = New System.Drawing.Point(13, 170)
                Curb_C.Location = New System.Drawing.Point(175, 25)
                Curb_D.Location = New System.Drawing.Point(595, 100)
                GP_Yaxis_1.Location = New System.Drawing.Point(465, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(465, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(221, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(361, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(114, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(18, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(114, 163)
                R.Location = New System.Drawing.Point(552, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()

            End If

        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.bt_tk_curb_in

                Curb_A.Location = New System.Drawing.Point(124, 284)
                Curb_B.Location = New System.Drawing.Point(578, 200)
                Curb_C.Location = New System.Drawing.Point(151, 25)
                Curb_D.Location = New System.Drawing.Point(8, 125)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 185)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 92)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 92)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(441, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(552, 74)

                bush_pkt_X1_axis.Hide()
                bush_pkt_X2_axis.Hide()
                bush_pkt_Y1_axis.Hide()
                R.Hide()
            Else
                Btm_TK_pic_box.Image = My.Resources.T2_BTM_TK_CURB

                Curb_A.Location = New System.Drawing.Point(160, 244)
                Curb_B.Location = New System.Drawing.Point(578, 165)
                Curb_C.Location = New System.Drawing.Point(445, 25)
                Curb_D.Location = New System.Drawing.Point(8, 95)
                GP_Yaxis_1.Location = New System.Drawing.Point(125, 120)
                GP_Yaxis_2.Location = New System.Drawing.Point(125, 175)
                GP_Xaxis_1.Location = New System.Drawing.Point(216, 97)
                GP_Xaxis_2.Location = New System.Drawing.Point(356, 97)
                OLTC_PKT_WDH.Location = New System.Drawing.Point(446, 111)
                OLTC_PKT_LEN.Location = New System.Drawing.Point(556, 70)
                bush_pkt_X1_axis.Location = New System.Drawing.Point(182, 284)
                bush_pkt_X2_axis.Location = New System.Drawing.Point(368, 284)
                bush_pkt_Y1_axis.Location = New System.Drawing.Point(433, 163)
                R.Location = New System.Drawing.Point(82, 217)

                bush_pkt_X1_axis.Show()
                bush_pkt_X2_axis.Show()
                bush_pkt_Y1_axis.Show()
                R.Show()
            End If

        End If


        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        Curb_A.Show()
        Curb_B.Show()
        Curb_C.Show()
        Curb_D.Show()

        GP_Yaxis_1.Show()
        GP_Yaxis_2.Show()
        GP_Xaxis_1.Show()
        GP_Xaxis_2.Show()

        OLTC_PKT_WDH.Show()
        OLTC_PKT_LEN.Show()

        bush_pkt_X1_axis.Show()
        bush_pkt_X2_axis.Show()
        bush_pkt_Y1_axis.Show()
        R.Show()

        tk_thk.Hide()
        BT_tk_win.Hide()
        bt_tk_lin.Hide()

        bt_tk_l.Hide()
        bt_tk_ht.Hide()
        W1.Hide()
        W2.Hide()

        bt_thickness1.Hide()
        bt_thickness.Hide()

    End Sub


    Private Sub No_of_Rails_SelectedIndexChanged(sender As Object, e As EventArgs) Handles No_of_Rails.SelectedIndexChanged






        If No_of_Rails.SelectedItem = "2" Then

            ' No_of_Rails1.Items.Clear()
            'No_of_Rails1.Items.Add("Option 1")
            ' No_of_Rails1.Items.Add("Option 2")

            'No_of_Rails1.SelectedItem = "Option 1"
            No_of_Rails1.Enabled = True
            Distance_A.Enabled = False
            Distance_B.Enabled = False
            prob_support_cmb.SelectedItem = "Yes"
            prob_support_cmb.Enabled = True
            BS_L.Enabled = True
            BS_R.Enabled = True

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                If prob_support_cmb.SelectedItem = "Yes" Then

                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False


                ElseIf prob_support_cmb.SelectedItem = "No" Then
                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = False
                    Prob_X2axis.Enabled = False
                    PROB_HT.Enabled = False
                    prob_no.Enabled = False


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False

                End If

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                If prob_support_cmb.SelectedItem = "Yes" Then

                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    'prob_no.Enabled = True


                    Prob_Y1axis.Enabled = True
                    Prob_Y2axis.Enabled = True


                ElseIf prob_support_cmb.SelectedItem = "No" Then
                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = False
                    Prob_X2axis.Enabled = False
                    PROB_HT.Enabled = False
                    prob_no.Enabled = False


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False

                End If



            End If





        ElseIf No_of_Rails.SelectedItem = "3" Then

            No_of_Rails1.Enabled = False

            Distance_A.Enabled = True
            Distance_B.Enabled = True


            prob_support_cmb.SelectedItem = "No"
            prob_support_cmb.Enabled = False
            prob_no.Enabled = False
            PROB_HT.Enabled = False
            Prob_X1axis.Enabled = False
            Prob_X2axis.Enabled = False
            Prob_Y1axis.Enabled = False
            Prob_Y2axis.Enabled = False

            BS_L.Enabled = False
            BS_R.Enabled = False

        End If

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

            UB_L_1.Show()
            UB_L1_1.Show()
            UB_W_1.Show()
            UB_T2_1.Show()

            UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If
    End Sub

    Private Sub No_of_Rails1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles No_of_Rails1.SelectedIndexChanged





        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(533, 425)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(536, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If




        If prob_support_cmb.SelectedItem = "Yes" Then
            If No_of_Rails.SelectedItem = "2" Then


                If No_of_Rails1.SelectedItem = "Option 1" Then



                    prob_no.SelectedItem = "2 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    'prob_no.Enabled = True


                    Prob_Y1axis.Enabled = False
                    Prob_Y2axis.Enabled = False




                ElseIf No_of_Rails1.SelectedItem = "Option 2" Then



                    prob_no.SelectedItem = "4 Nos"
                    prob_no.Enabled = False

                    Prob_X1axis.Enabled = True
                    Prob_X2axis.Enabled = True
                    PROB_HT.Enabled = True
                    ' prob_no.Enabled = True


                    Prob_Y1axis.Enabled = True
                    Prob_Y2axis.Enabled = True




                End If





            ElseIf No_of_Rails.SelectedItem = "3" Then

                prob_support_cmb.SelectedItem = "No"
                prob_support_cmb.Enabled = False
                prob_no.Enabled = False
                PROB_HT.Enabled = False
                Prob_X1axis.Enabled = False
                Prob_X2axis.Enabled = False
                Prob_Y1axis.Enabled = False
                Prob_Y2axis.Enabled = False

            End If


        ElseIf prob_support_cmb.SelectedItem = "No" Then
            prob_no.Enabled = False
            PROB_HT.Enabled = False
            Prob_X1axis.Enabled = False
            Prob_X2axis.Enabled = False
            Prob_Y1axis.Enabled = False
            Prob_Y2axis.Enabled = False

        End If

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub prob_no_SelectedIndexChanged(sender As Object, e As EventArgs) Handles prob_no.SelectedIndexChanged

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If prob_no.SelectedItem = "2 Nos" Then

            Prob_X1axis.Enabled = True
            Prob_X2axis.Enabled = True
            PROB_HT.Enabled = True


            Prob_Y1axis.Enabled = False
            Prob_Y2axis.Enabled = False

        Else prob_no.SelectedItem = "4 Nos"
            PROB_HT.Enabled = True

            Prob_X1axis.Enabled = True
            Prob_X2axis.Enabled = True

            Prob_Y1axis.Enabled = True
            Prob_Y2axis.Enabled = True
        End If
    End Sub

    Private Sub BS1_WDH_1_MouseClick(sender As Object, e As MouseEventArgs) Handles BS1_WDH_1.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub BS2_WDH_1_MouseClick(sender As Object, e As MouseEventArgs) Handles BS2_WDH_1.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub



    Private Sub TextBox12_MouseClick(sender As Object, e As MouseEventArgs) Handles X.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Distance_A_MouseClick(sender As Object, e As MouseEventArgs) Handles Distance_A.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Distance_B_MouseClick(sender As Object, e As MouseEventArgs) Handles Distance_B.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub TextBox10_MouseClick(sender As Object, e As MouseEventArgs) Handles BS_L.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub TextBox11_MouseClick(sender As Object, e As MouseEventArgs) Handles BS_R.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub MP_len_MouseClick(sender As Object, e As MouseEventArgs) Handles MP_len.MouseClick

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If


        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_L_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub MP_Wth_MouseClick(sender As Object, e As MouseEventArgs) Handles MP_Wth.MouseClick
        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Dim_A_MouseClick(sender As Object, e As MouseEventArgs) Handles Dim_L1.MouseClick
        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Prob_X1axis_MouseClick(sender As Object, e As MouseEventArgs) Handles Prob_X1axis.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Prob_X2axis_MouseClick(sender As Object, e As MouseEventArgs) Handles Prob_X2axis.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub Prob_Y1axis_MouseClick(sender As Object, e As MouseEventArgs) Handles Prob_Y1axis.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Prob_Y2axis_MouseClick(sender As Object, e As MouseEventArgs) Handles Prob_Y2axis.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        prob_H.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub HS_CC_len_MouseClick(sender As Object, e As MouseEventArgs) Handles HS_CC_len.MouseClick
        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

    End Sub

    Private Sub jack_pad_type_SelectedIndexChanged(sender As Object, e As EventArgs) Handles jack_pad_type.SelectedIndexChanged

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If
    End Sub

    Private Sub JP_A_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_A.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


    End Sub

    Private Sub JP_B_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_B.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_CH_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_B1.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub TextBox13_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_BB1.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub TextBox14_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_B2.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_C_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_C.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_D_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_D.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_E_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_E.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_G_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_G.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_H_MouseClick(sender As Object, e As MouseEventArgs)

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_T_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If jack_pad_type.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t1

            JP_HV_A.Location = New System.Drawing.Point(244, 93)
            JP_HV_B.Location = New System.Drawing.Point(83, 38)
            JP_HV_C.Location = New System.Drawing.Point(204, 131)
            JP_HV_C1.Location = New System.Drawing.Point(281, 132)
            JP_HV_D.Location = New System.Drawing.Point(243, 132)
            JP_HV_E.Location = New System.Drawing.Point(34, 93)
            JP_HV_E1.Location = New System.Drawing.Point(81, 124)
            JP_HV_G.Location = New System.Drawing.Point(83, 90)
            JP_HV_X.Location = New System.Drawing.Point(129, 102)
            JP_HV_T.Location = New System.Drawing.Point(17, 59)
            JP_HV_500.Location = New System.Drawing.Point(6, 123)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()

            JP_HV_B1.Hide()
            JP_HV_BB1.Hide()
            JP_HV_50.Hide()
            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = False
            JP_BB1.Enabled = False
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t2

            JP_HV_A.Location = New System.Drawing.Point(244, 106)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 48)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 48)
            JP_HV_C.Location = New System.Drawing.Point(204, 143)
            JP_HV_C1.Location = New System.Drawing.Point(281, 142)
            JP_HV_D.Location = New System.Drawing.Point(243, 142)
            JP_HV_E.Location = New System.Drawing.Point(28, 102)
            JP_HV_E1.Location = New System.Drawing.Point(78, 136)
            JP_HV_G.Location = New System.Drawing.Point(78, 99)
            JP_HV_X.Location = New System.Drawing.Point(126, 112)
            JP_HV_T.Location = New System.Drawing.Point(9, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(28, 60)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()

            JP_HV_B2.Hide()
            JP_HV_R100.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = False

        ElseIf jack_pad_type.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX.Image = My.Resources.jacking_pad_hv_t3


            JP_HV_A.Location = New System.Drawing.Point(244, 102)
            JP_HV_B.Location = New System.Drawing.Point(81, 28)
            JP_HV_B1.Location = New System.Drawing.Point(76, 45)
            JP_HV_B2.Location = New System.Drawing.Point(94, 179)
            JP_HV_BB1.Location = New System.Drawing.Point(124, 46)
            JP_HV_C.Location = New System.Drawing.Point(204, 166)
            JP_HV_C1.Location = New System.Drawing.Point(281, 165)
            JP_HV_D.Location = New System.Drawing.Point(243, 165)
            JP_HV_E.Location = New System.Drawing.Point(29, 98)
            JP_HV_G.Location = New System.Drawing.Point(75, 95)
            JP_HV_X.Location = New System.Drawing.Point(127, 111)
            JP_HV_T.Location = New System.Drawing.Point(13, 67)
            JP_HV_500.Location = New System.Drawing.Point(5, 136)
            JP_HV_50.Location = New System.Drawing.Point(32, 60)
            JP_HV_R100.Location = New System.Drawing.Point(24, 166)

            JP_HV_A.Show()
            JP_HV_B.Show()
            JP_HV_C.Show()
            JP_HV_C1.Show()
            JP_HV_D.Show()
            JP_HV_E.Show()
            'JP_HV_E1.Show()
            JP_HV_G.Show()
            JP_HV_X.Show()
            JP_HV_T.Show()
            JP_HV_500.Show()
            JP_HV_B1.Show()
            JP_HV_BB1.Show()
            JP_HV_50.Show()
            JP_HV_B2.Show()
            JP_HV_R100.Show()

            JP_HV_E1.Hide()

            JP_B1.Enabled = True
            JP_BB1.Enabled = True
            JP_B2.Enabled = True
        End If

        JP_HV_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_BB1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_B2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_C1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_E1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_G.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_X.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_T.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_HV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_HV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub jack_pad_type1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles jack_pad_type1.SelectedIndexChanged

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If
    End Sub

    Private Sub JP_LV_A_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_A.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub JP_LV_B_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_B.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_C_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_C.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_D_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_D.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()
        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_E_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_E.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()
        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_G_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_G.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()
        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_X_MouseClick(sender As Object, e As MouseEventArgs)

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()
        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_T_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_T.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_B1_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_B1.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_BB1_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_BB1.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub JP_LV_B2_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_LV_B2.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If jack_pad_type1.SelectedItem = "TYPE 1" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t1

            A_JP_LV.Location = New System.Drawing.Point(244, 93)
            B_JP_LV.Location = New System.Drawing.Point(83, 38)
            C_JP_LV.Location = New System.Drawing.Point(204, 131)
            C1_JP_LV.Location = New System.Drawing.Point(281, 132)
            D_JP_LV.Location = New System.Drawing.Point(243, 132)
            E_JP_LV.Location = New System.Drawing.Point(34, 93)
            E1_JP_LV.Location = New System.Drawing.Point(81, 124)
            G_JP_LV.Location = New System.Drawing.Point(83, 90)
            X_JP_LV.Location = New System.Drawing.Point(129, 102)
            T_JP_LV.Location = New System.Drawing.Point(17, 59)
            JP_LV_500.Location = New System.Drawing.Point(6, 123)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            G_JP_LV.Show()

            B1_JP_LV.Hide()
            BB1_JP_LV.Hide()
            JP_LV_50.Hide()
            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 2" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t2

            A_JP_LV.Location = New System.Drawing.Point(244, 106)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 48)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 48)
            C_JP_LV.Location = New System.Drawing.Point(204, 143)
            C1_JP_LV.Location = New System.Drawing.Point(281, 142)
            D_JP_LV.Location = New System.Drawing.Point(243, 142)
            E_JP_LV.Location = New System.Drawing.Point(28, 102)
            E1_JP_LV.Location = New System.Drawing.Point(78, 136)
            G_JP_LV.Location = New System.Drawing.Point(78, 99)
            X_JP_LV.Location = New System.Drawing.Point(126, 112)
            T_JP_LV.Location = New System.Drawing.Point(9, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(28, 60)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            E1_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()

            B2_JP_LV.Hide()
            JP_LV_R100.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = False

        ElseIf jack_pad_type1.SelectedItem = "TYPE 3" Then

            JP_PIC_BOX1.Image = My.Resources.jacking_pad_hv_t3


            A_JP_LV.Location = New System.Drawing.Point(244, 102)
            B_JP_LV.Location = New System.Drawing.Point(81, 28)
            B1_JP_LV.Location = New System.Drawing.Point(76, 45)
            B2_JP_LV.Location = New System.Drawing.Point(94, 179)
            BB1_JP_LV.Location = New System.Drawing.Point(124, 46)
            C_JP_LV.Location = New System.Drawing.Point(204, 166)
            C1_JP_LV.Location = New System.Drawing.Point(281, 165)
            D_JP_LV.Location = New System.Drawing.Point(243, 165)
            E_JP_LV.Location = New System.Drawing.Point(29, 98)
            G_JP_LV.Location = New System.Drawing.Point(75, 95)
            X_JP_LV.Location = New System.Drawing.Point(127, 111)
            T_JP_LV.Location = New System.Drawing.Point(13, 67)
            JP_LV_500.Location = New System.Drawing.Point(5, 136)
            JP_LV_50.Location = New System.Drawing.Point(32, 60)
            JP_LV_R100.Location = New System.Drawing.Point(24, 166)

            A_JP_LV.Show()
            B_JP_LV.Show()
            C_JP_LV.Show()
            C1_JP_LV.Show()
            D_JP_LV.Show()
            E_JP_LV.Show()
            G_JP_LV.Show()
            X_JP_LV.Show()
            T_JP_LV.Show()
            JP_LV_500.Show()
            B1_JP_LV.Show()
            BB1_JP_LV.Show()
            JP_LV_50.Show()
            B2_JP_LV.Show()
            JP_LV_R100.Show()

            E1_JP_LV.Hide()

            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True
        End If

        A_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        BB1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        B2_JP_LV.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        C_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        C1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        D_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        E1_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        G_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        X_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        T_JP_LV.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_50.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_LV_R100.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub jack_pad_HV_CheckedChanged(sender As Object, e As EventArgs)

        If jack_pad_HV.Checked Then

            jack_pad_type1.SelectedItem = jack_pad_type.SelectedItem
            JP_LV_A.Text = JP_A.Text
            JP_LV_B.Text = JP_B.Text
            JP_LV_C.Text = JP_C.Text
            JP_LV_D.Text = JP_D.Text
            JP_LV_E.Text = JP_E.Text
            JP_LV_G.Text = JP_G.Text

            JP_LV_T.Text = JP_T.Text
            JP_LV_B1.Text = JP_B1.Text
            JP_LV_BB1.Text = JP_BB1.Text
            JP_LV_B2.Text = JP_B2.Text

            NO_TOP_JP_LV.SelectedItem = NO_TOP_JP_HV.SelectedItem
            JP_T_L_LV.Text = JP_T_L.Text
            JP_T_T_LV.Text = JP_T_T.Text
            JP_T_H_LV.Text = JP_T_H.Text

            NO_JP_LV.SelectedItem = NO_OF_JP_HV.SelectedItem
            LV_JP1.Text = HV_JPD1.Text
            LV_JP2.Text = HV_JPD2.Text
            LV_JP3.Text = HV_JPD3.Text
            LV_JP4.Text = HV_JPD4.Text
            JP_Y_DIST_LV.Text = JP_Y_DIST.Text

            jack_pad_type1.Enabled = False
            JP_LV_A.Enabled = False
            JP_LV_B.Enabled = False
            JP_LV_C.Enabled = False
            JP_LV_D.Enabled = False
            JP_LV_E.Enabled = False
            JP_LV_G.Enabled = False

            JP_LV_T.Enabled = False
            JP_LV_B1.Enabled = False
            JP_LV_BB1.Enabled = False
            JP_LV_B2.Enabled = False

            NO_TOP_JP_LV.Enabled = False
            JP_T_L_LV.Enabled = False
            JP_T_T_LV.Enabled = False
            JP_T_H_LV.Enabled = False

            NO_JP_LV.Enabled = False
            LV_JP1.Enabled = False
            LV_JP2.Enabled = False
            LV_JP3.Enabled = False
            LV_JP4.Enabled = False
            JP_Y_DIST_LV.Enabled = False

        ElseIf Not jack_pad_HV.Checked Then

            jack_pad_type1.Enabled = True
            JP_LV_A.Enabled = True
            JP_LV_B.Enabled = True
            JP_LV_C.Enabled = True
            JP_LV_D.Enabled = True
            JP_LV_E.Enabled = True
            JP_LV_G.Enabled = True

            JP_LV_T.Enabled = True
            JP_LV_B1.Enabled = True
            JP_LV_BB1.Enabled = True
            JP_LV_B2.Enabled = True

            NO_TOP_JP_LV.Enabled = True
            JP_T_L_LV.Enabled = True
            JP_T_T_LV.Enabled = True
            JP_T_H_LV.Enabled = True

            NO_JP_LV.Enabled = True
            LV_JP1.Enabled = True
            LV_JP2.Enabled = True
            LV_JP3.Enabled = True
            LV_JP4.Enabled = True
            JP_Y_DIST_LV.Enabled = True

        End If
    End Sub

    Private Sub HV_JPD1_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_JPD1.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If NO_OF_JP_HV.SelectedItem = "2" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_2

            jp_hv_x1.Location = New System.Drawing.Point(228, 452)
            jp_hv_x2.Location = New System.Drawing.Point(348, 452)
            jp_h_500.Location = New System.Drawing.Point(224, 511)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_3

            jp_hv_x1.Location = New System.Drawing.Point(238, 468)
            jp_hv_x2.Location = New System.Drawing.Point(357, 467)
            jp_h_500.Location = New System.Drawing.Point(232, 529)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Show()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_4

            jp_hv_x1.Location = New System.Drawing.Point(167, 449)
            jp_hv_x2.Location = New System.Drawing.Point(406, 449)
            jp_hv_x3.Location = New System.Drawing.Point(236, 481)
            jp_hv_x4.Location = New System.Drawing.Point(341, 481)
            jp_h_500.Location = New System.Drawing.Point(601, 542)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_hv_x3.Show()
            jp_hv_x4.Show()
            jp_h_500.Show()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        End If

    End Sub

    Private Sub HV_JPD2_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_JPD2.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If NO_OF_JP_HV.SelectedItem = "2" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_2

            jp_hv_x1.Location = New System.Drawing.Point(228, 452)
            jp_hv_x2.Location = New System.Drawing.Point(348, 452)
            jp_h_500.Location = New System.Drawing.Point(224, 511)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_3

            jp_hv_x1.Location = New System.Drawing.Point(238, 468)
            jp_hv_x2.Location = New System.Drawing.Point(357, 467)
            jp_h_500.Location = New System.Drawing.Point(232, 529)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Show()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_4

            jp_hv_x1.Location = New System.Drawing.Point(167, 449)
            jp_hv_x2.Location = New System.Drawing.Point(406, 449)
            jp_hv_x3.Location = New System.Drawing.Point(236, 481)
            jp_hv_x4.Location = New System.Drawing.Point(341, 481)
            jp_h_500.Location = New System.Drawing.Point(601, 542)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_hv_x3.Show()
            jp_hv_x4.Show()
            jp_h_500.Show()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub HV_JPD3_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_JPD3.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If NO_OF_JP_HV.SelectedItem = "2" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_2

            jp_hv_x1.Location = New System.Drawing.Point(228, 452)
            jp_hv_x2.Location = New System.Drawing.Point(348, 452)
            jp_h_500.Location = New System.Drawing.Point(224, 511)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_3

            jp_hv_x1.Location = New System.Drawing.Point(238, 468)
            jp_hv_x2.Location = New System.Drawing.Point(357, 467)
            jp_hv_x3.Location = New System.Drawing.Point(304, 572)
            jp_h_500.Location = New System.Drawing.Point(232, 529)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Show()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False
            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_4

            jp_hv_x1.Location = New System.Drawing.Point(167, 449)
            jp_hv_x2.Location = New System.Drawing.Point(406, 449)
            jp_hv_x3.Location = New System.Drawing.Point(236, 481)
            jp_hv_x4.Location = New System.Drawing.Point(341, 481)
            jp_h_500.Location = New System.Drawing.Point(601, 542)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_hv_x3.Show()
            jp_hv_x4.Show()
            jp_h_500.Show()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        End If

    End Sub

    Private Sub HV_JPD4_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_JPD4.MouseClick

        JP_T_L_I.Hide()
        JP_T_H_I.Hide()
        JP_T_T_I.Hide()

        If NO_OF_JP_HV.SelectedItem = "2" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_2

            jp_hv_x1.Location = New System.Drawing.Point(228, 452)
            jp_hv_x2.Location = New System.Drawing.Point(348, 452)
            jp_h_500.Location = New System.Drawing.Point(224, 511)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "3" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_3

            jp_hv_x1.Location = New System.Drawing.Point(238, 468)
            jp_hv_x2.Location = New System.Drawing.Point(357, 467)
            jp_h_500.Location = New System.Drawing.Point(232, 529)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_h_500.Show()

            jp_hv_x3.Hide()
            jp_hv_x4.Hide()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_OF_JP_HV.SelectedItem = "4" Then

            JP_PIC_BOX_2.Image = My.Resources.jacking_pad_4

            jp_hv_x1.Location = New System.Drawing.Point(167, 449)
            jp_hv_x2.Location = New System.Drawing.Point(406, 449)
            jp_hv_x3.Location = New System.Drawing.Point(236, 481)
            jp_hv_x4.Location = New System.Drawing.Point(341, 481)
            jp_h_500.Location = New System.Drawing.Point(601, 542)

            jp_hv_x1.Show()
            jp_hv_x2.Show()
            jp_hv_x3.Show()
            jp_hv_x4.Show()
            jp_h_500.Show()

            JP_Y_DIST.Enabled = False

            jp_hv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_hv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_h_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If
    End Sub

    Private Sub LV_JP1_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_JP1.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If NO_JP_LV.SelectedItem = "2" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_2

            jp_lv_x1.Location = New System.Drawing.Point(228, 452)
            jp_lv_x2.Location = New System.Drawing.Point(348, 452)
            jp_l_500.Location = New System.Drawing.Point(224, 511)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "3" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_3

            jp_lv_x1.Location = New System.Drawing.Point(238, 468)
            jp_lv_x2.Location = New System.Drawing.Point(357, 467)
            jp_l_500.Location = New System.Drawing.Point(232, 529)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Show()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "4" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_4

            jp_lv_x1.Location = New System.Drawing.Point(167, 449)
            jp_lv_x2.Location = New System.Drawing.Point(406, 449)
            jp_lv_x3.Location = New System.Drawing.Point(236, 481)
            jp_lv_x4.Location = New System.Drawing.Point(341, 481)
            jp_l_500.Location = New System.Drawing.Point(601, 542)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_lv_x3.Show()
            jp_lv_x4.Show()
            jp_l_500.Show()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_JP2_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_JP2.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If NO_JP_LV.SelectedItem = "2" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_2

            jp_lv_x1.Location = New System.Drawing.Point(228, 452)
            jp_lv_x2.Location = New System.Drawing.Point(348, 452)
            jp_l_500.Location = New System.Drawing.Point(224, 511)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "3" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_3

            jp_lv_x1.Location = New System.Drawing.Point(238, 468)
            jp_lv_x2.Location = New System.Drawing.Point(357, 467)
            jp_l_500.Location = New System.Drawing.Point(232, 529)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Show()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "4" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_4

            jp_lv_x1.Location = New System.Drawing.Point(167, 449)
            jp_lv_x2.Location = New System.Drawing.Point(406, 449)
            jp_lv_x3.Location = New System.Drawing.Point(236, 481)
            jp_lv_x4.Location = New System.Drawing.Point(341, 481)
            jp_l_500.Location = New System.Drawing.Point(601, 542)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_lv_x3.Show()
            jp_lv_x4.Show()
            jp_l_500.Show()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_JP3_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_JP3.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If NO_JP_LV.SelectedItem = "2" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_2

            jp_lv_x1.Location = New System.Drawing.Point(228, 452)
            jp_lv_x2.Location = New System.Drawing.Point(348, 452)

            jp_l_500.Location = New System.Drawing.Point(224, 511)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "3" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_3

            jp_lv_x1.Location = New System.Drawing.Point(238, 468)
            jp_lv_x2.Location = New System.Drawing.Point(357, 467)
            jp_l_500.Location = New System.Drawing.Point(232, 529)
            jp_lv_x3.Location = New System.Drawing.Point(304, 572)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Show()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "4" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_4

            jp_lv_x1.Location = New System.Drawing.Point(167, 449)
            jp_lv_x2.Location = New System.Drawing.Point(406, 449)
            jp_lv_x3.Location = New System.Drawing.Point(236, 481)
            jp_lv_x4.Location = New System.Drawing.Point(341, 481)
            jp_l_500.Location = New System.Drawing.Point(601, 542)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_lv_x3.Show()
            jp_lv_x4.Show()
            jp_l_500.Show()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If
    End Sub

    Private Sub LV_JP4_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_JP4.MouseClick

        JP_T_L_LV_I.Hide()
        JP_T_H_LV_I.Hide()
        JP_T_T_LV_I.Hide()

        If NO_JP_LV.SelectedItem = "2" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_2

            jp_lv_x1.Location = New System.Drawing.Point(228, 452)
            jp_lv_x2.Location = New System.Drawing.Point(348, 452)
            jp_l_500.Location = New System.Drawing.Point(224, 511)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "3" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_3

            jp_lv_x1.Location = New System.Drawing.Point(238, 468)
            jp_lv_x2.Location = New System.Drawing.Point(357, 467)
            jp_l_500.Location = New System.Drawing.Point(232, 529)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_l_500.Show()

            jp_lv_x3.Hide()
            jp_lv_x4.Hide()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf NO_JP_LV.SelectedItem = "4" Then

            JP_PIC_BOX_LV.Image = My.Resources.jacking_pad_4

            jp_lv_x1.Location = New System.Drawing.Point(167, 449)
            jp_lv_x2.Location = New System.Drawing.Point(406, 449)
            jp_lv_x3.Location = New System.Drawing.Point(236, 481)
            jp_lv_x4.Location = New System.Drawing.Point(341, 481)
            jp_l_500.Location = New System.Drawing.Point(601, 542)

            jp_lv_x1.Show()
            jp_lv_x2.Show()
            jp_lv_x3.Show()
            jp_lv_x4.Show()
            jp_l_500.Show()

            JP_Y_DIST_LV.Enabled = False

            jp_lv_x1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            jp_lv_x4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            jp_l_500.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub TextBox13_MouseClick_1(sender As Object, e As MouseEventArgs) Handles PROB_HT.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        prob_H.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



    End Sub

    Private Sub OLTC_STNR_TYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OLTC_STNR_TYPE.SelectedIndexChanged
        S1.Show()
        S2.Show()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()


        End If
    End Sub

    Private Sub OLTC_STNR2_TYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OLTC_STNR2_TYPE.SelectedIndexChanged
        S1.Show()
        S2.Show()
        fs_A_2.Show()
        fs_B_2.Show()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        If OLTC_STNR2_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(319, 441)
            fs_A1.Location = New System.Drawing.Point(277, 423)
            fs_B.Location = New System.Drawing.Point(203, 553)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()


        End If

    End Sub

    Private Sub OLTC_END_STNR_TYP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OLTC_END_STNR_TYP.SelectedIndexChanged

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()


        If OLTC_END_STNR_TYP.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTCENDSIDE

            fs_A.Location = New System.Drawing.Point(263, 455)
            fs_B.Location = New System.Drawing.Point(363, 555)

            fs_B1.Location = New System.Drawing.Point(167, 535)
            fs_A1.Location = New System.Drawing.Point(277, 423)

            fs_A.Show()
            fs_A1.Hide()
            fs_B.Show()
            fs_B1.Hide()

            fs_A_2.Hide()
            fs_B_2.Hide()
            fs_A_2.Location = New System.Drawing.Point(319, 441)
            fs_B_2.Location = New System.Drawing.Point(277, 423)

        End If


    End Sub

    Private Sub HV_FS_A_MouseClick_1(sender As Object, e As MouseEventArgs) Handles HV_FS_A.MouseClick

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If




    End Sub

    Private Sub HV_FS_B_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_FS_B.MouseClick

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()

        T2PIC.Hide()
        T1PIC.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If
    End Sub

    Private Sub HV_FS_A1_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_FS_A1.MouseClick

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()

        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If

    End Sub

    Private Sub HV_FS_B1_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_FS_B1.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()

        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If

    End Sub

    Private Sub LV_FS_A_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_FS_A.MouseClick


        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If

    End Sub

    Private Sub LV_FS_B_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_FS_B.MouseClick

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If
    End Sub

    Private Sub LV_FS_A1_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_FS_A1.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If

    End Sub

    Private Sub LV_FS_B1_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_FS_B1.MouseClick

        Load_OLTC_Pkt_Position()

        T2PIC.Hide()
        T1PIC.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()

        S1_PIC.Hide()
        S2_PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If btm_tk_type.SelectedItem = "Type 1" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If LV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf LV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf LV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If


    End Sub

    Private Sub OLTC_S1_A_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S1_A.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()

        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        S1.Show()
        S2.Show()
        fs_A.Show()
        fs_B.Show()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()


        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Hide()
            fs_oltc_D2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        End If
    End Sub

    Private Sub OLTC_S1_B_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S1_B.MouseClick

        Load_OLTC_Pkt_Position()

        S1_PIC.Hide()
        S2_PIC.Hide()

        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        S1.Show()
        S2.Show()
        fs_A.Show()
        fs_B.Show()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE
            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Hide()
            fs_oltc_D2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        End If
    End Sub

    Private Sub OLTC_S1_A1_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S1_A1.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        S1.Show()
        S2.Show()
        fs_A.Show()
        fs_B.Show()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()


        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE
            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Hide()
            fs_oltc_D2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        End If
    End Sub

    Private Sub OLTC_S1_B1_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S1_B1.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        S1.Show()
        S2.Show()
        fs_A.Show()
        fs_B.Show()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        fs_oltc_D1.Show()
        fs_oltc_D2.Show()

        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            'fs_B1.Location = New System.Drawing.Point(167, 535)
            fs_oltc_D1.Location = New System.Drawing.Point(171, 545)
            fs_oltc_D2.Location = New System.Drawing.Point(188, 528)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Show()
            fs_oltc_D2.Show()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            'fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_oltc_D1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_oltc_D2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)



        End If
    End Sub

    Private Sub OLTC_S2_A_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S2_A.MouseClick


        Load_OLTC_Pkt_Position()

        S1.Show()
        S2.Show()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_A_2.Show()
        fs_B_2.Show()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()

        If OLTC_STNR2_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_B_2.Location = New System.Drawing.Point(562, 522)
            fs_A_2.Location = New System.Drawing.Point(531, 446)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            fs_A_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_B_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If
    End Sub

    Private Sub OLTC_S2_B_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S2_B.MouseClick


        Load_OLTC_Pkt_Position()

        S1.Show()
        S2.Show()
        fs_A_2.Show()
        fs_B_2.Show()
        S1_PIC.Hide()
        S2_PIC.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If OLTC_STNR2_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_B_2.Location = New System.Drawing.Point(562, 522)
            fs_A_2.Location = New System.Drawing.Point(531, 446)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            fs_A_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B_2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        End If

    End Sub

    Private Sub OLTC_S2_A1_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S2_A1.MouseClick

        Load_OLTC_Pkt_Position()


        S1.Show()
        S2.Show()
        fs_A_2.Show()
        fs_B_2.Show()
        S1_PIC.Hide()
        S2_PIC.Hide()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(639, 459)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        If OLTC_STNR2_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_B_2.Location = New System.Drawing.Point(562, 522)
            fs_A_2.Location = New System.Drawing.Point(531, 446)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Hide()
            fs_oltc_D2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            fs_A_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub


    Private Sub OLTC_END_A_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_END_A.MouseClick

        Load_OLTC_Pkt_Position()


        If OLTC_END_STNR_TYP.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTCENDSIDE

            fs_A.Location = New System.Drawing.Point(313, 446)
            fs_B.Location = New System.Drawing.Point(363, 555)


            fs_A1.Location = New System.Drawing.Point(277, 423)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_oltc_T1.Location = New System.Drawing.Point(207, 464)

            T2PIC.Hide()
            T1PIC.Show()
            T1PIC.Location = New System.Drawing.Point(165, 460)
            S1.Hide()
            S2.Hide()
            fs_A.Show()
            fs_A1.Hide()
            fs_B.Show()
            fs_B1.Hide()
            fs_A_2.Hide()
            fs_B_2.Hide()
            S1_PIC.Hide()
            S2_PIC.Hide()
            OD1.Hide()
            OD2.Hide()
            OD3.Hide()
            OD4.Hide()

            fs_oltc_T1.Show()
            fs_oltc_T2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        End If
    End Sub

    Private Sub OLTC_END_B_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_END_B.MouseClick

        Load_OLTC_Pkt_Position()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        S1.Hide()
        S2.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(165, 460)
        If OLTC_END_STNR_TYP.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTCENDSIDE

            fs_A.Location = New System.Drawing.Point(313, 446)
            fs_B.Location = New System.Drawing.Point(363, 555)

            fs_A1.Location = New System.Drawing.Point(277, 423)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_oltc_T1.Location = New System.Drawing.Point(207, 464)


            fs_A.Show()
            fs_A1.Hide()
            fs_B.Show()
            fs_B1.Hide()
            fs_A_2.Hide()
            fs_B_2.Hide()

            fs_oltc_T1.Show()
            fs_oltc_T2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        End If
    End Sub

    Private Sub OLTC_END_A1_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_END_A1.MouseClick

        Load_OLTC_Pkt_Position()


        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()
        S1.Hide()
        S2.Hide()
        T2PIC.Hide()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(165, 460)
        S1_PIC.Hide()
        S2_PIC.Hide()
        If OLTC_END_STNR_TYP.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTCENDSIDE

            fs_A.Location = New System.Drawing.Point(313, 446)
            fs_B.Location = New System.Drawing.Point(363, 555)

            fs_A1.Location = New System.Drawing.Point(277, 423)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_oltc_T1.Location = New System.Drawing.Point(207, 464)


            fs_A.Show()
            fs_A1.Hide()
            fs_B.Show()
            fs_B1.Hide()
            fs_A_2.Hide()
            fs_B_2.Hide()

            fs_oltc_T1.Show()
            fs_oltc_T2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


        End If

    End Sub

    Private Sub OLTC_END_B1_MouseClick(sender As Object, e As MouseEventArgs)

        If OLTC_END_STNR_TYP.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

            fs_A.Location = New System.Drawing.Point(319, 441)
            fs_A1.Location = New System.Drawing.Point(277, 423)
            fs_B.Location = New System.Drawing.Point(203, 553)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Show()


            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


        ElseIf OLTC_END_STNR_TYP.SelectedItem = "Type 2" Then

            STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

            fs_A.Location = New System.Drawing.Point(344, 441)
            fs_A1.Location = New System.Drawing.Point(316, 420)
            fs_B.Location = New System.Drawing.Point(213, 534)
            fs_B1.Location = New System.Drawing.Point(188, 532)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Show()


            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)


        ElseIf OLTC_END_STNR_TYP.SelectedItem = "Type 3" Then

            STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

            fs_A.Location = New System.Drawing.Point(332, 440)
            fs_A1.Location = New System.Drawing.Point(293, 421)
            fs_B.Location = New System.Drawing.Point(213, 534)
            fs_B1.Location = New System.Drawing.Point(174, 512)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Show()



            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        End If
    End Sub

    Private Sub CC_LENGTH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CC_LENGTH.SelectedIndexChanged

        If CC_LENGTH.SelectedItem = "Custom" Then
            Roller_Custom.Enabled = True

            HS_CC_len.Text = Roller_Custom.Text
        Else
            Roller_Custom.Enabled = False
            HS_CC_len.Text = CC_LENGTH.SelectedItem

        End If

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub


    Private Sub Roller_Custom_TextChanged(sender As Object, e As EventArgs) Handles Roller_Custom.TextChanged
        HS_CC_len.Text = Roller_Custom.Text

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub NO_TOP_JP_HV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NO_TOP_JP_HV.SelectedIndexChanged
        JP_HV_B.Hide()
        JP_HV_B1.Hide()
        JP_HV_BB1.Hide()
        JP_HV_50.Hide()
        JP_HV_T.Hide()
        JP_HV_E.Hide()
        JP_HV_E1.Hide()
        JP_HV_G.Hide()
        JP_HV_X.Hide()
        JP_HV_500.Hide()
        JP_HV_R100.Hide()
        JP_HV_B2.Hide()
        JP_HV_A.Hide()
        JP_HV_C.Hide()
        JP_HV_D.Hide()
        JP_HV_C1.Hide()
    End Sub

    Private Sub OS1D1_MouseClick(sender As Object, e As MouseEventArgs) Handles OS1D1.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Show()
        S2_PIC.Show()
        T2PIC.Hide()
        T1PIC.Hide()
        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()
        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        If OLTC_pkt_position = "Left" Then


            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2L

            OD1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD1.Location = New System.Drawing.Point(189, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(113, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(459, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(531, 607)

        ElseIf OLTC_pkt_position = "Right" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2R

            OD1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD1.Location = New System.Drawing.Point(456, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(531, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(189, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(107, 607)

        End If








    End Sub

    Private Sub OS1D2_MouseClick(sender As Object, e As MouseEventArgs) Handles OS1D2.MouseClick
        Load_OLTC_Pkt_Position()


        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        If OLTC_pkt_position = "Left" Then


            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2L



            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(189, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD2.Location = New System.Drawing.Point(113, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(459, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(531, 607)


        ElseIf OLTC_pkt_position = "Right" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2R

            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(456, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD2.Location = New System.Drawing.Point(531, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(189, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(107, 607)

        End If





    End Sub





    Private Sub OS2D1_MouseClick(sender As Object, e As MouseEventArgs) Handles OS2D1.MouseClick

        Load_OLTC_Pkt_Position()


        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        If OLTC_pkt_position = "Left" Then


            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2L

            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(189, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(113, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD3.Location = New System.Drawing.Point(459, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(531, 607)


        ElseIf OLTC_pkt_position = "Right" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2R

            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(456, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(531, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD3.Location = New System.Drawing.Point(189, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD4.Location = New System.Drawing.Point(107, 607)

        End If





    End Sub

    Private Sub OS2D2_MouseClick(sender As Object, e As MouseEventArgs) Handles OS2D2.MouseClick

        Load_OLTC_Pkt_Position()

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        If OLTC_pkt_position = "Left" Then


            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2L

            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(189, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(113, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(459, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD4.Location = New System.Drawing.Point(531, 607)

        ElseIf OLTC_pkt_position = "Right" Then

            STIFNER_PIC_BOX.Image = My.Resources.OLTC_S1S2R

            OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD1.Location = New System.Drawing.Point(456, 566)

            OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD2.Location = New System.Drawing.Point(531, 607)


            OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            OD3.Location = New System.Drawing.Point(189, 573)

            OD4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            OD4.Location = New System.Drawing.Point(107, 607)

        End If

    End Sub

    Private Sub OLTC_S1_T_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_S1_T.MouseClick

        Load_OLTC_Pkt_Position()


        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Show()
        T1PIC.Show()
        T1PIC.Location = New System.Drawing.Point(8, 459)
        fs_oltc_T1.Show()
        fs_oltc_T1.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        fs_oltc_T2.Show()
        fs_oltc_T2.Location = New System.Drawing.Point(51, 463)
        fs_oltc_T2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        S1.Show()
        S2.Show()
        fs_A.Show()
        fs_B.Show()
        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()

        If OLTC_STNR_TYPE.SelectedItem = "Type 1" Then

            STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_OLTCSIDE

            fs_A.Location = New System.Drawing.Point(223, 455)
            fs_A1.Location = New System.Drawing.Point(345, 507)
            fs_B.Location = New System.Drawing.Point(102, 546)
            fs_B1.Location = New System.Drawing.Point(167, 535)

            fs_A.Show()
            fs_A1.Show()
            fs_B.Show()
            fs_B1.Hide()
            fs_oltc_D1.Hide()
            fs_oltc_D2.Hide()

            fs_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)



        End If
    End Sub

    Private Sub OED1_MouseClick(sender As Object, e As MouseEventArgs) Handles OED1.MouseClick

        Load_OLTC_Pkt_Position()


        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()




        STIFNER_PIC_BOX.Image = My.Resources.OLTCES1

        OD1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD1.Location = New System.Drawing.Point(297, 588)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(340, 588)


        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(255, 630)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(388, 629)



    End Sub

    Private Sub OED2_MouseClick(sender As Object, e As MouseEventArgs) Handles OED2.MouseClick

        Load_OLTC_Pkt_Position()


        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()




        STIFNER_PIC_BOX.Image = My.Resources.OLTCES1

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(297, 588)

        OD2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD2.Location = New System.Drawing.Point(340, 588)


        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(255, 630)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(388, 629)

    End Sub

    Private Sub OED3_MouseClick(sender As Object, e As MouseEventArgs) Handles OED3.MouseClick

        Load_OLTC_Pkt_Position()


        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()




        STIFNER_PIC_BOX.Image = My.Resources.OLTCES1

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(297, 588)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(340, 588)


        OD3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD3.Location = New System.Drawing.Point(255, 630)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(388, 629)

    End Sub

    Private Sub OED4_MouseClick(sender As Object, e As MouseEventArgs) Handles OED4.MouseClick

        Load_OLTC_Pkt_Position()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()
        S1_PIC.Hide()
        S2_PIC.Hide()
        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        T2PIC.Hide()
        T1PIC.Hide()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()




        STIFNER_PIC_BOX.Image = My.Resources.OLTCES1

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(297, 588)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(340, 588)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(255, 630)

        OD4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD4.Location = New System.Drawing.Point(388, 629)

    End Sub

    Private Sub JP_T_L_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_L.MouseClick

        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1

        JP_HV_B.Hide()
        JP_HV_B1.Hide()
        JP_HV_BB1.Hide()
        JP_HV_50.Hide()
        JP_HV_T.Hide()
        JP_HV_E.Hide()
        JP_HV_E1.Hide()
        JP_HV_G.Hide()
        JP_HV_X.Hide()
        JP_HV_500.Hide()
        JP_HV_R100.Hide()
        JP_HV_B2.Hide()
        JP_HV_A.Hide()
        JP_HV_C.Hide()
        JP_HV_D.Hide()
        JP_HV_C1.Hide()

        JP_T_L_I.Show()
        JP_T_H_I.Show()
        JP_T_T_I.Show()

        JP_T_L_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_T_H_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_T_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub JP_T_H_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_H.MouseClick

        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1

        JP_HV_B.Hide()
        JP_HV_B1.Hide()
        JP_HV_BB1.Hide()
        JP_HV_50.Hide()
        JP_HV_T.Hide()
        JP_HV_E.Hide()
        JP_HV_E1.Hide()
        JP_HV_G.Hide()
        JP_HV_X.Hide()
        JP_HV_500.Hide()
        JP_HV_R100.Hide()
        JP_HV_B2.Hide()
        JP_HV_A.Hide()
        JP_HV_C.Hide()
        JP_HV_D.Hide()
        JP_HV_C1.Hide()

        JP_T_L_I.Show()
        JP_T_H_I.Show()
        JP_T_T_I.Show()

        JP_T_L_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_H_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_T_T_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub JP_T_T_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_T.MouseClick

        JP_PIC_BOX.Image = My.Resources.JP_TOP_STFNR1

        JP_HV_B.Hide()
        JP_HV_B1.Hide()
        JP_HV_BB1.Hide()
        JP_HV_50.Hide()
        JP_HV_T.Hide()
        JP_HV_E.Hide()
        JP_HV_E1.Hide()
        JP_HV_G.Hide()
        JP_HV_X.Hide()
        JP_HV_500.Hide()
        JP_HV_R100.Hide()
        JP_HV_B2.Hide()
        JP_HV_A.Hide()
        JP_HV_C.Hide()
        JP_HV_D.Hide()
        JP_HV_C1.Hide()

        JP_T_L_I.Show()
        JP_T_H_I.Show()
        JP_T_T_I.Show()

        JP_T_L_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_H_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_T_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub N2_filling_req_HV_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_filling_req_HV.SelectedIndexChanged

        If N2_filling_req_HV.SelectedItem = "Yes" Then

            N2_qty_HV.Enabled = True



            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

            If N2_qty_HV.SelectedItem = "1" Then


                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = False
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "2" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = False

            ElseIf N2_qty_HV.SelectedItem = "3" Then

                HV_X1_AXIS.Enabled = True
                HV_X2_AXIS.Enabled = True
                HV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_HV.SelectedItem = "No" Then
            N2_qty_HV.Enabled = False
            HV_X1_AXIS.Enabled = False
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

        End If
    End Sub

    Private Sub N2_filling_req_LV_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_filling_req_LV.SelectedIndexChanged
        If N2_filling_req_LV.SelectedItem = "Yes" Then

            N2_qty_LV.Enabled = True
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

            If N2_qty_LV.SelectedItem = "1" Then


                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = False
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "2" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = False

            ElseIf N2_qty_LV.SelectedItem = "3" Then

                LV_X1_AXIS.Enabled = True
                LV_X2_AXIS.Enabled = True
                LV_X3_AXIS.Enabled = True

            End If



        ElseIf N2_filling_req_LV.SelectedItem = "No" Then
            N2_qty_LV.Enabled = False
            LV_X1_AXIS.Enabled = False
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False


        End If
    End Sub

    Private Sub N2_filling_req_OLTC_opp_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_filling_req_OLTC_opp.SelectedIndexChanged
        If N2_filling_req_OLTC_opp.SelectedItem = "Yes" Then

            N2_qty_OLTC_opp.Enabled = True
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False

            If N2_qty_OLTC_opp.SelectedItem = "1" Then


                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = False
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "2" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC_opp.SelectedItem = "3" Then

                OLTC_opp_X1_AXIS.Enabled = True
                OLTC_opp_X2_AXIS.Enabled = True
                OLTC_opp_X3_AXIS.Enabled = True

            End If


        ElseIf N2_filling_req_OLTC_opp.SelectedItem = "No" Then
            N2_qty_OLTC_opp.Enabled = False
            OLTC_opp_X1_AXIS.Enabled = False
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False


        End If
    End Sub

    Private Sub N2_filling_req_OLTC_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_filling_req_OLTC.SelectedIndexChanged
        'oltc
        If N2_filling_req_OLTC.SelectedItem = "Yes" Then

            N2_qty_OLTC.Enabled = True

            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

            If N2_qty_OLTC.SelectedItem = "1" Then


                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = False
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "2" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = False

            ElseIf N2_qty_OLTC.SelectedItem = "3" Then

                OLTC_X1_AXIS.Enabled = True
                OLTC_X2_AXIS.Enabled = True
                OLTC_X3_AXIS.Enabled = True

            End If

        ElseIf N2_filling_req_OLTC.SelectedItem = "No" Then
            N2_qty_OLTC.Enabled = False
            OLTC_X1_AXIS.Enabled = False
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False


        End If
    End Sub

    Private Sub N2_qty_HV_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_qty_HV.SelectedIndexChanged

        If N2_qty_HV.SelectedItem = "1" Then

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_1Q

            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = False
            HV_X3_AXIS.Enabled = False

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_2Q

            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = True
            HV_X3_AXIS.Enabled = False

        ElseIf N2_qty_HV.SelectedItem = "3" Then

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_HV

            HV_X1_AXIS.Enabled = True
            HV_X2_AXIS.Enabled = True
            HV_X3_AXIS.Enabled = True

        End If

    End Sub

    Private Sub N2_qty_LV_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_qty_LV.SelectedIndexChanged


        If N2_qty_LV.SelectedItem = "1" Then

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_1Q

            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = False
            LV_X3_AXIS.Enabled = False

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_2Q

            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = True
            LV_X3_AXIS.Enabled = False

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_LV

            LV_X1_AXIS.Enabled = True
            LV_X2_AXIS.Enabled = True
            LV_X3_AXIS.Enabled = True

        End If

    End Sub

    Private Sub N2_qty_OLTC_opp_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_qty_OLTC_opp.SelectedIndexChanged
        If N2_qty_OLTC_opp.SelectedItem = "1" Then


            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = False
            OLTC_opp_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC_opp.SelectedItem = "2" Then

            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = True
            OLTC_opp_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC_opp.SelectedItem = "3" Then

            OLTC_opp_X1_AXIS.Enabled = True
            OLTC_opp_X2_AXIS.Enabled = True
            OLTC_opp_X3_AXIS.Enabled = True

        End If

    End Sub

    Private Sub N2_qty_OLTC_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles N2_qty_OLTC.SelectedIndexChanged

        If N2_qty_OLTC.SelectedItem = "1" Then


            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = False
            OLTC_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC.SelectedItem = "2" Then

            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = True
            OLTC_X3_AXIS.Enabled = False

        ElseIf N2_qty_OLTC.SelectedItem = "3" Then

            OLTC_X1_AXIS.Enabled = True
            OLTC_X2_AXIS.Enabled = True
            OLTC_X3_AXIS.Enabled = True

        End If
    End Sub

    Private Sub GAUDE_PIN_SIZE_MouseHover(sender As Object, e As EventArgs) Handles GAUDE_PIN_SIZE.MouseHover
        Curb_A.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_C.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Curb_D.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Xaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        GP_Yaxis_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_WDH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OLTC_PKT_LEN.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_X2_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bush_pkt_Y1_axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_thickness.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        bt_thickness1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub




    Private Sub OLTC_pkt_position_SelectedIndexChanged(sender As Object, e As EventArgs)

        Load_OLTC_Pkt_Position()

        If OLTC_pkt_position = "Left" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_R1

            ElseIf btm_tk_curb_pro.SelectedItem = "Type 2" Then
                Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2LHV 'T2_BTM_TK_LWT1
            End If



        ElseIf OLTC_pkt_position = "Right" Then

            If btm_tk_curb_pro.SelectedItem = "Type 1" Then
                Btm_TK_pic_box.Image = My.Resources.BTM_TK_CURB1_L
            ElseIf btm_tk_curb_pro.SelectedItem = "Type 2" Then
                Btm_TK_pic_box.Image = My.Resources.Btm_tk_t2RHV
            End If

        End If
    End Sub

    Private Sub TabControl1_MouseClick(sender As Object, e As MouseEventArgs) Handles TabControl1.MouseClick



        save_btm_tank_inputs()

        S1_PIC.Hide()
        S2_PIC.Hide()
        T2PIC.Hide()
        T1PIC.Hide()

        S1.Hide()
        S2.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()

        OD1.Hide()
        OD2.Hide()
        OD3.Hide()
        OD4.Hide()


        If btm_tk_type.SelectedItem = "Type 1" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_type1

                fs_A.Location = New System.Drawing.Point(319, 441)
                fs_A1.Location = New System.Drawing.Point(277, 423)
                fs_B.Location = New System.Drawing.Point(203, 553)
                fs_B1.Location = New System.Drawing.Point(167, 535)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t2

                fs_A.Location = New System.Drawing.Point(344, 441)
                fs_A1.Location = New System.Drawing.Point(316, 420)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(188, 532)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.flat_stiffner_t3

                fs_A.Location = New System.Drawing.Point(332, 440)
                fs_A1.Location = New System.Drawing.Point(293, 421)
                fs_B.Location = New System.Drawing.Point(213, 534)
                fs_B1.Location = New System.Drawing.Point(174, 512)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If



        ElseIf btm_tk_type.SelectedItem = "Type 2" Then

            If HV_STNR_TYPE.SelectedItem = "Type 1" Then

                STIFNER_PIC_BOX.Image = My.Resources.FALTSTIFFNER_HV

                fs_A.Location = New System.Drawing.Point(332, 442)
                fs_A1.Location = New System.Drawing.Point(296, 420)
                fs_B.Location = New System.Drawing.Point(193, 540)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()




            ElseIf HV_STNR_TYPE.SelectedItem = "Type 2" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T2

                fs_A.Location = New System.Drawing.Point(324, 460)
                fs_A1.Location = New System.Drawing.Point(289, 436)
                fs_B.Location = New System.Drawing.Point(158, 544)
                fs_B1.Location = New System.Drawing.Point(182, 556)

                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()

            ElseIf HV_STNR_TYPE.SelectedItem = "Type 3" Then

                STIFNER_PIC_BOX.Image = My.Resources.FLATSTIFF_T2_T3_1

                fs_A.Location = New System.Drawing.Point(338, 462)
                fs_A1.Location = New System.Drawing.Point(307, 440)
                fs_B.Location = New System.Drawing.Point(193, 539)
                fs_B1.Location = New System.Drawing.Point(215, 547)


                fs_A.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
                fs_A1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
                fs_B1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

                fs_A.Show()
                fs_A1.Show()
                fs_B.Show()
                fs_B1.Show()
            End If


        End If




    End Sub

    Private Sub OLTC_pck_len_TextChanged(sender As Object, e As EventArgs) Handles OLTC_pck_len.TextChanged
        OLTC_pck_len1 = OLTC_pck_len.Text
    End Sub

    Private Sub OLTC_pck_wdh_TextChanged(sender As Object, e As EventArgs) Handles OLTC_pck_wdh.TextChanged
        OLTC_pck_wdh1 = OLTC_pck_wdh.Text
    End Sub

    Private Sub BS1_WDH_1_TextChanged(sender As Object, e As EventArgs) Handles BS1_WDH_1.TextChanged
        BS3_WDH_1.Text = BS1_WDH_1.Text
    End Sub

    Private Sub BS1_HT_WS_TextChanged(sender As Object, e As EventArgs) Handles BS1_HT_WS.TextChanged
        BS3_HT_WS.Text = BS1_HT_WS.Text
    End Sub

    Private Sub BS1_THK_1_TextChanged(sender As Object, e As EventArgs) Handles BS1_THK_1.TextChanged
        BS3_THK_1.Text = BS1_THK_1.Text

    End Sub

    Private Sub BS1_HT_WS_MouseClick(sender As Object, e As MouseEventArgs) Handles BS1_HT_WS.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub BS1_THK_1_MouseClick(sender As Object, e As MouseEventArgs) Handles BS1_THK_1.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub BS2_HT_WS_MouseClick(sender As Object, e As MouseEventArgs) Handles BS2_HT_WS.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub BS2_THK_1_MouseClick(sender As Object, e As MouseEventArgs) Handles BS2_THK_1.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        If No_of_Rails.SelectedItem = "3" Then

            UB_picture_box.Image = My.Resources.UB_3rail

            box_stiff_w1.Location = New System.Drawing.Point(7, 363)
            box_stiff_w2.Location = New System.Drawing.Point(6, 481)
            Roller_C2_1.Location = New System.Drawing.Point(159, 592)
            Roller_C2_2.Location = New System.Drawing.Point(229, 592)
            Roller_C2_3.Location = New System.Drawing.Point(308, 592)
            Roller_C2_4.Location = New System.Drawing.Point(378, 592)
            Roller_C2_5.Location = New System.Drawing.Point(558, 433)
            Roller_C2_6.Location = New System.Drawing.Point(559, 517)
            Roller_center_2.Location = New System.Drawing.Point(584, 471)
            UB_X1.Location = New System.Drawing.Point(542, 462)
            UB_X2.Location = New System.Drawing.Point(542, 482)
            UB_D1.Location = New System.Drawing.Point(116, 592)
            UB_D2.Location = New System.Drawing.Point(462, 592)
            UB_roll_center_1.Location = New System.Drawing.Point(184, 616)
            UB_roll_center_2.Location = New System.Drawing.Point(339, 616)

            box_stiff_w1.Show()
            box_stiff_w2.Show()
            Roller_C2_1.Show()
            Roller_C2_2.Show()
            Roller_C2_3.Show()
            Roller_C2_4.Show()
            Roller_C2_5.Show()
            Roller_C2_6.Show()
            Roller_center_2.Show()
            UB_X1.Show()
            UB_X2.Show()
            UB_D1.Show()
            UB_D2.Show()
            UB_roll_center_1.Show()
            UB_roll_center_2.Show()

            UB_box_stiff1.Hide()
            UB_box_stiff2.Hide()
            Prop_supp_X1axis.Hide()
            Prop_supp_X2axis.Hide()
            Prop_supp_Y1axis.Hide()
            Prop_supp_Y2axis.Hide()
            Roller_center_1.Hide()
            UB_W.Hide()
            UB_L.Hide()
            UB_L1.Hide()

        End If


        If No_of_Rails.SelectedItem = "2" Then

            If No_of_Rails1.SelectedItem = "Option 1" Then

                UB_picture_box.Image = My.Resources.UB_opt_1

                box_stiff_w1.Location = New System.Drawing.Point(57, 358)
                box_stiff_w2.Location = New System.Drawing.Point(6, 548)
                UB_box_stiff1.Location = New System.Drawing.Point(176, 605)
                UB_box_stiff2.Location = New System.Drawing.Point(405, 605)
                Prop_supp_X1axis.Location = New System.Drawing.Point(168, 624)
                Prop_supp_X2axis.Location = New System.Drawing.Point(369, 624)
                Roller_C2_1.Location = New System.Drawing.Point(255, 571)
                Roller_C2_2.Location = New System.Drawing.Point(323, 571)
                Roller_center_1.Location = New System.Drawing.Point(284, 597)
                Roller_center_2.Location = New System.Drawing.Point(586, 467)
                UB_W.Location = New System.Drawing.Point(229, 508)
                UB_L.Location = New System.Drawing.Point(158, 548)
                UB_L1.Location = New System.Drawing.Point(177, 540)
                UB_X1.Location = New System.Drawing.Point(566, 451)
                UB_X2.Location = New System.Drawing.Point(565, 486)

                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                Prop_supp_Y1axis.Hide()
                Prop_supp_Y2axis.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            ElseIf No_of_Rails1.SelectedItem = "Option 2" Then

                UB_picture_box.Image = My.Resources.UB_opt_2

                box_stiff_w1.Location = New System.Drawing.Point(41, 361)
                box_stiff_w2.Location = New System.Drawing.Point(14, 458)
                UB_box_stiff1.Location = New System.Drawing.Point(155, 606)
                UB_box_stiff2.Location = New System.Drawing.Point(377, 606)
                Prop_supp_X1axis.Location = New System.Drawing.Point(150, 626)
                Prop_supp_X2axis.Location = New System.Drawing.Point(342, 626)
                Roller_C2_1.Location = New System.Drawing.Point(233, 571)
                Roller_C2_2.Location = New System.Drawing.Point(302, 571)
                Roller_center_1.Location = New System.Drawing.Point(261, 596)
                Roller_center_2.Location = New System.Drawing.Point(352, 458)
                UB_W.Location = New System.Drawing.Point(209, 505)
                UB_L.Location = New System.Drawing.Point(157, 546)
                UB_L1.Location = New System.Drawing.Point(177, 539)
                UB_X1.Location = New System.Drawing.Point(533, 452)
                UB_X2.Location = New System.Drawing.Point(533, 486)
                Prop_supp_Y1axis.Location = New System.Drawing.Point(537, 424)
                Prop_supp_Y2axis.Location = New System.Drawing.Point(537, 514)


                box_stiff_w1.Show()
                box_stiff_w2.Show()
                UB_box_stiff1.Show()
                UB_box_stiff2.Show()
                Prop_supp_X1axis.Show()
                Prop_supp_X2axis.Show()
                Roller_C2_1.Show()
                Roller_C2_2.Show()
                Roller_center_1.Show()
                Roller_center_2.Show()
                UB_W.Show()
                UB_L.Show()
                UB_L1.Show()
                UB_X1.Show()
                UB_X2.Show()
                Prop_supp_Y1axis.Show()
                Prop_supp_Y2axis.Show()

                Roller_C2_3.Hide()
                Roller_C2_4.Hide()
                Roller_C2_5.Hide()
                Roller_C2_6.Hide()
                UB_D1.Hide()
                UB_D2.Hide()
                UB_roll_center_1.Hide()
                UB_roll_center_2.Hide()

            End If
        End If

        box_stiff_w1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        box_stiff_w2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_box_stiff2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_X2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_W.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_X2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y1axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Prop_supp_Y2axis.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_5.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        Roller_C2_6.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_D2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_roll_center_2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub HV_STNR_TYPE_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_STNR_TYPE.MouseClick


        Load_OLTC_Pkt_Position()

    End Sub

    Private Sub LV_STNR_TYPE_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_STNR_TYPE.MouseClick

        Load_OLTC_Pkt_Position()

    End Sub

    Private Sub OLTC_STNR_TYPE_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_STNR_TYPE.MouseClick

        Load_OLTC_Pkt_Position()


    End Sub

    Private Sub OLTC_STNR2_TYPE_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_STNR2_TYPE.MouseClick

        Load_OLTC_Pkt_Position()


    End Sub

    Private Sub OLTC_END_STNR_TYP_MouseClick(sender As Object, e As MouseEventArgs) Handles OLTC_END_STNR_TYP.MouseClick

        Load_OLTC_Pkt_Position()


    End Sub

    Private Sub Stiffner_count_BT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Stiffner_count_BT.SelectedIndexChanged

        If Stiffner_count_BT.SelectedItem = "1" Then

            OED1.Enabled = True
            OED2.Enabled = False
            OED3.Enabled = False
            OED4.Enabled = False

        ElseIf Stiffner_count_BT.SelectedItem = "2" Then

            OED1.Enabled = True
            OED2.Enabled = True
            OED3.Enabled = False
            OED4.Enabled = False

        ElseIf Stiffner_count_BT.SelectedItem = "3" Then

            OED1.Enabled = True
            OED2.Enabled = True
            OED3.Enabled = True
            OED4.Enabled = False

        ElseIf Stiffner_count_BT.SelectedItem = "4" Then

            OED1.Enabled = True
            OED2.Enabled = True
            OED3.Enabled = True
            OED4.Enabled = True

        End If

    End Sub

    Private Sub oltc_stiffner1_count_SelectedIndexChanged(sender As Object, e As EventArgs) Handles oltc_stiffner1_count.SelectedIndexChanged

        If oltc_stiffner1_count.SelectedItem = "1" Then

            OS1D1.Enabled = True
            OS1D2.Enabled = False

        ElseIf oltc_stiffner1_count.SelectedItem = "2" Then

            OS1D1.Enabled = True
            OS1D2.Enabled = True

        End If

    End Sub

    Private Sub oltc_stiffner2_count_SelectedIndexChanged(sender As Object, e As EventArgs) Handles oltc_stiffner2_count.SelectedIndexChanged

        If oltc_stiffner2_count.SelectedItem = "1" Then

            OS2D1.Enabled = True
            OS2D2.Enabled = False

        ElseIf oltc_stiffner2_count.SelectedItem = "2" Then

            OS2D1.Enabled = True
            OS2D2.Enabled = True
        End If

    End Sub

    Private Sub stiffner_count_HV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles stiffner_count_HV.SelectedIndexChanged

        If stiffner_count_HV.SelectedItem = "1" Then

            HVD1.Enabled = True
            HVD2.Enabled = False
            HVD3.Enabled = False
            HVD4.Enabled = False

        ElseIf stiffner_count_HV.SelectedItem = "2" Then

            HVD1.Enabled = True
            HVD2.Enabled = True
            HVD3.Enabled = False
            HVD4.Enabled = False

        ElseIf stiffner_count_HV.SelectedItem = "3" Then

            HVD1.Enabled = True
            HVD2.Enabled = True
            HVD3.Enabled = True
            HVD4.Enabled = False

        ElseIf stiffner_count_HV.SelectedItem = "4" Then

            HVD1.Enabled = True
            HVD2.Enabled = True
            HVD3.Enabled = True
            HVD4.Enabled = True

        End If

    End Sub

    Private Sub stiffner_count_LV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles stiffner_count_LV.SelectedIndexChanged

        If stiffner_count_LV.SelectedItem = "1" Then

            LVD1.Enabled = True
            LVD2.Enabled = False
            LVD3.Enabled = False
            LVD4.Enabled = False

        ElseIf stiffner_count_LV.SelectedItem = "2" Then

            LVD1.Enabled = True
            LVD2.Enabled = True
            LVD3.Enabled = False
            LVD4.Enabled = False

        ElseIf stiffner_count_LV.SelectedItem = "3" Then

            LVD1.Enabled = True
            LVD2.Enabled = True
            LVD3.Enabled = True
            LVD4.Enabled = False

        ElseIf stiffner_count_LV.SelectedItem = "4" Then

            LVD1.Enabled = True
            LVD2.Enabled = True
            LVD3.Enabled = True
            LVD4.Enabled = True

        End If

    End Sub

    Private Sub SV_XAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles SV_XAXIS.MouseClick

        If SAM_VAL_POS.SelectedItem = "LV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1

            Accessories_BTX.Location = New System.Drawing.Point(326, 532)
            Accessories_BTY.Location = New System.Drawing.Point(416, 536)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf SAM_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1

            Accessories_BTX.Location = New System.Drawing.Point(411, 539)
            Accessories_BTY.Location = New System.Drawing.Point(508, 540)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf SAM_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2

            Accessories_BTX.Location = New System.Drawing.Point(272, 540)
            Accessories_BTY.Location = New System.Drawing.Point(226, 559)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf SAM_VAL_POS.SelectedItem = "OLTC" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(276, 553)
            Accessories_BTY.Location = New System.Drawing.Point(232, 569)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

        Accessories_BTX.Show()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

    End Sub

    Private Sub SV_YAXIS_MouseClick(sender As Object, e As MouseEventArgs)

        If SAM_VAL_POS.SelectedItem = "LV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1

            Accessories_BTX.Location = New System.Drawing.Point(326, 532)
            Accessories_BTY.Location = New System.Drawing.Point(416, 536)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf SAM_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1

            Accessories_BTX.Location = New System.Drawing.Point(411, 539)
            Accessories_BTY.Location = New System.Drawing.Point(508, 540)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf SAM_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2

            Accessories_BTX.Location = New System.Drawing.Point(272, 540)
            Accessories_BTY.Location = New System.Drawing.Point(226, 559)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf SAM_VAL_POS.SelectedItem = "OLTC" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(276, 553)
            Accessories_BTY.Location = New System.Drawing.Point(232, 569)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        End If

        Accessories_BTX.Hide()
        Accessories_BTY.Show()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

    End Sub

    Private Sub FV_XAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles FV_XAXIS.MouseClick

        If FILTER_VAL_POS.SelectedItem = "LV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1

            Accessories_BTX.Location = New System.Drawing.Point(326, 532)
            Accessories_BTY.Location = New System.Drawing.Point(416, 536)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf FILTER_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1

            Accessories_BTX.Location = New System.Drawing.Point(411, 539)
            Accessories_BTY.Location = New System.Drawing.Point(508, 540)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2

            Accessories_BTX.Location = New System.Drawing.Point(272, 540)
            Accessories_BTY.Location = New System.Drawing.Point(226, 559)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(276, 553)
            Accessories_BTY.Location = New System.Drawing.Point(232, 569)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

        Accessories_BTX.Show()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

    End Sub

    Private Sub FV_YAXIS_MouseClick(sender As Object, e As MouseEventArgs)

        If FILTER_VAL_POS.SelectedItem = "LV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_LV1

            Accessories_BTX.Location = New System.Drawing.Point(326, 532)
            Accessories_BTY.Location = New System.Drawing.Point(416, 536)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf FILTER_VAL_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_HV1

            Accessories_BTX.Location = New System.Drawing.Point(411, 539)
            Accessories_BTY.Location = New System.Drawing.Point(508, 540)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_END2

            Accessories_BTX.Location = New System.Drawing.Point(272, 540)
            Accessories_BTY.Location = New System.Drawing.Point(226, 559)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf FILTER_VAL_POS.SelectedItem = "OLTC" Then

            OP1_PIC_BOX.Image = My.Resources.sampling_valve_P725_OLTC_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(276, 553)
            Accessories_BTY.Location = New System.Drawing.Point(232, 569)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        End If

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

        Accessories_BTX.Hide()
        Accessories_BTY.Show()

    End Sub

    Private Sub DV_XAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles DV_XAXIS.MouseClick

        If DRAIN_VALVE_POS.SelectedItem = "LV" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_LV

            Accessories_BTX.Location = New System.Drawing.Point(410, 537)
            Accessories_BTY.Location = New System.Drawing.Point(484, 537)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_HV

            Accessories_BTX.Location = New System.Drawing.Point(413, 554)
            Accessories_BTY.Location = New System.Drawing.Point(546, 556)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_End

            Accessories_BTX.Location = New System.Drawing.Point(262, 564)
            Accessories_BTY.Location = New System.Drawing.Point(218, 564)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(262, 561)
            Accessories_BTY.Location = New System.Drawing.Point(212, 561)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

        Accessories_BTX.Show()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

    End Sub

    Private Sub DV_YAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles DV_YAXIS.MouseClick

        If DRAIN_VALVE_POS.SelectedItem = "LV" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_LV

            Accessories_BTX.Location = New System.Drawing.Point(410, 537)
            Accessories_BTY.Location = New System.Drawing.Point(484, 537)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "HV" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_HV

            Accessories_BTX.Location = New System.Drawing.Point(413, 554)
            Accessories_BTY.Location = New System.Drawing.Point(546, 556)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC OPP" Then

            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_End

            Accessories_BTX.Location = New System.Drawing.Point(262, 564)
            Accessories_BTY.Location = New System.Drawing.Point(218, 564)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        ElseIf DRAIN_VALVE_POS.SelectedItem = "OLTC" Then
            OP1_PIC_BOX.Image = My.Resources.drainvalve_oltc_SIDE

            Accessories_BTX.Location = New System.Drawing.Point(262, 561)
            Accessories_BTY.Location = New System.Drawing.Point(212, 561)

            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

        End If

        Accessories_BTX.Hide()
        Accessories_BTY.Show()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        bt_port1.Hide()
        bt_port2.Hide()
        bt_port3.Hide()

    End Sub


    Private Sub HV_X1_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_X1_AXIS.MouseClick

        If N2_qty_HV.SelectedItem = "1" Then

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_1Q

            Accessories_BTX1.Location = New System.Drawing.Point(284, 507)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_2Q

            Accessories_BTX1.Location = New System.Drawing.Point(158, 462)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_HV.SelectedItem = "3" Then

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_HV

            Accessories_BTX1.Location = New System.Drawing.Point(290, 582)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub HV_X2_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_X2_AXIS.MouseClick

        If N2_qty_HV.SelectedItem = "1" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Show()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_2Q

            Accessories_BTX2.Location = New System.Drawing.Point(379, 462)

            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        ElseIf N2_qty_HV.SelectedItem = "3" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Show()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_HV

            Accessories_BTX2.Location = New System.Drawing.Point(315, 495)

            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub HV_X3_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_X3_AXIS.MouseClick

        If N2_qty_HV.SelectedItem = "1" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

        ElseIf N2_qty_HV.SelectedItem = "3" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Show()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_HV

            Accessories_BTX3.Location = New System.Drawing.Point(241, 496)

            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        End If

    End Sub

    Private Sub HV_Y1_AXIS_MouseClick(sender As Object, e As MouseEventArgs)

        If N2_qty_HV.SelectedItem = "1" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_1Q

            Accessories_BTY1.Location = New System.Drawing.Point(261, 564)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_HV.SelectedItem = "2" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_HV_2Q

            Accessories_BTY1.Location = New System.Drawing.Point(436, 532)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_HV.SelectedItem = "3" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_HV

            Accessories_BTY1.Location = New System.Drawing.Point(397, 530)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_X1_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_X1_AXIS.MouseClick

        If N2_qty_LV.SelectedItem = "1" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_1Q

            Accessories_BTX1.Location = New System.Drawing.Point(355, 509)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_2Q

            Accessories_BTX1.Location = New System.Drawing.Point(217, 474)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Show()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_LV

            Accessories_BTX1.Location = New System.Drawing.Point(353, 580)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_X2_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_X2_AXIS.MouseClick

        If N2_qty_LV.SelectedItem = "1" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Show()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_2Q

            Accessories_BTX2.Location = New System.Drawing.Point(379, 474)

            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Show()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_LV

            Accessories_BTX2.Location = New System.Drawing.Point(328, 494)

            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        End If

    End Sub

    Private Sub LV_X3_AXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_X3_AXIS.MouseClick

        If N2_qty_LV.SelectedItem = "1" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Show()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_LV

            Accessories_BTX3.Location = New System.Drawing.Point(404, 494)

            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_Y1_AXIS_MouseClick(sender As Object, e As MouseEventArgs)

        If N2_qty_LV.SelectedItem = "1" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_1Q

            Accessories_BTY1.Location = New System.Drawing.Point(374, 539)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_LV.SelectedItem = "2" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.NIFPS_LV_2Q

            Accessories_BTY1.Location = New System.Drawing.Point(411, 535)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        ElseIf N2_qty_LV.SelectedItem = "3" Then

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Show()

            Accessories_BTX.Hide()
            Accessories_BTY.Hide()

            OP1_PIC_BOX.Image = My.Resources.N2_PIPE_LV

            Accessories_BTY1.Location = New System.Drawing.Point(247, 528)

            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub HV_XAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_XAXIS.MouseClick

        If EARTHINGPAD_HV.SelectedItem = "Yes" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Show()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.EARTHINGPAD_HV

            Accessories_BTX.Location = New System.Drawing.Point(343, 552)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub HV_YAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles HV_YAXIS.MouseClick

        If EARTHINGPAD_HV.SelectedItem = "Yes" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Show()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.EARTHINGPAD_HV

            Accessories_BTY.Location = New System.Drawing.Point(547, 514)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_XAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_XAXIS.MouseClick

        If EARTHINGPAD_POS.SelectedItem = "Yes" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Show()
            Accessories_BTY.Hide()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.EARTHINGPAD_LV1

            Accessories_BTX.Location = New System.Drawing.Point(528, 554)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub LV_YAXIS_MouseClick(sender As Object, e As MouseEventArgs) Handles LV_YAXIS.MouseClick

        If EARTHINGPAD_POS.SelectedItem = "Yes" Then

            Accessories_BTX1.Hide()
            Accessories_BTX2.Hide()
            Accessories_BTX3.Hide()
            Accessories_BTY1.Hide()

            Accessories_BTX.Hide()
            Accessories_BTY.Show()

            bt_port1.Hide()
            bt_port2.Hide()
            bt_port3.Hide()

            OP1_PIC_BOX.Image = My.Resources.EARTHINGPAD_LV1

            Accessories_BTY.Location = New System.Drawing.Point(634, 503)

            Accessories_BTX1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTX3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            Accessories_BTY.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            Accessories_BTX.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

        End If

    End Sub

    Private Sub NO_OF_PORT_BOT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NO_OF_PORT_BOT.SelectedIndexChanged

        If NO_OF_PORT_BOT.SelectedItem = "2" Then

            NO_OF_PORT_BOT_X1.Enabled = True
            NO_OF_PORT_BOT_X2.Enabled = False
            NO_OF_PORT_BOT_X3.Enabled = True

            NO_OF_PORT_BOT_Y1.Enabled = True

        ElseIf NO_OF_PORT_BOT.SelectedItem = "3" Then

            NO_OF_PORT_BOT_X1.Enabled = True
            NO_OF_PORT_BOT_X2.Enabled = True
            NO_OF_PORT_BOT_X3.Enabled = True

            NO_OF_PORT_BOT_Y1.Enabled = True

        End If
            Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        If NO_OF_PORT_BOT.SelectedItem = "3" Then
            OP1_PIC_BOX.Image = My.Resources.BT_CP_3Port

            bt_port1.Location = New System.Drawing.Point(132, 458)
            bt_port3.Location = New System.Drawing.Point(256, 458)
            bt_port2.Location = New System.Drawing.Point(390, 458)

            bt_port1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            bt_port1.Show()
            bt_port2.Show()
            bt_port3.Show()

        ElseIf NO_OF_PORT_BOT.SelectedItem = "2" Then

            OP1_PIC_BOX.Image = My.Resources.BT_CP_2Port

            bt_port1.Location = New System.Drawing.Point(132, 458)
            bt_port2.Location = New System.Drawing.Point(390, 458)

            bt_port1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            bt_port1.Show()
            bt_port2.Show()
            bt_port3.Hide()

        End If

    End Sub

    Private Sub NO_OF_PORT_BOT_LV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NO_OF_PORT_BOT_LV.SelectedIndexChanged

        If NO_OF_PORT_BOT_LV.SelectedItem = "2" Then

            NO_OF_PORT_BOT_X1_LV.Enabled = True
            NO_OF_PORT_BOT_X2_LV.Enabled = False
            NO_OF_PORT_BOT_X3_LV.Enabled = True

            NO_OF_PORT_BOT_Y1.Enabled = True

        ElseIf NO_OF_PORT_BOT_LV.SelectedItem = "3" Then

            NO_OF_PORT_BOT_X1_LV.Enabled = True
            NO_OF_PORT_BOT_X2_LV.Enabled = True
            NO_OF_PORT_BOT_X3_LV.Enabled = True

            NO_OF_PORT_BOT_Y1.Enabled = True

        End If
        Accessories_BTX.Hide()
        Accessories_BTY.Hide()

        Accessories_BTX1.Hide()
        Accessories_BTX2.Hide()
        Accessories_BTX3.Hide()
        Accessories_BTY1.Hide()

        If NO_OF_PORT_BOT_LV.SelectedItem = "3" Then
            OP1_PIC_BOX.Image = My.Resources.BT_CP_3Port

            bt_port1.Location = New System.Drawing.Point(132, 458)
            bt_port3.Location = New System.Drawing.Point(256, 458)
            bt_port2.Location = New System.Drawing.Point(390, 458)

            bt_port1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            bt_port1.Show()
            bt_port2.Show()
            bt_port3.Show()

        ElseIf NO_OF_PORT_BOT_LV.SelectedItem = "2" Then

            OP1_PIC_BOX.Image = My.Resources.BT_CP_2Port

            bt_port1.Location = New System.Drawing.Point(132, 458)
            bt_port2.Location = New System.Drawing.Point(390, 458)

            bt_port1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            bt_port2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            bt_port1.Show()
            bt_port2.Show()
            bt_port3.Hide()

        End If

    End Sub

    Private Sub T2_MouseClick(sender As Object, e As MouseEventArgs) Handles T2.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

    End Sub

    Private Sub MP_Thk_MouseClick(sender As Object, e As MouseEventArgs) Handles MP_Thk.MouseClick

        UB_L_1.Show()
        UB_L1_1.Show()
        UB_W_1.Show()
        UB_T2_1.Show()

        UB_W_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L1_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_L_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        UB_T2_1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


    End Sub

    Private Sub JP_T_L_LV_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_L_LV.MouseClick

        JP_PIC_BOX1.Image = My.Resources.JP_TOP_STFNR1

        A_JP_LV.Hide()
        B_JP_LV.Hide()
        C_JP_LV.Hide()
        C1_JP_LV.Hide()
        D_JP_LV.Hide()
        E_JP_LV.Hide()
        E1_JP_LV.Hide()
        X_JP_LV.Hide()
        T_JP_LV.Hide()
        JP_LV_500.Hide()
        G_JP_LV.Hide()
        B1_JP_LV.Hide()
        BB1_JP_LV.Hide()
        JP_LV_50.Hide()
        B2_JP_LV.Hide()
        JP_LV_R100.Hide()

        JP_T_L_LV_I.Show()
        JP_T_H_LV_I.Show()
        JP_T_T_LV_I.Show()

        JP_T_L_LV_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_T_H_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_T_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub JP_T_H_LV_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_H_LV.MouseClick

        JP_PIC_BOX1.Image = My.Resources.JP_TOP_STFNR1

        A_JP_LV.Hide()
        B_JP_LV.Hide()
        C_JP_LV.Hide()
        C1_JP_LV.Hide()
        D_JP_LV.Hide()
        E_JP_LV.Hide()
        E1_JP_LV.Hide()
        X_JP_LV.Hide()
        T_JP_LV.Hide()
        JP_LV_500.Hide()
        G_JP_LV.Hide()
        B1_JP_LV.Hide()
        BB1_JP_LV.Hide()
        JP_LV_50.Hide()
        B2_JP_LV.Hide()
        JP_LV_R100.Hide()

        JP_T_L_LV_I.Show()
        JP_T_H_LV_I.Show()
        JP_T_T_LV_I.Show()

        JP_T_L_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_H_LV_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        JP_T_T_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

    End Sub

    Private Sub JP_T_T_LV_MouseClick(sender As Object, e As MouseEventArgs) Handles JP_T_T_LV.MouseClick

        JP_PIC_BOX1.Image = My.Resources.JP_TOP_STFNR1

        A_JP_LV.Hide()
        B_JP_LV.Hide()
        C_JP_LV.Hide()
        C1_JP_LV.Hide()
        D_JP_LV.Hide()
        E_JP_LV.Hide()
        E1_JP_LV.Hide()
        X_JP_LV.Hide()
        T_JP_LV.Hide()
        JP_LV_500.Hide()
        G_JP_LV.Hide()
        B1_JP_LV.Hide()
        BB1_JP_LV.Hide()
        JP_LV_50.Hide()
        B2_JP_LV.Hide()
        JP_LV_R100.Hide()

        JP_T_L_LV_I.Show()
        JP_T_H_LV_I.Show()
        JP_T_T_LV_I.Show()

        JP_T_L_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_H_LV_I.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        JP_T_T_LV_I.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

    End Sub

    Private Sub HVD1_MouseClick(sender As Object, e As MouseEventArgs) Handles HVD1.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_HVD

        OD1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD1.Location = New System.Drawing.Point(304, 445)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(259, 445)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(137, 454)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(402, 455)

    End Sub

    Private Sub HVD2_MouseClick(sender As Object, e As MouseEventArgs) Handles HVD2.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()


        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_HVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(304, 445)

        OD2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD2.Location = New System.Drawing.Point(259, 445)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(137, 454)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(402, 455)

    End Sub

    Private Sub HVD3_MouseClick(sender As Object, e As MouseEventArgs) Handles HVD3.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()


        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_HVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(304, 445)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(259, 445)

        OD3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD3.Location = New System.Drawing.Point(137, 454)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(402, 455)

    End Sub

    Private Sub HVD4_MouseClick(sender As Object, e As MouseEventArgs) Handles HVD4.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_HVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(304, 445)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(259, 445)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(137, 454)

        OD4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD4.Location = New System.Drawing.Point(402, 455)

    End Sub

    Private Sub LVD1_MouseClick(sender As Object, e As MouseEventArgs) Handles LVD1.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_LVD

        OD1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD1.Location = New System.Drawing.Point(371, 439)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(327, 438)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(218, 455)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(463, 455)

    End Sub

    Private Sub LVD2_MouseClick(sender As Object, e As MouseEventArgs) Handles LVD2.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()


        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_LVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(371, 439)

        OD2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD2.Location = New System.Drawing.Point(327, 438)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(218, 455)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(463, 455)

    End Sub

    Private Sub LVD3_MouseClick(sender As Object, e As MouseEventArgs) Handles LVD3.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()


        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_LVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(371, 439)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(327, 438)

        OD3.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD3.Location = New System.Drawing.Point(218, 455)

        OD4.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD4.Location = New System.Drawing.Point(463, 455)

    End Sub

    Private Sub LVD4_MouseClick(sender As Object, e As MouseEventArgs) Handles LVD4.MouseClick

        S1.Hide()
        S2.Hide()
        fs_A.Hide()
        fs_A1.Hide()
        fs_B1.Hide()
        fs_B.Hide()

        fs_oltc_T1.Hide()
        fs_oltc_T2.Hide()
        fs_oltc_D1.Hide()
        fs_oltc_D2.Hide()

        T2PIC.Hide()
        T1PIC.Hide()
        fs_A_2.Hide()
        fs_B_2.Hide()
        S1_PIC.Show()
        S2_PIC.Show()
        OD1.Show()
        OD2.Show()
        OD3.Show()
        OD4.Show()

        STIFNER_PIC_BOX.Image = My.Resources.FLAT_Stiff_LVD

        OD1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD1.Location = New System.Drawing.Point(371, 439)

        OD2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD2.Location = New System.Drawing.Point(327, 438)

        OD3.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        OD3.Location = New System.Drawing.Point(218, 455)

        OD4.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
        OD4.Location = New System.Drawing.Point(463, 455)

    End Sub

    Private Sub FS_HV_REQ_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FS_HV_REQ.SelectedIndexChanged

        If FS_HV_REQ.SelectedItem = "Yes" Then

            HV_STNR_TYPE.Enabled = True
            stiffner_count_HV.Enabled = True
            HV_FS_A.Enabled = True
            HV_FS_A1.Enabled = True
            HV_FS_B.Enabled = True
            HV_FS_B1.Enabled = True
            HVD1.Enabled = True
            HVD2.Enabled = True
            HVD3.Enabled = True
            HVD4.Enabled = True

        ElseIf FS_HV_REQ.SelectedItem = "No" Then

            HV_STNR_TYPE.Enabled = False
            stiffner_count_HV.Enabled = False
            HV_FS_A.Enabled = False
            HV_FS_A1.Enabled = False
            HV_FS_B.Enabled = False
            HV_FS_B1.Enabled = False
            HVD1.Enabled = False
            HVD2.Enabled = False
            HVD3.Enabled = False
            HVD4.Enabled = False

        End If

    End Sub

    Private Sub FS_LV_REQ_SelectedIndexChanged(sender As Object, e As EventArgs) Handles FS_LV_REQ.SelectedIndexChanged

        If FS_LV_REQ.SelectedItem = "Yes" Then

            LV_STNR_TYPE.Enabled = True
            stiffner_count_LV.Enabled = True
            LV_FS_A.Enabled = True
            LV_FS_A1.Enabled = True
            LV_FS_B.Enabled = True
            LV_FS_B1.Enabled = True
            LVD1.Enabled = True
            LVD2.Enabled = True
            LVD3.Enabled = True
            LVD4.Enabled = True

        ElseIf FS_LV_REQ.SelectedItem = "No" Then

            LV_STNR_TYPE.Enabled = False
            stiffner_count_LV.Enabled = False
            LV_FS_A.Enabled = False
            LV_FS_A1.Enabled = False
            LV_FS_B.Enabled = False
            LV_FS_B1.Enabled = False
            LVD1.Enabled = False
            LVD2.Enabled = False
            LVD3.Enabled = False
            LVD4.Enabled = False

        End If

    End Sub
End Class
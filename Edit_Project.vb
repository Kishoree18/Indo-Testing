Imports System.Data.SqlClient
Imports System.Management
Imports System.Runtime.InteropServices

Public Class Edit_Project

    Public Shared wk_odr_No = "0000"

    Private Sub back_btn_Click(sender As Object, e As EventArgs) Handles back_btn.Click
        Main_Panel.Show()
        Me.Close()
    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        user_lbl.Text = String.Empty
        user_lbl.Text = Main_Panel.userTxt.Text
        load_data()

    End Sub

    Private Sub Form5_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Main_Panel.Show()
    End Sub

    Public Sub load_data()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Create_Project"
            Dim cmd As New SqlCommand(query, connection)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            DataGridView1.Rows.Clear()

            Dim sno As Integer = 1

            While reader.Read()
                Dim wk_odr_No As String = reader("wk_odr_No").ToString()
                Dim Type As String = "Bell"
                Dim createdBy As String = reader("Username").ToString()
                Dim dateCreated As String = DateTime.Now.ToString("yyyy-MM-dd")
                Dim customer As String = reader("Customer").ToString()

                DataGridView1.Rows.Add(sno, wk_odr_No, Type, createdBy, dateCreated, customer)
                sno += 1
            End While

        Catch ex As Exception
            MsgBox("err " & ex.ToString())
        End Try

    End Sub



    Public Sub Load_CreateProject()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Create_Project WHERE wk_odr_No = @wk_odr_No"

            Dim command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = command.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Create_Project.dwg_title.Text = reader("dwg_title").ToString()
                Create_Project.wk_odr_No.Text = reader("wk_odr_No").ToString()
                Create_Project.Customer.Text = reader("Customer").ToString()
                Create_Project.Rating.Text = reader("Rating").ToString()
                Create_Project.Voltage.Text = reader("Voltage").ToString()
                Create_Project.Outputlocation.Text = reader("Outputlocation").ToString()
                Create_Project.librarylocation.Text = reader("librarylocation").ToString()

            End While




            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_BasicInputs()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Basic_Input WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Basic_inputs.core_limp.SelectedItem = reader("core_limp").ToString()
                Basic_inputs.core_dia.Text = reader("core_dia").ToString()
                Basic_inputs.stack_thk.Text = reader("stack_thk").ToString()
                Basic_inputs.max_step.Text = reader("max_step").ToString()
                Basic_inputs.leg_centre_1.Text = reader("leg_centre_1").ToString()
                Basic_inputs.leg_centre_2.Text = reader("leg_centre_2").ToString()

                Basic_inputs.Max_winding_od.Text = reader("Max_winding_od").ToString()
                Basic_inputs.window_ht.Text = reader("window_ht").ToString()
                Basic_inputs.end_limb_1.Text = reader("end_limb_1").ToString()
                Basic_inputs.Tp_btm_el_1.Text = reader("Tp_btm_el_1").ToString()
                Basic_inputs.vector_group.Text = reader("vector_group").ToString()
                Basic_inputs.type_of_cooling.SelectedItem = reader("type_of_cooling").ToString()
                Basic_inputs.transform_ft_view.SelectedItem = reader("transform_ft_view").ToString()
                Basic_inputs.OLTC_pkt_position.SelectedItem = reader("OLTC_pkt_position").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_BottomTank_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Bottom_Tank_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.btm_tk_type.SelectedItem = reader("btm_tk_type").ToString()
                Bottom_Tank.btm_tk_curb_pro.SelectedItem = reader("btm_tk_curb_pro").ToString()
                Bottom_Tank.btm_tk_curb_pro.SelectedItem = reader("btm_tk_curb_pro").ToString()
                Bottom_Tank.curb_size.SelectedItem = reader("curb_size").ToString()

                Bottom_Tank.btm_plt_ht.SelectedItem = reader("btm_plt_ht").ToString()
                Bottom_Tank.Cus_ht.Text = reader("Cus_ht").ToString()
                Bottom_Tank.btm_tk_i_len.Text = reader("btm_tk_i_len").ToString()
                Bottom_Tank.btm_plt_len.Text = reader("btm_plt_len").ToString()
                Bottom_Tank.btm_plt_thk.Text = reader("btm_plt_thk").ToString()
                Bottom_Tank.btm_plt_width.Text = reader("btm_plt_width").ToString()
                Bottom_Tank.Side_Plt_thk.Text = reader("Side_Plt_thk").ToString()
                Bottom_Tank.btm_tk_i_wdh.Text = reader("btm_tk_i_wdh").ToString()

                Bottom_Tank.OLTC_pck_len.Text = reader("OLTC_pck_len").ToString()
                Bottom_Tank.OLTC_pck_wdh.Text = reader("OLTC_pck_wdh").ToString()
                Bottom_Tank.tk_curb_wdh.Text = reader("tk_curb_wdh").ToString()
                Bottom_Tank.Curb_wdh_LV.Text = reader("Curb_wdh_LV").ToString()
                Bottom_Tank.Curb_OLTC_End.Text = reader("Curb_OLTC_End").ToString()

                Bottom_Tank.Curb_wdh_D.Text = reader("Curb_wdh_D").ToString()

                Bottom_Tank.GAUDE_PIN_SIZE.SelectedItem = reader("GAUDE_PIN_SIZE").ToString()
                Bottom_Tank.GPX1.Text = reader("GPX1").ToString()
                Bottom_Tank.GPX2.Text = reader("GPX2").ToString()
                Bottom_Tank.GPY1.Text = reader("GPY1").ToString()
                Bottom_Tank.GPY2.Text = reader("GPY2").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_UnderBase_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from UnderBase_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.No_of_Rails.SelectedItem = reader("No_of_Rails").ToString()
                Bottom_Tank.No_of_Rails1.SelectedItem = reader("No_of_Rails1").ToString()
                Bottom_Tank.CC_LENGTH.SelectedItem = reader("CC_LENGTH").ToString()
                Bottom_Tank.Roller_Custom.Text = reader("Roller_Custom").ToString()
                Bottom_Tank.roller_gauge_cmb.SelectedItem = reader("roller_gauge_cmb").ToString()
                Bottom_Tank.roller_orientation.SelectedItem = reader("roller_orientation").ToString()

                Bottom_Tank.BS1_WDH_1.Text = reader("BS1_WDH_1").ToString()
                Bottom_Tank.BS1_HT_WS.Text = reader("BS1_HT_WS").ToString()
                Bottom_Tank.BS1_THK_1.Text = reader("BS1_THK_1").ToString()
                Bottom_Tank.BS2_WDH_1.Text = reader("BS2_WDH_1").ToString()
                Bottom_Tank.BS2_HT_WS.Text = reader("BS2_HT_WS").ToString()
                Bottom_Tank.BS2_THK_1.Text = reader("BS2_THK_1").ToString()
                Bottom_Tank.BS3_WDH_1.Text = reader("BS3_WDH_1").ToString()
                Bottom_Tank.BS3_HT_WS.Text = reader("BS3_HT_WS").ToString()
                Bottom_Tank.BS3_THK_1.Text = reader("BS3_THK_1").ToString()

                Bottom_Tank.HS_CC_len.Text = reader("HS_CC_len").ToString()
                Bottom_Tank.X.Text = reader("X").ToString()
                Bottom_Tank.Distance_A.Text = reader("Distance_A").ToString()
                Bottom_Tank.Distance_B.Text = reader("Distance_B").ToString()

                Bottom_Tank.MP_len.Text = reader("MP_len").ToString()
                Bottom_Tank.MP_Wth.Text = reader("MP_Wth").ToString()
                Bottom_Tank.MP_Thk.Text = reader("MP_Thk").ToString()
                Bottom_Tank.T2.Text = reader("T2").ToString()
                Bottom_Tank.Dim_L1.Text = reader("Dim_L1").ToString()

                Bottom_Tank.prob_support_cmb.SelectedItem = reader("prob_support_cmb").ToString()
                Bottom_Tank.prob_no.SelectedItem = reader("prob_no").ToString()
                Bottom_Tank.PROB_HT.Text = reader("PROB_HT").ToString()
                Bottom_Tank.Prob_X1axis.Text = reader("Prob_X1axis").ToString()
                Bottom_Tank.Prob_X2axis.Text = reader("Prob_X2axis").ToString()
                Bottom_Tank.Prob_Y1axis.Text = reader("Prob_Y1axis").ToString()
                Bottom_Tank.Prob_Y2axis.Text = reader("Prob_Y2axis").ToString()

                Bottom_Tank.BS_L.Text = reader("BS_L").ToString()
                Bottom_Tank.BS_R.Text = reader("BS_R").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub



    Public Sub Load_FlatStiffner_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from FlatStiffner_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.HV_STNR_TYPE.SelectedItem = reader("HV_STNR_TYPE").ToString()
                Bottom_Tank.HV_FS_A.Text = reader("HV_FS_A").ToString()
                Bottom_Tank.HV_FS_A1.Text = reader("HV_FS_A1").ToString()
                Bottom_Tank.HV_FS_B.Text = reader("HV_FS_B").ToString()
                Bottom_Tank.HV_FS_B1.Text = reader("HV_FS_B1").ToString()

                Bottom_Tank.LV_STNR_TYPE.SelectedItem = reader("LV_STNR_TYPE").ToString()
                Bottom_Tank.LV_FS_A.Text = reader("LV_FS_A").ToString()
                Bottom_Tank.LV_FS_A1.Text = reader("LV_FS_A1").ToString()
                Bottom_Tank.LV_FS_B.Text = reader("LV_FS_B").ToString()
                Bottom_Tank.LV_FS_B1.Text = reader("LV_FS_B1").ToString()

                Bottom_Tank.OLTC_STNR_TYPE.SelectedItem = reader("OLTC_STNR_TYPE").ToString()
                Bottom_Tank.OLTC_S1_A.Text = reader("OLTC_S1_A").ToString()
                Bottom_Tank.OLTC_S1_A1.Text = reader("OLTC_S1_A1").ToString()
                Bottom_Tank.OLTC_S1_B.Text = reader("OLTC_S1_B").ToString()
                Bottom_Tank.OLTC_S1_B1.Text = reader("OLTC_S1_B1").ToString()
                Bottom_Tank.OLTC_S1_T.Text = reader("OLTC_S1_T").ToString()
                Bottom_Tank.OS1D1.Text = reader("OS1D1").ToString()
                Bottom_Tank.OS1D2.Text = reader("OS1D2").ToString()

                Bottom_Tank.OLTC_STNR2_TYPE.SelectedItem = reader("OLTC_STNR2_TYPE").ToString()
                Bottom_Tank.OLTC_S2_A.Text = reader("OLTC_S2_A").ToString()
                Bottom_Tank.OLTC_S2_B.Text = reader("OLTC_S2_B").ToString()
                Bottom_Tank.OLTC_S2_A1.Text = reader("OLTC_S2_A1").ToString()
                Bottom_Tank.OS2D1.Text = reader("OS2D1").ToString()
                Bottom_Tank.OS2D2.Text = reader("OS2D2").ToString()

                Bottom_Tank.OLTC_END_STNR_TYP.SelectedItem = reader("OLTC_END_STNR_TYP").ToString()
                Bottom_Tank.Stiffner_count_BT.SelectedItem = reader("Stiffner_count_BT").ToString()
                Bottom_Tank.oltc_stiffner1_count.SelectedItem = reader("oltc_stiffner1_count").ToString()
                Bottom_Tank.oltc_stiffner2_count.SelectedItem = reader("oltc_stiffner2_count").ToString()
                Bottom_Tank.OLTC_END_A.Text = reader("OLTC_END_A").ToString()
                Bottom_Tank.OLTC_END_B.Text = reader("OLTC_END_B").ToString()
                Bottom_Tank.OLTC_END_A1.Text = reader("OLTC_END_A1").ToString()
                Bottom_Tank.OED1.Text = reader("OED1").ToString()
                Bottom_Tank.OED2.Text = reader("OED2").ToString()
                Bottom_Tank.OED3.Text = reader("OED3").ToString()
                Bottom_Tank.OED4.Text = reader("OED4").ToString()

                Bottom_Tank.stiffner_count_HV.SelectedItem = reader("stiffner_count_HV").ToString()
                Bottom_Tank.HVD1.Text = reader("HVD1").ToString()
                Bottom_Tank.HVD2.Text = reader("HVD2").ToString()
                Bottom_Tank.HVD3.Text = reader("HVD3").ToString()
                Bottom_Tank.HVD4.Text = reader("HVD4").ToString()

                Bottom_Tank.stiffner_count_LV.SelectedItem = reader("stiffner_count_LV").ToString()
                Bottom_Tank.LVD1.Text = reader("LVD1").ToString()
                Bottom_Tank.LVD2.Text = reader("LVD2").ToString()
                Bottom_Tank.LVD3.Text = reader("LVD3").ToString()
                Bottom_Tank.LVD4.Text = reader("LVD4").ToString()

                Bottom_Tank.FS_HV_REQ.SelectedItem = reader("FS_HV_REQ").ToString()
                Bottom_Tank.FS_LV_REQ.SelectedItem = reader("FS_LV_REQ").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_JackingPad_HV_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from JackingPad_HV_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.jack_pad_type.SelectedItem = reader("jack_pad_type").ToString()
                Bottom_Tank.JP_A.Text = reader("JP_A").ToString()
                Bottom_Tank.JP_B.Text = reader("JP_B").ToString()
                Bottom_Tank.JP_T.Text = reader("JP_T").ToString()
                Bottom_Tank.JP_G.Text = reader("JP_G").ToString()
                Bottom_Tank.JP_E.Text = reader("JP_E").ToString()
                Bottom_Tank.JP_D.Text = reader("JP_D").ToString()
                Bottom_Tank.JP_C.Text = reader("JP_C").ToString()
                Bottom_Tank.JP_B1.Text = reader("JP_B1").ToString()
                Bottom_Tank.JP_BB1.Text = reader("JP_BB1").ToString()
                Bottom_Tank.JP_B2.Text = reader("JP_B2").ToString()

                Bottom_Tank.NO_TOP_JP_HV.SelectedItem = reader("NO_TOP_JP_HV").ToString()
                Bottom_Tank.JP_T_L.Text = reader("JP_T_L").ToString()
                Bottom_Tank.JP_T_H.Text = reader("JP_T_H").ToString()
                Bottom_Tank.JP_T_T.Text = reader("JP_T_T").ToString()

                Bottom_Tank.NO_OF_JP_HV.SelectedItem = reader("NO_OF_JP_HV").ToString()
                Bottom_Tank.HV_JPD1.Text = reader("HV_JPD1").ToString()
                Bottom_Tank.HV_JPD2.Text = reader("HV_JPD2").ToString()
                Bottom_Tank.HV_JPD3.Text = reader("HV_JPD3").ToString()
                Bottom_Tank.HV_JPD4.Text = reader("HV_JPD4").ToString()
                Bottom_Tank.JP_Y_DIST.Text = reader("JP_Y_DIST").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_JakingPad_LV_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from JakingPad_LV_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.jack_pad_type1.SelectedItem = reader("jack_pad_type1").ToString()
                Bottom_Tank.JP_LV_A.Text = reader("JP_LV_A").ToString()
                Bottom_Tank.JP_LV_B.Text = reader("JP_LV_B").ToString()
                Bottom_Tank.JP_LV_C.Text = reader("JP_LV_C").ToString()
                Bottom_Tank.JP_LV_D.Text = reader("JP_LV_D").ToString()
                Bottom_Tank.JP_LV_E.Text = reader("JP_LV_E").ToString()
                Bottom_Tank.JP_LV_G.Text = reader("JP_LV_G").ToString()
                Bottom_Tank.JP_LV_T.Text = reader("JP_LV_T").ToString()
                Bottom_Tank.JP_LV_B1.Text = reader("JP_LV_B1").ToString()
                Bottom_Tank.JP_LV_BB1.Text = reader("JP_LV_BB1").ToString()
                Bottom_Tank.JP_LV_B2.Text = reader("JP_LV_B2").ToString()

                Bottom_Tank.NO_TOP_JP_LV.SelectedItem = reader("NO_TOP_JP_LV").ToString()
                Bottom_Tank.JP_T_L_LV.Text = reader("JP_T_L_LV").ToString()
                Bottom_Tank.JP_T_T_LV.Text = reader("JP_T_T_LV").ToString()
                Bottom_Tank.JP_T_H_LV.Text = reader("JP_T_H_LV").ToString()

                Bottom_Tank.NO_JP_LV.SelectedItem = reader("NO_JP_LV").ToString()
                Bottom_Tank.JP_Y_DIST_LV.Text = reader("JP_Y_DIST_LV").ToString()
                Bottom_Tank.LV_JP1.Text = reader("LV_JP1").ToString()
                Bottom_Tank.LV_JP2.Text = reader("LV_JP2").ToString()
                Bottom_Tank.LV_JP3.Text = reader("LV_JP3").ToString()
                Bottom_Tank.LV_JP4.Text = reader("LV_JP4").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Accessories_BT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories_BT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Bottom_Tank.SAMP_PIPE_SIZE.SelectedItem = reader("SAMP_PIPE_SIZE").ToString()
                Bottom_Tank.SAM_VAL_POS.SelectedItem = reader("SAM_VAL_POS").ToString()
                Bottom_Tank.SV_XAXIS.Text = reader("SV_XAXIS").ToString()

                Bottom_Tank.FILTER_VALUVE_SIZE.SelectedItem = reader("FILTER_VALUVE_SIZE").ToString()
                Bottom_Tank.FILTER_VAL_POS.SelectedItem = reader("FILTER_VAL_POS").ToString()
                Bottom_Tank.FV_XAXIS.Text = reader("FV_XAXIS").ToString()

                Bottom_Tank.DRAIN_VALVE_SIZE.SelectedItem = reader("DRAIN_VALVE_SIZE").ToString()
                Bottom_Tank.DRAIN_VALVE_POS.SelectedItem = reader("DRAIN_VALVE_POS").ToString()
                Bottom_Tank.DV_XAXIS.Text = reader("DV_XAXIS").ToString()
                Bottom_Tank.DV_YAXIS.Text = reader("DV_YAXIS").ToString()

                Bottom_Tank.NIFPS_REQ.SelectedItem = reader("NIFPS_REQ").ToString()
                Bottom_Tank.NIFPS_VALSIZE.SelectedItem = reader("NIFPS_VALSIZE").ToString()

                Bottom_Tank.N2_filling_req_HV.SelectedItem = reader("N2_filling_req_HV").ToString()
                Bottom_Tank.N2_qty_HV.SelectedItem = reader("N2_qty_HV").ToString()
                Bottom_Tank.HV_X1_AXIS.Text = reader("HV_X1_AXIS").ToString()
                Bottom_Tank.HV_X2_AXIS.Text = reader("HV_X2_AXIS").ToString()
                Bottom_Tank.HV_X3_AXIS.Text = reader("HV_X3_AXIS").ToString()

                Bottom_Tank.N2_filling_req_LV.SelectedItem = reader("N2_filling_req_LV").ToString()
                Bottom_Tank.N2_qty_LV.SelectedItem = reader("N2_qty_LV").ToString()
                Bottom_Tank.LV_X1_AXIS.Text = reader("LV_X1_AXIS").ToString()
                Bottom_Tank.LV_X2_AXIS.Text = reader("LV_X2_AXIS").ToString()
                Bottom_Tank.LV_X3_AXIS.Text = reader("LV_X3_AXIS").ToString()


                Bottom_Tank.N2_filling_req_OLTC_opp.SelectedItem = reader("N2_filling_req_OLTC_opp").ToString()
                Bottom_Tank.N2_qty_OLTC_opp.SelectedItem = reader("N2_qty_OLTC_opp").ToString()
                Bottom_Tank.OLTC_opp_X1_AXIS.Text = reader("OLTC_opp_X1_AXIS").ToString()
                Bottom_Tank.OLTC_opp_X2_AXIS.Text = reader("OLTC_opp_X2_AXIS").ToString()
                Bottom_Tank.OLTC_opp_X3_AXIS.Text = reader("OLTC_opp_X3_AXIS").ToString()

                Bottom_Tank.N2_filling_req_OLTC.SelectedItem = reader("N2_filling_req_OLTC").ToString()
                Bottom_Tank.N2_qty_OLTC.SelectedItem = reader("N2_qty_OLTC").ToString()
                Bottom_Tank.OLTC_X1_AXIS.Text = reader("OLTC_X1_AXIS").ToString()
                Bottom_Tank.OLTC_X2_AXIS.Text = reader("OLTC_X2_AXIS").ToString()
                Bottom_Tank.OLTC_X3_AXIS.Text = reader("OLTC_X3_AXIS").ToString()

                Bottom_Tank.EARTHINGPAD_HV.SelectedItem = reader("EARTHINGPAD_HV").ToString()
                Bottom_Tank.HV_XAXIS.Text = reader("HV_XAXIS").ToString()
                Bottom_Tank.HV_YAXIS.Text = reader("HV_YAXIS").ToString()

                Bottom_Tank.EARTHINGPAD_POS.SelectedItem = reader("EARTHINGPAD_POS").ToString()
                Bottom_Tank.LV_XAXIS.Text = reader("LV_XAXIS").ToString()
                Bottom_Tank.LV_YAXIS.Text = reader("LV_YAXIS").ToString()

                Bottom_Tank.earthing_pad_HV_type.SelectedItem = reader("earthing_pad_HV_type").ToString()
                Bottom_Tank.earthing_pad_LV_type.SelectedItem = reader("earthing_pad_LV_type").ToString()

                Bottom_Tank.NO_OF_PORT_BOT.SelectedItem = reader("NO_OF_PORT_BOT").ToString()

                'ONAN
                Cooler_Bank.ONAN_no_port.SelectedItem = reader("NO_OF_PORT_BOT").ToString()
                'ONAF
                Cooler_Bank.NO_OF_PORT_ONAF.SelectedItem = reader("NO_OF_PORT_BOT").ToString()

                Bottom_Tank.NO_OF_PORT_BOT_X1.Text = reader("NO_OF_PORT_BOT_X1").ToString()
                Bottom_Tank.NO_OF_PORT_BOT_X2.Text = reader("NO_OF_PORT_BOT_X2").ToString()
                Bottom_Tank.NO_OF_PORT_BOT_X3.Text = reader("NO_OF_PORT_BOT_X3").ToString()

                Bottom_Tank.NO_OF_PORT_BOT_Y1.Text = reader("NO_OF_PORT_BOT_Y1").ToString()

                Bottom_Tank.NO_OF_PORT_BOT_LV.SelectedItem = reader("NO_OF_PORT_BOT_LV").ToString()

                Bottom_Tank.NO_OF_PORT_BOT_X1_LV.Text = reader("NO_OF_PORT_BOT_X1_LV").ToString()
                Bottom_Tank.NO_OF_PORT_BOT_X2_LV.Text = reader("NO_OF_PORT_BOT_X2_LV").ToString()
                Bottom_Tank.NO_OF_PORT_BOT_X3_LV.Text = reader("NO_OF_PORT_BOT_X3_LV").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Tank_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Tank_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.bt_tk_wdh_2.Text = reader("bt_tk_wdh_2").ToString()
                Top_Tank.curb_wid.Text = reader("curb_wid").ToString()
                Top_Tank.curb_len.Text = reader("curb_len").ToString()

                Top_Tank.T_TANK_WTH.Text = reader("T_TANK_WTH").ToString()
                Top_Tank.T_TANK_WTH2.Text = reader("T_TANK_WTH2").ToString()
                Top_Tank.T_IN_LEN.Text = reader("T_IN_LEN").ToString()
                Top_Tank.T_T_PLT_THK.Text = reader("T_T_PLT_THK").ToString()
                Top_Tank.TNK_HT.Text = reader("TNK_HT").ToString()
                Top_Tank.T_TOP_CVR_THK.Text = reader("T_TOP_CVR_THK").ToString()

                Top_Tank.T_SLOPE_SIDE.SelectedItem = reader("T_SLOPE_SIDE").ToString()
                Top_Tank.tank_slope_val.SelectedItem = reader("tank_slope_val").ToString()

                Top_Tank.SLOPE_CUS.Text = reader("SLOPE_CUS").ToString()
                Top_Tank.Bushing_poc_A.Text = reader("Bushing_poc_A").ToString()
                Top_Tank.Bushing_poc_B.Text = reader("Bushing_poc_B").ToString()
                Top_Tank.Bushing_poc_C.Text = reader("Bushing_poc_C").ToString()
                Top_Tank.NT_X.Text = reader("NT_X").ToString()
                Top_Tank.NT_H.Text = reader("NT_H").ToString()

                Top_Tank.DM_BOX_ORIN.SelectedItem = reader("DM_BOX_ORIN").ToString()
                Top_Tank.TAP_CHANGER_TYPE.SelectedItem = reader("TAP_CHANGER_TYPE").ToString()
                Top_Tank.TAP_make.SelectedItem = reader("TAP_make").ToString()
                Top_Tank.TAP_model_no.SelectedItem = reader("TAP_model_no").ToString()
                Top_Tank.TIE_IN_REQ.SelectedItem = reader("TIE_IN_REQ").ToString()
                Top_Tank.Tie_in_resist.SelectedItem = reader("Tie_in_resist").ToString()
                Top_Tank.MOUNTING.SelectedItem = reader("MOUNTING").ToString()

                Top_Tank.OLTC_POC_LEN.Text = reader("OLTC_POC_LEN").ToString()
                Top_Tank.OLTC_POC_WDH.Text = reader("OLTC_POC_WDH").ToString()

                Top_Tank.OLTC_HEAD_VANT.SelectedItem = reader("OLTC_HEAD_VANT").ToString()


                Top_Tank.OLTC_POC_HT.Text = reader("OLTC_POC_HT").ToString()
                Top_Tank.UP_TO_CL.Text = reader("UP_TO_CL").ToString()
                Top_Tank.VP_TO_CL.Text = reader("VP_TO_CL").ToString()
                Top_Tank.WP_TO_CL.Text = reader("WP_TO_CL").ToString()

                Top_Tank.LV_TERM.SelectedItem = reader("LV_TERM").ToString()
                Top_Tank.LV_TERM_OPN.SelectedItem = reader("LV_TERM_OPN").ToString()
                Top_Tank.LV_T1.SelectedItem = reader("LV_T1").ToString()
                Top_Tank.LV_T2.SelectedItem = reader("LV_T2").ToString()
                Top_Tank.OLTC_conservator_type.SelectedItem = reader("OLTC_conservator_type").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_HVTermination_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from HVTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.HV_Termiantion.SelectedItem = reader("HV_Termiantion").ToString()
                Top_Tank.HV_Termiantion_Mounting.SelectedItem = reader("HV_Termiantion_Mounting").ToString()

                'U
                Top_Tank.HV_bush_rating_U.SelectedItem = reader("HV_bush_rating_U").ToString()
                Top_Tank.HV_bush_BCT_U.SelectedItem = reader("HV_bush_BCT_U").ToString()
                Top_Tank.Bush_Mounting_U.SelectedItem = reader("Bush_Mounting_U").ToString()
                Top_Tank.HV_turret_dia_U.SelectedItem = reader("HV_turret_dia_U").ToString()
                Top_Tank.HV_turret_typ_U.SelectedItem = reader("HV_turret_typ_U").ToString()
                Top_Tank.HV_turret_ht_U.SelectedItem = reader("HV_turret_ht_U").ToString()

                Top_Tank.HV_ht_sml_turret_body_U.Text = reader("HV_ht_sml_turret_body_U").ToString()
                Top_Tank.HV_welded_trunk_ht_U.Text = reader("HV_welded_trunk_ht_U").ToString()
                Top_Tank.HV_turret_ang_U.Text = reader("HV_turret_ang_U").ToString()
                Top_Tank.HV_UH.Text = reader("HV_UH").ToString()
                Top_Tank.HV_ht_CT_MP_U.Text = reader("HV_ht_CT_MP_U").ToString()
                Top_Tank.HV_ht_CT_box_U.Text = reader("HV_ht_CT_box_U").ToString()
                Top_Tank.HV_ID_CT_MP_U.Text = reader("HV_ID_CT_MP_U").ToString()

                Top_Tank.HV_bush_MTG_U.SelectedItem = reader("HV_bush_MTG_U").ToString()
                Top_Tank.HV_U_CT_ter_box.SelectedItem = reader("HV_U_CT_ter_box").ToString()
                Top_Tank.CT_BOX_U.SelectedItem = reader("CT_BOX_U").ToString()

                Top_Tank.HV_U_CT_ter_box_angle.Text = reader("HV_U_CT_ter_box_angle").ToString()

                Top_Tank.HV_U_inspection_cover_cmb.SelectedItem = reader("HV_U_inspection_cover_cmb").ToString()
                Top_Tank.IC_partno_U.SelectedItem = reader("IC_partno_U").ToString()

                Top_Tank.HV_U_inspection_cover_angle.Text = reader("HV_U_inspection_cover_angle").ToString()
                Top_Tank.HV_U_equalizing_cover_angle.Text = reader("HV_U_equalizing_cover_angle").ToString()

                Top_Tank.HV_ht_turret_body_plt_U.SelectedItem = reader("HV_ht_turret_body_plt_U").ToString()


                'V
                Top_Tank.HV_bush_rating_V.SelectedItem = reader("HV_bush_rating_V").ToString()
                Top_Tank.HV_bush_BCT_V.SelectedItem = reader("HV_bush_BCT_V").ToString()
                Top_Tank.Bush_Mounting_V.SelectedItem = reader("Bush_Mounting_V").ToString()
                Top_Tank.HV_turret_dia_V.SelectedItem = reader("HV_turret_dia_V").ToString()
                Top_Tank.HV_turret_typ_V.SelectedItem = reader("HV_turret_typ_V").ToString()
                Top_Tank.HV_turret_ht_V.SelectedItem = reader("HV_turret_ht_V").ToString()

                Top_Tank.HV_ht_sml_turret_body_V.Text = reader("HV_ht_sml_turret_body_V").ToString()
                Top_Tank.HV_welded_trunk_ht_V.Text = reader("HV_welded_trunk_ht_V").ToString()
                Top_Tank.HV_turret_ang_V.Text = reader("HV_turret_ang_V").ToString()
                Top_Tank.HV_VH.Text = reader("HV_VH").ToString()
                Top_Tank.HV_ht_CT_MP_V.Text = reader("HV_ht_CT_MP_V").ToString()
                Top_Tank.HV_ht_CT_box_V.Text = reader("HV_ht_CT_box_V").ToString()
                Top_Tank.HV_ID_CT_MP_V.Text = reader("HV_ID_CT_MP_V").ToString()

                Top_Tank.HV_bush_MTG_V.SelectedItem = reader("HV_bush_MTG_V").ToString()
                Top_Tank.HV_V_CT_ter_box.SelectedItem = reader("HV_V_CT_ter_box").ToString()
                Top_Tank.CT_BOX_V.SelectedItem = reader("CT_BOX_V").ToString()

                Top_Tank.HV_V_CT_ter_box_angle.Text = reader("HV_V_CT_ter_box_angle").ToString()

                Top_Tank.HV_V_inspection_cover_cmb.SelectedItem = reader("HV_V_inspection_cover_cmb").ToString()
                Top_Tank.IC_partno_V.SelectedItem = reader("IC_partno_V").ToString()

                Top_Tank.HV_V_inspection_cover_angle.Text = reader("HV_V_inspection_cover_angle").ToString()
                Top_Tank.HV_V_equalizing_cover_angle.Text = reader("HV_V_equalizing_cover_angle").ToString()

                Top_Tank.HV_ht_turret_body_plt_V.SelectedItem = reader("HV_ht_turret_body_plt_V").ToString()

                'W



                Top_Tank.HV_bush_rating_W.SelectedItem = reader("HV_bush_rating_W").ToString()
                Top_Tank.HV_bush_BCT_W.SelectedItem = reader("HV_bush_BCT_W").ToString()
                Top_Tank.Bush_Mounting_W.SelectedItem = reader("Bush_Mounting_W").ToString()
                Top_Tank.HV_turret_dia_W.SelectedItem = reader("HV_turret_dia_W").ToString()
                Top_Tank.HV_turret_typ_W.SelectedItem = reader("HV_turret_typ_W").ToString()
                Top_Tank.HV_turret_ht_W.SelectedItem = reader("HV_turret_ht_W").ToString()

                Top_Tank.HV_ht_sml_turret_body_W.Text = reader("HV_ht_sml_turret_body_W").ToString()
                Top_Tank.HV_welded_trunk_ht_W.Text = reader("HV_welded_trunk_ht_W").ToString()
                Top_Tank.HV_turret_ang_W.Text = reader("HV_turret_ang_W").ToString()
                Top_Tank.HV_WH.Text = reader("HV_WH").ToString()
                Top_Tank.HV_ht_CT_MP_W.Text = reader("HV_ht_CT_MP_W").ToString()
                Top_Tank.HV_ht_CT_box_W.Text = reader("HV_ht_CT_box_W").ToString()
                Top_Tank.HV_ID_CT_MP_W.Text = reader("HV_ID_CT_MP_W").ToString()

                Top_Tank.HV_bush_MTG_W.SelectedItem = reader("HV_bush_MTG_W").ToString()
                Top_Tank.HV_W_CT_ter_box.SelectedItem = reader("HV_W_CT_ter_box").ToString()
                Top_Tank.CT_BOX_W.SelectedItem = reader("CT_BOX_W").ToString()

                Top_Tank.HV_W_CT_ter_box_angle.Text = reader("HV_W_CT_ter_box_angle").ToString()

                Top_Tank.HV_W_inspection_cover_cmb.SelectedItem = reader("HV_W_inspection_cover_cmb").ToString()
                Top_Tank.IC_partno_W.SelectedItem = reader("IC_partno_W").ToString()

                Top_Tank.HV_W_inspection_cover_angle.Text = reader("HV_W_inspection_cover_angle").ToString()
                Top_Tank.HV_W_equalizing_cover_angle.Text = reader("HV_W_equalizing_cover_angle").ToString()

                Top_Tank.HV_ht_turret_body_plt_W.SelectedItem = reader("HV_ht_turret_body_plt_W").ToString()

                Top_Tank.Bush_Mounting_UL.SelectedItem = reader("Bush_Mounting_UL").ToString()
                Top_Tank.Bush_Mounting_VL.SelectedItem = reader("Bush_Mounting_VL").ToString()
                Top_Tank.Bush_Mounting_WL.SelectedItem = reader("Bush_Mounting_WL").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_HVNTermination_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from HVNTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.HVN_Termination_Mounting.SelectedItem = reader("HVN_Termination_Mounting").ToString()

                'U
                Top_Tank.HVN_bush_rating_U.SelectedItem = reader("HVN_bush_rating_U").ToString()
                Top_Tank.MOUNTING_HVN.SelectedItem = reader("MOUNTING_HVN").ToString()
                Top_Tank.HVN_bush_BCT_U.SelectedItem = reader("HVN_bush_BCT_U").ToString()
                Top_Tank.HVN_turret_dia_U.SelectedItem = reader("HVN_turret_dia_U").ToString()
                Top_Tank.HVN_turret_ty_U.SelectedItem = reader("HVN_turret_ty_U").ToString()
                Top_Tank.HVN_turret_ht_U.SelectedItem = reader("HVN_turret_ht_U").ToString()

                Top_Tank.HVN_welded_trunk_ht_U.Text = reader("HVN_welded_trunk_ht_U").ToString()
                Top_Tank.HVN_ht_small_turret_U.Text = reader("HVN_ht_small_turret_U").ToString()
                Top_Tank.HVN_turret_ang_U.Text = reader("HVN_turret_ang_U").ToString()
                Top_Tank.HVN_H.Text = reader("HVN_H").ToString()
                Top_Tank.HVN_ht_CT_MP_U.Text = reader("HVN_ht_CT_MP_U").ToString()
                Top_Tank.HVN_ht_CT_box_U.Text = reader("HVN_ht_CT_box_U").ToString()
                Top_Tank.HVN_ID_CT_MP_U.Text = reader("HVN_ID_CT_MP_U").ToString()

                Top_Tank.HVN_bush_MTG_pad_U.SelectedItem = reader("HVN_bush_MTG_pad_U").ToString()
                Top_Tank.HVN_CT_terminal_U.SelectedItem = reader("HVN_CT_terminal_U").ToString()
                Top_Tank.HVN_CTT_TYP.SelectedItem = reader("HVN_CTT_TYP").ToString()

                Top_Tank.HV_CT_terminal_ang_U.Text = reader("HV_CT_terminal_ang_U").ToString()

                Top_Tank.HV_inspection_cover_U.SelectedItem = reader("HV_inspection_cover_U").ToString()
                'Top_Tank.HVN_IC_TYP.SelectedItem = reader("HVN_IC_TYP").ToString()

                Top_Tank.HV_inspection_cover_ang_U.Text = reader("HV_inspection_cover_ang_U").ToString()
                Top_Tank.HVN_equaliznig_ang_U.Text = reader("HVN_equaliznig_ang_U").ToString()

                Top_Tank.turret_incl_ang.Text = reader("turret_incl_ang").ToString()



                Top_Tank.HVN_ht_turret_bd_pl_U.SelectedItem = reader("HVN_ht_turret_bd_pl_U").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_LVTermination_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from LVTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.LV_Termination.SelectedItem = reader("LV_Termination").ToString()
                Top_Tank.LV_Termination_Mounting.SelectedItem = reader("LV_Termination_Mounting").ToString()
                Top_Tank.LV1_Termination.SelectedItem = reader("LV1_Termination").ToString()
                Top_Tank.LV2_Termination.SelectedItem = reader("LV2_Termination").ToString()

                'U
                Top_Tank.LV_bush_rating_U.SelectedItem = reader("LV_bush_rating_U").ToString()
                Top_Tank.MOUNT_LVU.SelectedItem = reader("MOUNT_LVU").ToString()
                Top_Tank.LV_bush_BCT_U.SelectedItem = reader("LV_bush_BCT_U").ToString()
                Top_Tank.LV_turret_dia_U.SelectedItem = reader("LV_turret_dia_U").ToString()
                Top_Tank.LV_turret_type_U.SelectedItem = reader("LV_turret_type_U").ToString()
                Top_Tank.LV_turret_ht_U.SelectedItem = reader("LV_turret_ht_U").ToString()

                Top_Tank.LV_welded_trunk_ht_U.Text = reader("LV_welded_trunk_ht_U").ToString()
                Top_Tank.LV_ht_small_turret_U.Text = reader("LV_ht_small_turret_U").ToString()
                Top_Tank.LV_turret_angle_U.Text = reader("LV_turret_angle_U").ToString()
                Top_Tank.LV_UH.Text = reader("LV_UH").ToString()
                Top_Tank.LV_ht_CT_MP_U.Text = reader("LV_ht_CT_MP_U").ToString()
                Top_Tank.LV_ht_CT_box_U.Text = reader("LV_ht_CT_box_U").ToString()
                Top_Tank.LV_ID_CT_MP_U.Text = reader("LV_ID_CT_MP_U").ToString()

                Top_Tank.LV_bush_MTG_pad_U.SelectedItem = reader("LV_bush_MTG_pad_U").ToString()
                Top_Tank.LV_CT_terminal_box_U.SelectedItem = reader("LV_CT_terminal_box_U").ToString()
                Top_Tank.LV_U1_CT_TYP.SelectedItem = reader("LV_U1_CT_TYP").ToString()

                Top_Tank.LV_CT_terminal_box_ang_U.Text = reader("LV_CT_terminal_box_ang_U").ToString()

                Top_Tank.LV_inspection_cover_U.SelectedItem = reader("LV_inspection_cover_U").ToString()
                Top_Tank.LV_U1_IC_TYP.SelectedItem = reader("LV_U1_IC_TYP").ToString()

                Top_Tank.LV_inspection_cover_ang_U.Text = reader("LV_inspection_cover_ang_U").ToString()
                Top_Tank.LV_equalizing_cover_U.Text = reader("LV_equalizing_cover_U").ToString()

                Top_Tank.LV_ht_turret_body_U.SelectedItem = reader("LV_ht_turret_body_U").ToString()

                'V
                Top_Tank.LV_bush_rating_V.SelectedItem = reader("LV_bush_rating_V").ToString()
                Top_Tank.MOUNT_LVV.SelectedItem = reader("MOUNT_LVV").ToString()
                Top_Tank.LV_bush_BCT_V.SelectedItem = reader("LV_bush_BCT_V").ToString()
                Top_Tank.LV_turret_dia_V.SelectedItem = reader("LV_turret_dia_V").ToString()
                Top_Tank.LV_turret_type_V.SelectedItem = reader("LV_turret_type_V").ToString()
                Top_Tank.LV_turret_ht_V.SelectedItem = reader("LV_turret_ht_V").ToString()

                Top_Tank.LV_welded_trunk_ht_V.Text = reader("LV_welded_trunk_ht_V").ToString()
                Top_Tank.LV_ht_small_turret_V.Text = reader("LV_ht_small_turret_V").ToString()
                Top_Tank.LV_turret_angle_V.Text = reader("LV_turret_angle_V").ToString()
                Top_Tank.LV_VH.Text = reader("LV_VH").ToString()
                Top_Tank.LV_ht_CT_MP_V.Text = reader("LV_ht_CT_MP_V").ToString()
                Top_Tank.LV_ht_CT_box_V.Text = reader("LV_ht_CT_box_V").ToString()
                Top_Tank.LV_ID_CT_MP_V.Text = reader("LV_ID_CT_MP_V").ToString()

                Top_Tank.LV_bush_MTG_pad_V.SelectedItem = reader("LV_bush_MTG_pad_V").ToString()
                Top_Tank.LV_CT_terminal_box_V.SelectedItem = reader("LV_CT_terminal_box_V").ToString()
                Top_Tank.LV_V1_CT_TYP.SelectedItem = reader("LV_V1_CT_TYP").ToString()

                Top_Tank.LV_CT_terminal_box_ang_V.Text = reader("LV_CT_terminal_box_ang_V").ToString()

                Top_Tank.LV_inspection_cover_V.SelectedItem = reader("LV_inspection_cover_V").ToString()
                Top_Tank.LV_V1_IC_TYP.SelectedItem = reader("LV_V1_IC_TYP").ToString()

                Top_Tank.LV_inspection_cover_ang_V.Text = reader("LV_inspection_cover_ang_V").ToString()
                Top_Tank.LV_equalizing_cover_V.Text = reader("LV_equalizing_cover_V").ToString()

                Top_Tank.LV_ht_turret_body_V.SelectedItem = reader("LV_ht_turret_body_V").ToString()

                'U
                Top_Tank.LV_bush_rating_W.SelectedItem = reader("LV_bush_rating_W").ToString()
                Top_Tank.MOUNT_LVW.SelectedItem = reader("MOUNT_LVW").ToString()
                Top_Tank.LV_bush_BCT_W.SelectedItem = reader("LV_bush_BCT_W").ToString()
                Top_Tank.LV_turret_dia_W.SelectedItem = reader("LV_turret_dia_W").ToString()
                Top_Tank.LV_turret_type_W.SelectedItem = reader("LV_turret_type_W").ToString()
                Top_Tank.LV_turret_ht_W.SelectedItem = reader("LV_turret_ht_W").ToString()

                Top_Tank.LV_welded_trunk_ht_W.Text = reader("LV_welded_trunk_ht_W").ToString()
                Top_Tank.LV_ht_small_turret_W.Text = reader("LV_ht_small_turret_W").ToString()
                Top_Tank.LV_turret_angle_W.Text = reader("LV_turret_angle_W").ToString()
                Top_Tank.LV_WH.Text = reader("LV_WH").ToString()
                Top_Tank.LV_ht_CT_MP_W.Text = reader("LV_ht_CT_MP_W").ToString()
                Top_Tank.LV_ht_CT_box_W.Text = reader("LV_ht_CT_box_W").ToString()
                Top_Tank.LV_ID_CT_MP_W.Text = reader("LV_ID_CT_MP_W").ToString()

                Top_Tank.LV_bush_MTG_pad_W.SelectedItem = reader("LV_bush_MTG_pad_W").ToString()
                Top_Tank.LV_CT_terminal_box_W.SelectedItem = reader("LV_CT_terminal_box_W").ToString()
                Top_Tank.LV_W1_CT_TYP.SelectedItem = reader("LV_W1_CT_TYP").ToString()

                Top_Tank.LV_CT_terminal_box_ang_W.Text = reader("LV_CT_terminal_box_ang_W").ToString()

                Top_Tank.LV_inspection_cover_W.SelectedItem = reader("LV_inspection_cover_W").ToString()
                Top_Tank.LV_W1_IC_TYP.SelectedItem = reader("LV_W1_IC_TYP").ToString()

                Top_Tank.LV_inspection_cover_ang_W.Text = reader("LV_inspection_cover_ang_W").ToString()
                Top_Tank.LV_equalizing_cover_W.Text = reader("LV_equalizing_cover_W").ToString()

                Top_Tank.LV_ht_turret_body_W.SelectedItem = reader("LV_ht_turret_body_W").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_BareBushing_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from BareBushing_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Top_Tank.LV1_W.Text = reader("LV1_W").ToString()
                Top_Tank.LV2_W.Text = reader("LV2_W").ToString()
                Top_Tank.LV1_V.Text = reader("LV1_V").ToString()
                Top_Tank.LV2_V.Text = reader("LV2_V").ToString()
                Top_Tank.LV1_U.Text = reader("LV1_U").ToString()
                Top_Tank.LV2_U.Text = reader("LV2_U").ToString()
                Top_Tank.LVN1_C.Text = reader("LVN1_C").ToString()
                Top_Tank.LVN2.Text = reader("LVN2").ToString()
                Top_Tank.LVN_C.Text = reader("LVN_C").ToString()
                Top_Tank.LVBC.Text = reader("LVBC").ToString()
                Top_Tank.LVNBC.Text = reader("LVNBC").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_CableBox_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from CableBox_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Top_Tank.LV1_CB.SelectedItem = reader("LV1_CB").ToString()
                Top_Tank.LV1_CB_Mounting.SelectedItem = reader("LV1_CB_Mounting").ToString()
                Top_Tank.LV1_CB_Position.SelectedItem = reader("LV1_CB_Position").ToString()
                Top_Tank.LVN_Termination_Mounting.SelectedItem = reader("LVN_Termination_Mounting").ToString()
                Top_Tank.LVN_SIDE.SelectedItem = reader("LVN_SIDE").ToString()

                Top_Tank.LVN_CB_Z.Text = reader("LVN_CB_Z").ToString()
                Top_Tank.LVN_CB_TP.Text = reader("LVN_CB_TP").ToString()
                Top_Tank.LVN_CB_Z.Text = reader("LVN_CB_Z").ToString()
                Top_Tank.LVN_CB_X.Text = reader("LVN_CB_X").ToString()
                Top_Tank.LV1_trunk_offset.Text = reader("LV1_trunk_offset").ToString()
                Top_Tank.LV1_trunk_height.Text = reader("LV1_trunk_height").ToString()

                Top_Tank.LV1_CB_Clearance.SelectedItem = reader("LV1_CB_Clearance").ToString()
                Top_Tank.LV1_CB_PH.SelectedItem = reader("LV1_CB_PH").ToString()
                Top_Tank.LV1_CB_ER.SelectedItem = reader("LV1_CB_ER").ToString()

                Top_Tank.LV1_bush_rating_U.SelectedItem = reader("LV1_bush_rating_U").ToString()
                Top_Tank.LV1_bush_mountain_pad_U.SelectedItem = reader("LV1_bush_mountain_pad_U").ToString()

                Top_Tank.LV_bushing_U.Text = reader("LV_bushing_U").ToString()
                Top_Tank.LV1_welded_trunk_ht_U.Text = reader("LV1_welded_trunk_ht_U").ToString()

                Top_Tank.LV1_CT_terminal_box_U.SelectedItem = reader("LV1_CT_terminal_box_U").ToString()
                Top_Tank.LV1_inspec_cover1_U.Text = reader("LV1_inspec_cover1_U").ToString()
                Top_Tank.LV1_inspec_cover2_U.Text = reader("LV1_inspec_cover2_U").ToString()
                Top_Tank.LV1_top_gusset_U.Text = reader("LV1_top_gusset_U").ToString()
                Top_Tank.LV1_bot_gusset_U.Text = reader("LV1_bot_gusset_U").ToString()
                Top_Tank.LV1_side1_gusset_U.Text = reader("LV1_side1_gusset_U").ToString()
                Top_Tank.LV1_side2_gusset_U.Text = reader("LV1_side2_gusset_U").ToString()

                Top_Tank.CB_Width.Text = reader("CB_Width").ToString()
                Top_Tank.CB_Height.Text = reader("CB_Height").ToString()
                Top_Tank.CB_TopLength.Text = reader("CB_TopLength").ToString()
                Top_Tank.CB_Length1.Text = reader("CB_Length1").ToString()
                Top_Tank.CB_Length2.Text = reader("CB_Length2").ToString()
                Top_Tank.DC_Length.Text = reader("DC_Length").ToString()
                Top_Tank.Trunk_Length.Text = reader("Trunk_Length").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_LV2CableBox_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from LV2CableBox_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Top_Tank.LV2_CB.SelectedItem = reader("LV2_CB").ToString()
                Top_Tank.LV2_CB_Mounting.SelectedItem = reader("LV2_CB_Mounting").ToString()
                Top_Tank.LV2_CB_Position.SelectedItem = reader("LV2_CB_Position").ToString()

                Top_Tank.CB1_X.Text = reader("CB1_X").ToString()
                Top_Tank.CB2_X.Text = reader("CB2_X").ToString()
                Top_Tank.LVN_CB_TP2.Text = reader("LVN_CB_TP2").ToString()
                Top_Tank.LVN_CB_X2.Text = reader("LVN_CB_X2").ToString()
                Top_Tank.LVN2_CB_Z.Text = reader("LVN2_CB_Z").ToString()
                Top_Tank.LVN2_CB_Z2.Text = reader("LVN2_CB_Z2").ToString()
                Top_Tank.CB2_THT.Text = reader("CB2_THT").ToString()
                Top_Tank.CB2_BHT.Text = reader("CB2_BHT").ToString()

                Top_Tank.LV2_CB_Clearance.SelectedItem = reader("LV2_CB_Clearance").ToString()
                Top_Tank.LV2_CB_PH.SelectedItem = reader("LV2_CB_PH").ToString()
                Top_Tank.LV2_CB_ER.SelectedItem = reader("LV2_CB_ER").ToString()

                Top_Tank.LV2_bush_rating.SelectedItem = reader("LV2_bush_rating").ToString()
                Top_Tank.LV2_bush_mounting_pad.SelectedItem = reader("LV2_bush_mounting_pad").ToString()

                Top_Tank.LV2_bushing.Text = reader("LV2_bushing").ToString()
                Top_Tank.LV2_welded_trunk_ht.Text = reader("LV2_welded_trunk_ht").ToString()

                Top_Tank.LV2_CT_terminal_box.SelectedItem = reader("LV2_CT_terminal_box").ToString()
                Top_Tank.LV2_inspec_cover1.Text = reader("LV2_inspec_cover1").ToString()
                Top_Tank.LV2_inspec_cover2.Text = reader("LV2_inspec_cover2").ToString()
                Top_Tank.LV2_top_gusset.Text = reader("LV2_top_gusset").ToString()
                Top_Tank.LV2_bot_gusset.Text = reader("LV2_bot_gusset").ToString()
                Top_Tank.LV2_side1_gusset.Text = reader("LV2_side1_gusset").ToString()
                Top_Tank.LV2_side2_gusset.Text = reader("LV2_side2_gusset").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_CB_Term_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String =
        DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()

            Dim query As String =
        "SELECT * FROM CB_Term_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)

            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)

            Dim reader As SqlDataReader = cmd.ExecuteReader()


            While reader.Read()

                '====================================================
                ' TOP DETAILS
                '====================================================

                'CB Type
                Top_Tank.CB_Type.SelectedItem =
            reader("CB_Type").ToString()

                Top_Tank.CB_X.Text =
            reader("CB_X").ToString()

                Top_Tank.CB_Y.Text =
            reader("CB_Y").ToString()

                'CB Mounting Position
                Top_Tank.CB_Mounting_POS.SelectedItem =
            reader("CB_Mounting_POS").ToString()


                'CB Plate Thickness
                Top_Tank.CB_BTM_Length.Text =
            reader("CB_BTM_Length").ToString()

                '====================================================
                ' CABLE BOX DETAILS
                '====================================================

                'Cable Box Width
                Top_Tank.CB_WTH.Text =
            reader("CB_WTH").ToString()


                'Cable Box Length
                Top_Tank.CB_Length.Text =
            reader("CB_Length").ToString()


                'Bushing Center to Gland Height
                Top_Tank.Bushing_C_Gland_HT.Text =
            reader("Bushing_C_Gland_HT").ToString()



                '====================================================
                ' SPACE HEATER / BREATHER
                '====================================================

                'Space Heater
                Top_Tank.CB_Space_Heater.SelectedItem =
            reader("CB_Space_Heater").ToString()


                'Breather Pipe
                Top_Tank.CB_BreatherPipe.SelectedItem =
            reader("CB_BreatherPipe").ToString()


                'Breather Pipe Position
                Top_Tank.CB_BreatherPipe_POS.SelectedItem =
            reader("CB_BreatherPipe_POS").ToString()



                '====================================================
                ' FRONT INSPECTION COVER
                '====================================================

                Top_Tank.CB_Front_Inspec_Cover.SelectedItem =
            reader("CB_Front_Inspec_Cover").ToString()



                '====================================================
                ' BUSHING TO TOP PLATE HEIGHT
                '====================================================

                Top_Tank.Bushing_Top_Plt_HT.Text =
            reader("Bushing_Top_Plt_HT").ToString()



                '====================================================
                ' GLAND PLATE
                '====================================================

                'Gland Plate Material
                Top_Tank.CB_Gland_Plt_Mat.SelectedItem =
            reader("CB_Gland_Plt_Mat").ToString()


                'Gland Position
                Top_Tank.CB_Gland_POS.SelectedItem =
            reader("CB_Gland_POS").ToString()



                '====================================================
                ' INSULATION SUPPORT
                '====================================================

                Top_Tank.CB_Insulation_Distance1.Text =
            reader("CB_Insulation_Distance1").ToString()

                Top_Tank.CB_Insulation_Distance2.Text =
            reader("CB_Insulation_Distance2").ToString()

                Top_Tank.CB_Insulation_Distance3.Text =
            reader("CB_Insulation_Distance3").ToString()



                '====================================================
                ' DESICCATORS
                '====================================================

                Top_Tank.CB_Desiccators.SelectedItem =
            reader("CB_Desiccators").ToString()


            End While


            reader.Close()
            connection.Close()


        Catch ex As Exception

            MsgBox("Error" & ex.ToString)

        End Try

    End Sub

    Public Sub Load_Trunk_DC_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If


            ' Get selected Work Order No
            Dim wk_odr_No As String =
            DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()


            Dim query As String =
            "SELECT * FROM Trunk_DC_TT WHERE wk_odr_No = @wk_odr_No"


            Dim cmd As New SqlCommand(query, connection)

            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)


            Dim reader As SqlDataReader = cmd.ExecuteReader()


            While reader.Read()


                '====================================================
                ' PLATE / TRUNK DETAILS
                '====================================================

                'CB Plate Thickness
                Top_Tank.CB_Plt_Thk.Text =
                reader("CB_Plt_Thk").ToString()


                'CB Welded Trunk Height
                Top_Tank.CB_Welded_Trunk_HT.Text =
                reader("CB_Welded_Trunk_HT").ToString()


                'CB Trunk Height
                Top_Tank.CB_Trunk_Ht.Text =
                reader("CB_Trunk_Ht").ToString()


                'CB Trunk Width
                Top_Tank.CB_Trunk_Wth.Text =
                reader("CB_Trunk_Wth").ToString()


                'CB Trunk Length
                Top_Tank.CB_Trunk_Lth.Text =
                reader("CB_Trunk_Lth").ToString()



                '====================================================
                ' CT BOX DETAILS
                '====================================================

                'CT Box Reference
                Top_Tank.CTBox_Ref.SelectedItem =
                reader("CTBox_Ref").ToString()


                'CT Box Quantity
                Top_Tank.CTBox_Qty.SelectedItem =
                reader("CTBox_Qty").ToString()


                'CB CT Quantity
                Top_Tank.CB_CT_Qty.Text =
                reader("CB_CT_Qty").ToString()


                'CT Box ID
                Top_Tank.CTBox_ID.Text =
                reader("CTBox_ID").ToString()


                'CT Box OD
                Top_Tank.CTBox_OD.Text =
                reader("CTBox_OD").ToString()


                'CT Box Height
                Top_Tank.CTBox_HT.Text =
                reader("CTBox_HT").ToString()



                '====================================================
                ' INSPECTION COVER / MOUNTING PAD
                '====================================================

                'CB Inspection Cover Type
                Top_Tank.CB_Inspec_Cover_Type.SelectedItem =
                reader("CB_Inspec_Cover_Type").ToString()


                'CB Trunk Mounting Pad Selection
                Top_Tank.CB_MP_Sel.SelectedItem =
                reader("CB_MP_Sel").ToString()


                'CB Trunk Mounting Pad
                Top_Tank.CB_Trunk_MP.SelectedItem =
                reader("CB_Trunk_MP").ToString()


                'CB Mounting Pad Center to Center
                Top_Tank.CB_MP_CC.Text =
                reader("CB_MP_CC").ToString()



                '====================================================
                ' DC DETAILS
                '====================================================

                Top_Tank.DC_Side_HT.Text =
                reader("DC_Side_HT").ToString()

                Top_Tank.DC_Side_WTH.Text =
                reader("DC_Side_WTH").ToString()

                Top_Tank.DC_Side_LG.Text =
                reader("DC_Side_LG").ToString()

                'DC Inside Length
                Top_Tank.DC_Inside_LG.Text =
                reader("DC_Inside_LG").ToString()


                'DC Depth
                Top_Tank.DC_Depth.Text =
                reader("DC_Depth").ToString()


                'DC Depth Offside
                Top_Tank.DC_Depth_Offside.Text =
                reader("DC_Depth_Offside").ToString()


                'DC Inside Height
                Top_Tank.DC_Inside_HT.Text =
                reader("DC_Inside_HT").ToString()


                'DC Flange Height
                Top_Tank.DC_Flange_HT.Text =
                reader("DC_Flange_HT").ToString()

                Top_Tank.DC_Top_Radius.Text =
                reader("DC_Top_Radius").ToString()


                'DC Inspection Cover
                Top_Tank.DC_Inspec_Cover.SelectedItem =
                reader("DC_Inspec_Cover").ToString()



                '====================================================
                ' FIRE / FILE SENSOR MOUNTING
                '====================================================

                'File Sensor Mounting
                Top_Tank.File_Sensor_Mounting.SelectedItem =
                reader("File_Sensor_Mounting").ToString()


                'File Sensor Mounting Make
                Top_Tank.File_Sensor_Mounting_Make.SelectedItem =
                reader("File_Sensor_Mounting_Make").ToString()


            End While


            reader.Close()
            connection.Close()


        Catch ex As Exception

            MsgBox("Error" & ex.ToString)

        End Try

    End Sub

    Public Sub Load_LVNTermination_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from LVNTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.LVN_Termination.SelectedItem = reader("LVN_Termination").ToString()
                Top_Tank.LVN1_Termination.SelectedItem = reader("LVN1_Termination").ToString()

                Top_Tank.LVN_Mounting.SelectedItem = reader("LVN_Mounting").ToString()
                Top_Tank.LVN_POS.SelectedItem = reader("LVN_POS").ToString()
                Top_Tank.LVN_X.Text = reader("LVN_X").ToString()
                Top_Tank.LVN_Y.Text = reader("LVN_Y").ToString()

                'U
                Top_Tank.LVN1_bush_rating.SelectedItem = reader("LVN1_bush_rating").ToString()
                Top_Tank.LVN1_Mounting.SelectedItem = reader("LVN1_Mounting").ToString()
                Top_Tank.LVN1_bush_BCT.SelectedItem = reader("LVN1_bush_BCT").ToString()
                Top_Tank.LVN1_turret_type.SelectedItem = reader("LVN1_turret_type").ToString()
                Top_Tank.LVN1_turret_ht.SelectedItem = reader("LVN1_turret_ht").ToString()
                Top_Tank.LVN1_turret_dia.SelectedItem = reader("LVN1_turret_dia").ToString()

                Top_Tank.LVN1_welded_trunk_ht.Text = reader("LVN1_welded_trunk_ht").ToString()
                Top_Tank.LVN1_ht_small_turret_body.Text = reader("LVN1_ht_small_turret_body").ToString()
                Top_Tank.LVN1_turret_angle.Text = reader("LVN1_turret_angle").ToString()
                Top_Tank.LVN1_H.Text = reader("LVN1_H").ToString()
                Top_Tank.LVN1_ht_CT_MP.Text = reader("LVN1_ht_CT_MP").ToString()
                Top_Tank.LVN1_ht_CT_box.Text = reader("LVN1_ht_CT_box").ToString()
                Top_Tank.LVN1_ID_CT_MP.Text = reader("LVN1_ID_CT_MP").ToString()

                Top_Tank.LVN1_bush_MTG_pad.SelectedItem = reader("LVN1_bush_MTG_pad").ToString()
                Top_Tank.LVN1_CT_terminal_box.SelectedItem = reader("LVN1_CT_terminal_box").ToString()
                Top_Tank.LVN1_CT_TYP.SelectedItem = reader("LVN1_CT_TYP").ToString()

                Top_Tank.LVN1_CT_terminal_box_ang.Text = reader("LVN1_CT_terminal_box_ang").ToString()

                Top_Tank.LVN1_inspection_cover.SelectedItem = reader("LVN1_inspection_cover").ToString()
                Top_Tank.LVN1_IC_TYP.SelectedItem = reader("LVN1_IC_TYP").ToString()

                Top_Tank.LVN1_inspection_cover_ang.Text = reader("LVN1_inspection_cover_ang").ToString()
                Top_Tank.LV1_equalizing_cover.Text = reader("LV1_equalizing_cover").ToString()

                Top_Tank.LVN1_ht_turret_body.SelectedItem = reader("LVN1_ht_turret_body").ToString()

                'V
                Top_Tank.LVN2_bush_rating.SelectedItem = reader("LVN2_bush_rating").ToString()
                Top_Tank.LVN2_Mounting.SelectedItem = reader("LVN2_Mounting").ToString()
                Top_Tank.LVN2_bush_BCT.SelectedItem = reader("LVN2_bush_BCT").ToString()
                Top_Tank.LVN2_turret_type.SelectedItem = reader("LVN2_turret_type").ToString()
                Top_Tank.LVN2_turret_ht.SelectedItem = reader("LVN2_turret_ht").ToString()
                Top_Tank.LVN2_turret_dia.SelectedItem = reader("LVN2_turret_dia").ToString()

                Top_Tank.LVN2_welded_trunk_ht.Text = reader("LVN2_welded_trunk_ht").ToString()
                Top_Tank.LVN2_ht_small_turret_body.Text = reader("LVN2_ht_small_turret_body").ToString()
                Top_Tank.LVN2_turret_angle.Text = reader("LVN2_turret_angle").ToString()
                Top_Tank.LVN2_H.Text = reader("LVN2_H").ToString()
                Top_Tank.LVN2_ht_CT_MP.Text = reader("LVN2_ht_CT_MP").ToString()
                Top_Tank.LVN2_ht_CT_box.Text = reader("LVN2_ht_CT_box").ToString()
                Top_Tank.LVN2_ID_CT_MP.Text = reader("LVN2_ID_CT_MP").ToString()

                Top_Tank.LVN2_bush_MTG_pad.SelectedItem = reader("LVN2_bush_MTG_pad").ToString()
                Top_Tank.LVN2_CT_terminal_box.SelectedItem = reader("LVN2_CT_terminal_box").ToString()
                Top_Tank.LVN2_CT_TYP.SelectedItem = reader("LVN2_CT_TYP").ToString()

                Top_Tank.LVN2_CT_terminal_box_ang.Text = reader("LVN2_CT_terminal_box_ang").ToString()

                Top_Tank.LVN2_inspection_cover.SelectedItem = reader("LVN2_inspection_cover").ToString()
                Top_Tank.LVN2_IC_TYP.SelectedItem = reader("LVN2_IC_TYP").ToString()

                Top_Tank.LVN2_inspection_cover_ang.Text = reader("LVN2_inspection_cover_ang").ToString()
                Top_Tank.LV2_equalizing_cover.Text = reader("LV2_equalizing_cover").ToString()

                Top_Tank.LVN2_ht_turret_body.SelectedItem = reader("LVN2_ht_turret_body").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Stiffner1_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Stiffner1_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.BS1_A.Text = reader("BS1_A").ToString()
                Top_Tank.BS1_A1.Text = reader("BS1_A1").ToString()
                Top_Tank.BS1_B.Text = reader("BS1_B").ToString()
                Top_Tank.BS1_C.Text = reader("BS1_C").ToString()
                Top_Tank.BS1_D.Text = reader("BS1_D").ToString()
                Top_Tank.BS1_I.Text = reader("BS1_I").ToString()
                Top_Tank.BS1_E.Text = reader("BS1_E").ToString()
                Top_Tank.BS1_F.Text = reader("BS1_F").ToString()
                Top_Tank.BS1_G.Text = reader("BS1_G").ToString()
                Top_Tank.BS1_H.Text = reader("BS1_H").ToString()
                Top_Tank.BS1_J.Text = reader("BS1_J").ToString()
                Top_Tank.BS1_K.Text = reader("BS1_K").ToString()

                Top_Tank.HV_DIM1.Text = reader("HV_DIM1").ToString()
                Top_Tank.HV_DIM2.Text = reader("HV_DIM2").ToString()
                Top_Tank.HV_DIM3.Text = reader("HV_DIM3").ToString()
                Top_Tank.HV_DIM4.Text = reader("HV_DIM4").ToString()

                Top_Tank.LV_S1A.Text = reader("LV_S1A").ToString()
                Top_Tank.LV_S1B.Text = reader("LV_S1B").ToString()
                Top_Tank.LV_S1C.Text = reader("LV_S1C").ToString()
                Top_Tank.LV_S1D.Text = reader("LV_S1D").ToString()
                Top_Tank.LV2_S1A.Text = reader("LV2_S1A").ToString()
                Top_Tank.LV2_S1B.Text = reader("LV2_S1B").ToString()
                Top_Tank.LV2_S1C.Text = reader("LV2_S1C").ToString()
                Top_Tank.LV2_S1D.Text = reader("LV2_S1D").ToString()
                Top_Tank.LH1_S1A.Text = reader("LH1_S1A").ToString()
                Top_Tank.LH1_S1B.Text = reader("LH1_S1B").ToString()
                Top_Tank.LH1_S1C.Text = reader("LH1_S1C").ToString()
                Top_Tank.LH1_S1D.Text = reader("LH1_S1D").ToString()
                Top_Tank.LV3_S1A.Text = reader("LV3_S1A").ToString()
                Top_Tank.LV3_S1B.Text = reader("LV3_S1B").ToString()
                Top_Tank.LV3_S1C.Text = reader("LV3_S1C").ToString()
                Top_Tank.LV3_S1D.Text = reader("LV3_S1D").ToString()
                Top_Tank.LH2_S1A.Text = reader("LH2_S1A").ToString()
                Top_Tank.LH2_S1B.Text = reader("LH2_S1B").ToString()
                Top_Tank.LH2_S1C.Text = reader("LH2_S1C").ToString()
                Top_Tank.LH2_S1D.Text = reader("LH2_S1D").ToString()

                Top_Tank.LV_D1.Text = reader("LV_D1").ToString()
                Top_Tank.LV_D2.Text = reader("LV_D2").ToString()
                Top_Tank.LV_D3.Text = reader("LV_D3").ToString()
                Top_Tank.LV_D4.Text = reader("LV_D4").ToString()
                Top_Tank.LV_D5.Text = reader("LV_D5").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_OLTC_Stiffner1_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from OLTC_Stiffner1_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.OLTC_FS1_WDTH.Text = reader("OLTC_FS1_WDTH").ToString()
                Top_Tank.OLTC_FS1_THK.Text = reader("OLTC_FS1_THK").ToString()
                Top_Tank.OLTC_FS1_OTY.Text = reader("OLTC_FS1_OTY").ToString()
                Top_Tank.OLTC_FS2_WDTH.Text = reader("OLTC_FS2_WDTH").ToString()
                Top_Tank.OLTC_FS2_THK.Text = reader("OLTC_FS2_THK").ToString()
                Top_Tank.OLTC_FS2_OTY.Text = reader("OLTC_FS2_OTY").ToString()
                Top_Tank.OLTC_FS3_WDTH.Text = reader("OLTC_FS3_WDTH").ToString()
                Top_Tank.OLTC_FS3_THK.Text = reader("OLTC_FS3_THK").ToString()
                Top_Tank.OLTC_FS3_OTY.Text = reader("OLTC_FS3_OTY").ToString()
                Top_Tank.OLTC_FS4_WDTH.Text = reader("OLTC_FS4_WDTH").ToString()
                Top_Tank.OLTC_FS4_THK.Text = reader("OLTC_FS4_THK").ToString()
                Top_Tank.OLTC_FS4_OTY.Text = reader("OLTC_FS4_OTY").ToString()

                Top_Tank.OLTC_DIM1.Text = reader("OLTC_DIM1").ToString()
                Top_Tank.OLTC_DIM2.Text = reader("OLTC_DIM2").ToString()
                Top_Tank.OLTC_DIM3.Text = reader("OLTC_DIM3").ToString()
                Top_Tank.OLTC_DIM4.Text = reader("OLTC_DIM4").ToString()

                Top_Tank.OLTC_BS1_QTY.Text = reader("OLTC_BS1_QTY").ToString()
                Top_Tank.OLTC_BS1_WDTH.Text = reader("OLTC_BS1_WDTH").ToString()
                Top_Tank.OLTC_BS1_HT.Text = reader("OLTC_BS1_HT").ToString()
                Top_Tank.OLTC_BS1_THK.Text = reader("OLTC_BS1_THK").ToString()
                Top_Tank.OLTC_BS2_QTY.Text = reader("OLTC_BS2_QTY").ToString()
                Top_Tank.OLTC_BS2_WDTH.Text = reader("OLTC_BS2_WDTH").ToString()
                Top_Tank.OLTC_BS2_HT.Text = reader("OLTC_BS2_HT").ToString()
                Top_Tank.OLTC_BS2_THK.Text = reader("OLTC_BS2_THK").ToString()
                Top_Tank.OLTC_BS3_QTY.Text = reader("OLTC_BS3_QTY").ToString()
                Top_Tank.OLTC_BS3_WDTH.Text = reader("OLTC_BS3_WDTH").ToString()
                Top_Tank.OLTC_BS3_HT.Text = reader("OLTC_BS3_HT").ToString()
                Top_Tank.OLTC_BS3_THK.Text = reader("OLTC_BS3_THK").ToString()

                Top_Tank.OLTC_D1.Text = reader("OLTC_D1").ToString()
                Top_Tank.OLTC_D2.Text = reader("OLTC_D2").ToString()
                Top_Tank.OLTC_D3.Text = reader("OLTC_D3").ToString()
                Top_Tank.OLTC_D4.Text = reader("OLTC_D4").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Lashing_leg_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Lashing_Leg_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.HV_LL_cmb.SelectedItem = reader("HV_LL_cmb").ToString()
                Top_Tank.LV_LL_cmb.SelectedItem = reader("LV_LL_cmb").ToString()
                Top_Tank.OLTC_LL_cmb.SelectedItem = reader("OLTC_LL_cmb").ToString()
                Top_Tank.OLTC_OPP_LL_cmb.SelectedItem = reader("OLTC_OPP_LL_cmb").ToString()

                Top_Tank.HV_LL_H1.Text = reader("HV_LL_H1").ToString()

                Top_Tank.LV_LL_H1.Text = reader("LV_LL_H1").ToString()
                Top_Tank.LV_LL_H2.Text = reader("LV_LL_H2").ToString()
                Top_Tank.LV_LL_H3.Text = reader("LV_LL_H3").ToString()
                Top_Tank.LV_LL_H4.Text = reader("LV_LL_H4").ToString()
                Top_Tank.LV_LL_H5.Text = reader("LV_LL_H5").ToString()

                'Top_Tank.top_tank_ladder_cmb.SelectedItem = reader("top_tank_ladder_cmb").ToString()
                'Top_Tank.A_frame_ladder_cmb.SelectedItem = reader("A_frame_ladder_cmb").ToString()
                'Top_Tank.top_tank_ladder_ht.Text = reader("top_tank_ladder_ht").ToString()
                Top_Tank.top_tank_ladder_dis.Text = reader("top_tank_ladder_dis").ToString()
                Top_Tank.top_tank_ladder_dis_LV.Text = reader("top_tank_ladder_dis_LV").ToString()
                Top_Tank.top_tank_ladder_dis_OLTC.Text = reader("top_tank_ladder_dis_OLTC").ToString()
                Top_Tank.top_tank_ladder_dis_OLTC_OPP.Text = reader("top_tank_ladder_dis_OLTC_OPP").ToString()

                Top_Tank.rating_plate_HV_X.Text = reader("rating_plate_HV_X").ToString()
                Top_Tank.rating_plate_LV_X.Text = reader("rating_plate_LV_X").ToString()
                Top_Tank.rating_plate_OLTC_X.Text = reader("rating_plate_OLTC_X").ToString()
                Top_Tank.rating_plate_OLTC_OPP_X.Text = reader("rating_plate_OLTC_OPP_X").ToString()

                Top_Tank.rating_plate_HV_Y.Text = reader("rating_plate_HV_Y").ToString()
                Top_Tank.rating_plate_LV_Y.Text = reader("rating_plate_LV_Y").ToString()
                Top_Tank.rating_plate_OLTC_Y.Text = reader("rating_plate_OLTC_Y").ToString()
                Top_Tank.rating_plate_OLTC_OPP_Y.Text = reader("rating_plate_OLTC_OPP_Y").ToString()

                Top_Tank.rating_plate_HV_X2.Text = reader("rating_plate_HV_X2").ToString()
                Top_Tank.rating_plate_LV_X2.Text = reader("rating_plate_LV_X2").ToString()
                Top_Tank.rating_plate_OLTC_X2.Text = reader("rating_plate_OLTC_X2").ToString()
                Top_Tank.rating_plate_OLTC_OPP_X2.Text = reader("rating_plate_OLTC_OPP_X2").ToString()

                Top_Tank.rating_plate_HV_Y2.Text = reader("rating_plate_HV_Y2").ToString()
                Top_Tank.rating_plate_LV_Y2.Text = reader("rating_plate_LV_Y2").ToString()
                Top_Tank.rating_plate_OLTC_Y2.Text = reader("rating_plate_OLTC_Y2").ToString()
                Top_Tank.rating_plate_OLTC_OPP_Y2.Text = reader("rating_plate_OLTC_OPP_Y2").ToString()

                Top_Tank.rating_plate_HV_X3.Text = reader("rating_plate_HV_X3").ToString()
                Top_Tank.rating_plate_LV_X3.Text = reader("rating_plate_LV_X3").ToString()
                Top_Tank.rating_plate_OLTC_X3.Text = reader("rating_plate_OLTC_X3").ToString()
                Top_Tank.rating_plate_OLTC_OPP_X3.Text = reader("rating_plate_OLTC_OPP_X3").ToString()

                Top_Tank.rating_plate_HV_Y3.Text = reader("rating_plate_HV_Y3").ToString()
                Top_Tank.rating_plate_LV_Y3.Text = reader("rating_plate_LV_Y3").ToString()
                Top_Tank.rating_plate_OLTC_Y3.Text = reader("rating_plate_OLTC_Y3").ToString()
                Top_Tank.rating_plate_OLTC_OPP_Y3.Text = reader("rating_plate_OLTC_OPP_Y3").ToString()

                Top_Tank.HV_LL_H2.Text = reader("HV_LL_H2").ToString()
                Top_Tank.HV_LL_H3.Text = reader("HV_LL_H3").ToString()
                Top_Tank.HV_LL_H4.Text = reader("HV_LL_H4").ToString()
                Top_Tank.HV_LL_H5.Text = reader("HV_LL_H5").ToString()

                Top_Tank.OLTC_LL_H1.Text = reader("OLTC_LL_H1").ToString()
                Top_Tank.OLTC_LL_H2.Text = reader("OLTC_LL_H2").ToString()

                Top_Tank.OLTC_OPP_LL_H1.Text = reader("OLTC_OPP_LL_H1").ToString()
                Top_Tank.OLTC_OPP_LL_H2.Text = reader("OLTC_OPP_LL_H2").ToString()

                Top_Tank.LADDER_MOUNTING.SelectedItem = reader("LADDER_MOUNTING").ToString()

                Top_Tank.rating_plate_HV_X4.Text = reader("rating_plate_HV_X4").ToString()
                Top_Tank.rating_plate_LV_X4.Text = reader("rating_plate_LV_X4").ToString()
                Top_Tank.rating_plate_OLTC_X4.Text = reader("rating_plate_OLTC_X4").ToString()
                Top_Tank.rating_plate_OLTC_OPP_X4.Text = reader("rating_plate_OLTC_OPP_X4").ToString()

                Top_Tank.rating_plate_HV_Y4.Text = reader("rating_plate_HV_Y4").ToString()
                Top_Tank.rating_plate_LV_Y4.Text = reader("rating_plate_LV_Y4").ToString()
                Top_Tank.rating_plate_OLTC_Y4.Text = reader("rating_plate_OLTC_Y4").ToString()
                Top_Tank.rating_plate_OLTC_OPP_Y4.Text = reader("rating_plate_OLTC_OPP_Y4").ToString()

                Top_Tank.LADDER_ANGLE.Text = reader("LADDER_ANGLE").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_GCD_PARTS_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from GCD_PARTS WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.IC_HV_REQ.SelectedItem = reader("IC_HV_REQ").ToString()
                Top_Tank.IC_HV_COUNT.SelectedItem = reader("IC_HV_COUNT").ToString()
                Top_Tank.IC_HV_X1.Text = reader("IC_HV_X1").ToString()
                Top_Tank.IC_HV_X2.Text = reader("IC_HV_X2").ToString()
                Top_Tank.IC_HV_X3.Text = reader("IC_HV_X3").ToString()
                Top_Tank.IC_HV_Y1.Text = reader("IC_HV_Y1").ToString()
                Top_Tank.IC_HV_Y2.Text = reader("IC_HV_Y2").ToString()
                Top_Tank.IC_HV_Y3.Text = reader("IC_HV_Y3").ToString()

                Top_Tank.IC_HVOP_REQ.SelectedItem = reader("IC_HVOP_REQ").ToString()
                Top_Tank.IC_HVOP_COUNT.SelectedItem = reader("IC_HVOP_COUNT").ToString()
                Top_Tank.IC_HVOP_X1.Text = reader("IC_HVOP_X1").ToString()
                Top_Tank.IC_HVOP_X2.Text = reader("IC_HVOP_X2").ToString()
                Top_Tank.IC_HVOP_Y1.Text = reader("IC_HVOP_Y1").ToString()
                Top_Tank.IC_HVOP_Y2.Text = reader("IC_HVOP_Y2").ToString()

                Top_Tank.IC_LV_REQ.SelectedItem = reader("IC_LV_REQ").ToString()
                Top_Tank.IC_LV_COUNT.SelectedItem = reader("IC_LV_COUNT").ToString()
                Top_Tank.IC_LV_X.Text = reader("IC_LV_X").ToString()
                Top_Tank.IC_LV_Y.Text = reader("IC_LV_Y").ToString()

                Top_Tank.IC_LVOP_REQ.SelectedItem = reader("IC_LVOP_REQ").ToString()
                Top_Tank.IC_LVOP_COUNT.SelectedItem = reader("IC_LVOP_COUNT").ToString()
                Top_Tank.IC_LVOP_X.Text = reader("IC_LVOP_X").ToString()
                Top_Tank.IC_LVOP_Y.Text = reader("IC_LVOP_Y").ToString()

                Top_Tank.IC_HV_TYPE.SelectedItem = reader("IC_HV_TYPE").ToString()
                Top_Tank.IC_HVOP_TYPE.SelectedItem = reader("IC_HVOP_TYPE").ToString()
                Top_Tank.IC_LV_TYPE.SelectedItem = reader("IC_LV_TYPE").ToString()
                Top_Tank.IC_LVOP_TYPE.SelectedItem = reader("IC_LVOP_TYPE").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Stiffner2_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Stiffner2_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.no_flat_stiff_U.SelectedItem = reader("no_flat_stiff_U").ToString()
                Top_Tank.flat_stiff1_U.Text = reader("flat_stiff1_U").ToString()
                Top_Tank.flat_stiff2_U.Text = reader("flat_stiff2_U").ToString()
                Top_Tank.flat_stiff3_U.Text = reader("flat_stiff3_U").ToString()
                Top_Tank.flat_stiff4_U.Text = reader("flat_stiff4_U").ToString()

                Top_Tank.flat_stiff5_U.Text = reader("flat_stiff5_U").ToString()
                Top_Tank.box_stiffner_H.Text = reader("box_stiffner_H").ToString()
                Top_Tank.box_stiffner_T.Text = reader("box_stiffner_T").ToString()
                Top_Tank.mid_insp_cvr_U.Text = reader("mid_insp_cvr_U").ToString()

                Top_Tank.top_insp_cvr_U.Text = reader("top_insp_cvr_U").ToString()
                Top_Tank.T2_D2HV.Text = reader("T2_D2HV").ToString()
                Top_Tank.T2_D3HV.Text = reader("T2_D3HV").ToString()
                Top_Tank.T2_D4HV.Text = reader("T2_D4HV").ToString()

                Top_Tank.flat_stiff5_V.Text = reader("flat_stiff5_V").ToString()
                Top_Tank.LVBSH1.Text = reader("LVBSH1").ToString()
                Top_Tank.LVBST1.Text = reader("LVBST1").ToString()
                Top_Tank.T2_LVSL1.Text = reader("T2_LVSL1").ToString()
                Top_Tank.box_stiffner_V2_H.Text = reader("box_stiffner_V2_H").ToString()
                Top_Tank.box_stiffner_V2_W.Text = reader("box_stiffner_V2_W").ToString()
                Top_Tank.box_stiffner_V2_T.Text = reader("box_stiffner_V2_T").ToString()
                Top_Tank.T2_LVSL2.Text = reader("T2_LVSL2").ToString()

                Top_Tank.top_insp_cvr_V.Text = reader("top_insp_cvr_V").ToString()
                Top_Tank.T2LVD2.Text = reader("T2LVD2").ToString()
                Top_Tank.T2LVD3.Text = reader("T2LVD3").ToString()
                Top_Tank.T2LVD4.Text = reader("T2LVD4").ToString()
                Top_Tank.T2LVD5.Text = reader("T2LVD5").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_TopCover_Stiffner_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopCover_Stiffner_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.top_cover_types.SelectedItem = reader("top_cover_types").ToString()
                Top_Tank.HV_W_dist_X_axis.Text = reader("HV_W_dist_X_axis").ToString()
                Top_Tank.OLTC_flange_dist_X_axis.Text = reader("OLTC_flange_dist_X_axis").ToString()
                Top_Tank.OLTC_flange_dist_Y_axis.Text = reader("OLTC_flange_dist_Y_axis").ToString()

                Top_Tank.TC_BSW.Text = reader("TC_BSW").ToString()
                Top_Tank.TC_BSH.Text = reader("TC_BSH").ToString()
                Top_Tank.TC_BST.Text = reader("TC_BST").ToString()

                Top_Tank.TC_FSH.Text = reader("TC_FSH").ToString()
                Top_Tank.TC_FST.Text = reader("TC_FST").ToString()

                Top_Tank.dist1.Text = reader("dist1").ToString()
                Top_Tank.dist2.Text = reader("dist2").ToString()
                Top_Tank.dist3.Text = reader("dist3").ToString()
                Top_Tank.dist4.Text = reader("dist4").ToString()
                Top_Tank.dist5.Text = reader("dist5").ToString()
                Top_Tank.FS_2.Text = reader("FS_2").ToString()
                Top_Tank.FS_3.Text = reader("FS_3").ToString()
                Top_Tank.FS_4.Text = reader("FS_4").ToString()
                Top_Tank.FS_5.Text = reader("FS_5").ToString()
                Top_Tank.FS_6.Text = reader("FS_6").ToString()

                Top_Tank.BUCHHOLZ_RELAY_REQ.SelectedItem = reader("BUCHHOLZ_RELAY_REQ").ToString()
                Top_Tank.BUCHHOLZ_RELAY_X.Text = reader("BUCHHOLZ_RELAY_X").ToString()
                Top_Tank.BUCHHOLZ_RELAY_Y.Text = reader("BUCHHOLZ_RELAY_Y").ToString()

                Top_Tank.NO_OF_PORT_TOP_HV.SelectedItem = reader("NO_OF_PORT_TOP_HV").ToString()
                Top_Tank.NO_OF_PORT_TOP_LV.SelectedItem = reader("NO_OF_PORT_TOP_LV").ToString()

                Cooler_Bank.ODAF_no_of_port.SelectedItem = reader("NO_OF_PORT_TOP_HV").ToString()
                Cooler_Bank.ODAF_no_of_port1.SelectedItem = reader("NO_OF_PORT_TOP_LV").ToString()

                Top_Tank.TopLiftingLeg_X1.Text = reader("TopLiftingLeg_X1").ToString()
                Top_Tank.TopLiftingLeg_X2.Text = reader("TopLiftingLeg_X2").ToString()
                Top_Tank.TopLiftingLeg_X3.Text = reader("TopLiftingLeg_X3").ToString()
                Top_Tank.TopLiftingLeg_X4.Text = reader("TopLiftingLeg_X4").ToString()

                Top_Tank.TopLiftingLeg_Y1.Text = reader("TopLiftingLeg_Y1").ToString()
                Top_Tank.TopLiftingLeg_Y2.Text = reader("TopLiftingLeg_Y2").ToString()
                Top_Tank.TopLiftingLeg_Y3.Text = reader("TopLiftingLeg_Y3").ToString()
                Top_Tank.TopLiftingLeg_Y4.Text = reader("TopLiftingLeg_Y4").ToString()

                Top_Tank.HV_PORT_X1.Text = reader("HV_PORT_X1").ToString()
                Top_Tank.HV_PORT_X2.Text = reader("HV_PORT_X2").ToString()
                Top_Tank.HV_PORT_X3.Text = reader("HV_PORT_X3").ToString()

                Top_Tank.HV_PORT_Y1.Text = reader("HV_PORT_Y1").ToString()
                Top_Tank.HV_PORT_Y2.Text = reader("HV_PORT_Y2").ToString()
                Top_Tank.HV_PORT_Y3.Text = reader("HV_PORT_Y3").ToString()

                Top_Tank.LV_PORT_X1.Text = reader("LV_PORT_X1").ToString()
                Top_Tank.LV_PORT_X2.Text = reader("LV_PORT_X2").ToString()
                Top_Tank.LV_PORT_X3.Text = reader("LV_PORT_X3").ToString()

                Top_Tank.LV_PORT_Y1.Text = reader("LV_PORT_Y1").ToString()
                Top_Tank.LV_PORT_Y2.Text = reader("LV_PORT_Y2").ToString()
                Top_Tank.LV_PORT_Y3.Text = reader("LV_PORT_Y3").ToString()

                Top_Tank.DM_BOX_DETAILS.SelectedItem = reader("DM_BOX_DETAILS").ToString()
                Top_Tank.DM_BOX_ANGLE.Text = reader("DM_BOX_ANGLE").ToString()

                Top_Tank.Tie_in_Resistor1.SelectedItem = reader("Tie_in_Resistor1").ToString()
                Top_Tank.Tie_in_Resistor2.Text = reader("Tie_in_Resistor2").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_Accessories1_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories1_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.PRD_mounting.SelectedItem = reader("PRD_mounting").ToString()

                Top_Tank.PRD_x1.Text = reader("PRD_x1").ToString()
                Top_Tank.PRD_Y1.Text = reader("PRD_Y1").ToString()
                Top_Tank.PRD_x2.Text = reader("PRD_x2").ToString()
                Top_Tank.PRD_Y2.Text = reader("PRD_Y2").ToString()
                Top_Tank.PRD_x3.Text = reader("PRD_x3").ToString()
                Top_Tank.PRD_Y3.Text = reader("PRD_Y3").ToString()
                Top_Tank.PRD_x4.Text = reader("PRD_x4").ToString()
                Top_Tank.PRD_Y4.Text = reader("PRD_Y4").ToString()

                Top_Tank.Inspection_cover.SelectedItem = reader("Inspection_cover").ToString()
                Top_Tank.ins_cover_size.SelectedItem = reader("ins_cover_size").ToString()

                Top_Tank.Inspection_cover_x1.Text = reader("Inspection_cover_x1").ToString()
                Top_Tank.Inspection_cover_y1.Text = reader("Inspection_cover_y1").ToString()
                Top_Tank.Inspection_cover_x2.Text = reader("Inspection_cover_x2").ToString()
                Top_Tank.Inspection_cover_y2.Text = reader("Inspection_cover_y2").ToString()
                Top_Tank.Inspection_cover_x3.Text = reader("Inspection_cover_x3").ToString()
                Top_Tank.Inspection_cover_y3.Text = reader("Inspection_cover_y3").ToString()
                Top_Tank.Inspection_cover_x4.Text = reader("Inspection_cover_x4").ToString()
                Top_Tank.Inspection_cover_y4.Text = reader("Inspection_cover_y4").ToString()

                Top_Tank.Buffer_Locking.SelectedItem = reader("Buffer_Locking").ToString()

                Top_Tank.Buffer_Locking_x1.Text = reader("Buffer_Locking_x1").ToString()
                Top_Tank.Buffer_Locking_y1.Text = reader("Buffer_Locking_y1").ToString()
                Top_Tank.Buffer_Locking_x2.Text = reader("Buffer_Locking_x2").ToString()
                Top_Tank.Buffer_Locking_y2.Text = reader("Buffer_Locking_y2").ToString()
                Top_Tank.Buffer_Locking_x3.Text = reader("Buffer_Locking_x3").ToString()
                Top_Tank.Buffer_Locking_y3.Text = reader("Buffer_Locking_y3").ToString()

                Top_Tank.Safety_pole.SelectedItem = reader("Safety_pole").ToString()

                Top_Tank.Safety_pole_x1.Text = reader("Safety_pole_x1").ToString()
                Top_Tank.Safety_pole_y1.Text = reader("Safety_pole_y1").ToString()
                Top_Tank.Safety_pole_x2.Text = reader("Safety_pole_x2").ToString()
                Top_Tank.Safety_pole_y2.Text = reader("Safety_pole_y2").ToString()
                Top_Tank.Safety_pole_x3.Text = reader("Safety_pole_x3").ToString()
                Top_Tank.Safety_pole_y3.Text = reader("Safety_pole_y3").ToString()
                Top_Tank.Safety_pole_x4.Text = reader("Safety_pole_x4").ToString()
                Top_Tank.Safety_pole_y4.Text = reader("Safety_pole_y4").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Accessories2_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories2_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.FOS_REQ.SelectedItem = reader("FOS_REQ").ToString()
                Top_Tank.FOS_MAKE.SelectedItem = reader("FOS_MAKE").ToString()
                Top_Tank.FOS_Channel_Mounting.SelectedItem = reader("FOS_Channel_Mounting").ToString()
                Top_Tank.FOS_Box_Mounting.SelectedItem = reader("FOS_Box_Mounting").ToString()
                Top_Tank.FOS_Inspection_cover.SelectedItem = reader("FOS_Inspection_cover").ToString()
                Top_Tank.FOS_X1.Text = reader("FOS_X1").ToString()
                Top_Tank.FOS_Y1.Text = reader("FOS_Y1").ToString()
                Top_Tank.FOS_X2.Text = reader("FOS_X2").ToString()
                Top_Tank.FOS_Y2.Text = reader("FOS_Y2").ToString()

                Top_Tank.CF_TE_X.Text = reader("CF_TE_X").ToString()
                Top_Tank.CF_TE_Y.Text = reader("CF_TE_Y").ToString()
                Top_Tank.SPRR_MOUNTING_X.Text = reader("SPRR_MOUNTING_X").ToString()
                Top_Tank.SPRR_MOUNTING_Y.Text = reader("SPRR_MOUNTING_Y").ToString()
                Top_Tank.GOR_CP_X.Text = reader("GOR_CP_X").ToString()
                Top_Tank.GOR_CP_Y.Text = reader("GOR_CP_Y").ToString()


                Top_Tank.Thermosyphon_filter_Req.SelectedItem = reader("Thermosyphon_filter_Req").ToString()
                Top_Tank.Thermosyphon_filter_Make.SelectedItem = reader("Thermosyphon_filter_Make").ToString()
                Top_Tank.Thermosyphon_filter_Size.SelectedItem = reader("Thermosyphon_filter_Size").ToString()
                Top_Tank.Thermosyphon_filter_POS.SelectedItem = reader("Thermosyphon_filter_POS").ToString()
                Top_Tank.Thermo_X.Text = reader("Thermo_X").ToString()
                Top_Tank.Thermo_Y.Text = reader("Thermo_Y").ToString()
                Top_Tank.Thermosyphon_filter_Mounting.SelectedItem = reader("Thermosyphon_filter_Mounting").ToString()
                Top_Tank.Thermosyphon_filter_X.Text = reader("Thermosyphon_filter_X").ToString()
                Top_Tank.Thermosyphon_filter_Y.Text = reader("Thermosyphon_filter_Y").ToString()

                Top_Tank.N2_cylinder_REQ.SelectedItem = reader("N2_cylinder_REQ").ToString()
                Top_Tank.N2_cylinder_QTY.SelectedItem = reader("N2_cylinder_QTY").ToString()
                Top_Tank.N2_cylinder_MP.SelectedItem = reader("N2_cylinder_MP").ToString()
                Top_Tank.N2_cylinder_AR.SelectedItem = reader("N2_cylinder_AR").ToString()
                Top_Tank.N2_X.Text = reader("N2_X").ToString()
                Top_Tank.N2_Y.Text = reader("N2_Y").ToString()
                Top_Tank.N2_X1.Text = reader("N2_X1").ToString()
                Top_Tank.N2_Y1.Text = reader("N2_Y1").ToString()

                Top_Tank.Marshallin_box_Mounting.SelectedItem = reader("Marshallin_box_Mounting").ToString()
                Top_Tank.Marshallin_box_X1.Text = reader("Marshallin_box_X1").ToString()
                Top_Tank.Marshallin_box_Y1.Text = reader("Marshallin_box_Y1").ToString()

                Top_Tank.FOS_MTG_REQ.SelectedItem = reader("FOS_MTG_REQ").ToString()
                Top_Tank.FOS_MTG_POS.SelectedItem = reader("FOS_MTG_POS").ToString()
                Top_Tank.FOS_MTG_X.Text = reader("FOS_MTG_X").ToString()
                Top_Tank.FOS_MTG_Y.Text = reader("FOS_MTG_Y").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_Accessories3_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories3_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.Thermometer_pocket.SelectedItem = reader("Thermometer_pocket").ToString()

                Top_Tank.Thermometer_pocket_X1.Text = reader("Thermometer_pocket_X1").ToString()
                Top_Tank.Thermometer_pocket_Y1.Text = reader("Thermometer_pocket_Y1").ToString()

                Top_Tank.Vaccum_braking_valve.SelectedItem = reader("Vaccum_braking_valve").ToString()
                Top_Tank.Vaccum_braking_valve_X.Text = reader("Vaccum_braking_valve_X").ToString()
                Top_Tank.Vaccum_braking_valve_Y.Text = reader("Vaccum_braking_valve_Y").ToString()

                Top_Tank.Air_release_plug.SelectedItem = reader("Air_release_plug").ToString()
                Top_Tank.Air_release_plug_X1.Text = reader("Air_release_plug_X1").ToString()
                Top_Tank.Thermo_Y.Text = reader("Air_release_plug_Y1").ToString()
                Top_Tank.Air_release_plug_X2.Text = reader("Air_release_plug_X2").ToString()
                Top_Tank.Air_release_plug_Y2.Text = reader("Air_release_plug_Y2").ToString()

                Top_Tank.NIFPES_sensor_box.SelectedItem = reader("NIFPES_sensor_box").ToString()
                Top_Tank.NIFPES_sensor_box_X.Text = reader("NIFPES_sensor_box_X").ToString()
                Top_Tank.NIFPES_sensor_box_Y.Text = reader("NIFPES_sensor_box_Y").ToString()

                Top_Tank.quick_drain_val_1.SelectedItem = reader("quick_drain_val_1").ToString()
                Top_Tank.quick_drain_val_x.Text = reader("quick_drain_val_x").ToString()
                Top_Tank.quick_drain_val_y.Text = reader("quick_drain_val_y").ToString()

                Top_Tank.NIFPES_Mounting1.SelectedItem = reader("NIFPES_Mounting1").ToString()
                Top_Tank.NIFPES_Mounting_X.Text = reader("NIFPES_Mounting_X").ToString()
                Top_Tank.NIFPES_Mounting_Y.Text = reader("NIFPES_Mounting_Y").ToString()

                Top_Tank.Impact_recorder_req.SelectedItem = reader("Impact_recorder_req").ToString()
                Top_Tank.Impact_recorder_qty.SelectedItem = reader("Impact_recorder_qty").ToString()
                Top_Tank.Impact_recorder_mounting.SelectedItem = reader("Impact_recorder_mounting").ToString()
                Top_Tank.Impact_recorder_X.Text = reader("Impact_recorder_X").ToString()
                Top_Tank.Impact_recorder_Y.Text = reader("Impact_recorder_Y").ToString()

                Top_Tank.OLTC_CONSERVATOR_REQ.SelectedItem = reader("OLTC_CONSERVATOR_REQ").ToString()
                Top_Tank.OLTC_CONSERVATOR_POS.SelectedItem = reader("OLTC_CONSERVATOR_POS").ToString()
                Top_Tank.OLTC_CONSERVATOR_X.Text = reader("OLTC_CONSERVATOR_X").ToString()
                Top_Tank.OLTC_CONSERVATOR_Y.Text = reader("OLTC_CONSERVATOR_Y").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Accessories4_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accessories4_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.ODS_REQ.SelectedItem = reader("ODS_REQ").ToString()
                Top_Tank.ODS_MAKE.SelectedItem = reader("ODS_MAKE").ToString()
                Top_Tank.ODS_SIZE.SelectedItem = reader("ODS_SIZE").ToString()
                Top_Tank.ODS_POS.SelectedItem = reader("ODS_POS").ToString()
                Top_Tank.ODS_X1.Text = reader("ODS_X1").ToString()
                Top_Tank.ODS_Y1.Text = reader("ODS_Y1").ToString()
                Top_Tank.ODS_Y2.Text = reader("ODS_Y2").ToString()
                Top_Tank.ODS_MOUNTING.SelectedItem = reader("ODS_MOUNTING").ToString()
                Top_Tank.ODS_X.Text = reader("ODS_X").ToString()
                Top_Tank.ODS_Y.Text = reader("ODS_Y").ToString()

                Top_Tank.DGA_REQ.SelectedItem = reader("DGA_REQ").ToString()
                Top_Tank.DGA_MAKE.SelectedItem = reader("DGA_MAKE").ToString()
                Top_Tank.DGA_SIZE.SelectedItem = reader("DGA_SIZE").ToString()
                Top_Tank.DGA_POS.SelectedItem = reader("DGA_POS").ToString()
                Top_Tank.DGA_X1.Text = reader("DGA_X1").ToString()
                Top_Tank.DGA_Y1.Text = reader("DGA_Y1").ToString()
                Top_Tank.DGA_Y2.Text = reader("DGA_Y2").ToString()
                Top_Tank.DGA_MOUNTING.SelectedItem = reader("DGA_MOUNTING").ToString()
                Top_Tank.DGA_X.Text = reader("DGA_X").ToString()
                Top_Tank.DGA_Y.Text = reader("DGA_Y").ToString()


                Top_Tank.UHF_Sensor_Valve_REQ.SelectedItem = reader("UHF_Sensor_Valve_REQ").ToString()
                Top_Tank.UHF_Sensor_Valve_MAKE.SelectedItem = reader("UHF_Sensor_Valve_MAKE").ToString()
                Top_Tank.UHF_Sensor_Valve_SIZE.SelectedItem = reader("UHF_Sensor_Valve_SIZE").ToString()
                Top_Tank.UHF_Sensor_Valve_POS.SelectedItem = reader("UHF_Sensor_Valve_POS").ToString()
                Top_Tank.UHF_QTY.SelectedItem = reader("UHF_QTY").ToString()
                Top_Tank.UHF_Y1.Text = reader("UHF_Y1").ToString()
                Top_Tank.UHF_X1.Text = reader("UHF_X1").ToString()
                Top_Tank.UHF_X2.Text = reader("UHF_X2").ToString()
                Top_Tank.UHF_X3.Text = reader("UHF_X3").ToString()
                Top_Tank.UHF_Sensor_Valve_Mounting.SelectedItem = reader("UHF_Sensor_Valve_Mounting").ToString()
                Top_Tank.UHF_Sensor_Valve_X.Text = reader("UHF_Sensor_Valve_X").ToString()
                Top_Tank.UHF_Sensor_Valve_Y.Text = reader("UHF_Sensor_Valve_Y").ToString()

                Top_Tank.Moisture_sensor_REQ.SelectedItem = reader("Moisture_sensor_REQ").ToString()
                Top_Tank.Moisture_sensor_MAKE.SelectedItem = reader("Moisture_sensor_MAKE").ToString()
                Top_Tank.Moisture_sensor_SIZE.SelectedItem = reader("Moisture_sensor_SIZE").ToString()
                Top_Tank.Moisture_sensor_POS.SelectedItem = reader("Moisture_sensor_POS").ToString()
                Top_Tank.Sensor_X.Text = reader("Sensor_X").ToString()
                Top_Tank.Sensor_Y.Text = reader("Sensor_Y").ToString()
                Top_Tank.Moisture_sensor_Mounting.SelectedItem = reader("Moisture_sensor_Mounting").ToString()
                Top_Tank.Moisture_sensor_X.Text = reader("Moisture_sensor_X").ToString()
                Top_Tank.Moisture_sensor_Y.Text = reader("Moisture_sensor_Y").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Accessories5_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Accesories5_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.MPS_REQ.SelectedItem = reader("MPS_REQ").ToString()
                Top_Tank.MPS_POS.SelectedItem = reader("MPS_POS").ToString()
                Top_Tank.MPS_X.Text = reader("MPS_X").ToString()
                Top_Tank.MPS_Y.Text = reader("MPS_Y").ToString()

                Top_Tank.LBA_REQ.SelectedItem = reader("LBA_REQ").ToString()
                Top_Tank.LBA_HV_COUNT.SelectedItem = reader("LBA_HV_COUNT").ToString()
                Top_Tank.LBA_HVX1.Text = reader("LBA_HVX1").ToString()
                Top_Tank.LBA_HVX2.Text = reader("LBA_HVX2").ToString()
                Top_Tank.LBA_HVY1.Text = reader("LBA_HVY1").ToString()
                Top_Tank.LBA_HVY2.Text = reader("LBA_HVY2").ToString()
                Top_Tank.LBA_LV_COUNT.SelectedItem = reader("LBA_LV_COUNT").ToString()
                Top_Tank.LBA_HV_X1.Text = reader("LBA_HV_X1").ToString()
                Top_Tank.LBA_HV_X2.Text = reader("LBA_HV_X2").ToString()
                Top_Tank.LBA_HV_Y1.Text = reader("LBA_HV_Y1").ToString()
                Top_Tank.LBA_HV_Y2.Text = reader("LBA_HV_Y2").ToString()

                Top_Tank.TOLG_REQ.SelectedItem = reader("TOLG_REQ").ToString()
                Top_Tank.TOLG_POS.SelectedItem = reader("TOLG_POS").ToString()
                Top_Tank.TOLG_X.Text = reader("TOLG_X").ToString()
                Top_Tank.TOLG_Y.Text = reader("TOLG_Y").ToString()

                Top_Tank.GCDM_REQ.SelectedItem = reader("GCDM_REQ").ToString()
                Top_Tank.GCDM_POS.SelectedItem = reader("GCDM_POS").ToString()
                Top_Tank.GCDM_X.Text = reader("GCDM_X").ToString()
                Top_Tank.GCDM_Y.Text = reader("GCDM_Y").ToString()

                Top_Tank.TFV_REQ.SelectedItem = reader("TFV_REQ").ToString()
                Top_Tank.TFV_POS.SelectedItem = reader("TFV_POS").ToString()
                Top_Tank.TFV_X.Text = reader("TFV_X").ToString()
                Top_Tank.TFV_Y.Text = reader("TFV_Y").ToString()

                Top_Tank.TSV_REQ.SelectedItem = reader("TSV_REQ").ToString()
                Top_Tank.TSV_POS.SelectedItem = reader("TSV_POS").ToString()
                Top_Tank.TSV_X.Text = reader("TSV_X").ToString()
                Top_Tank.TSV_Y.Text = reader("TSV_Y").ToString()

                Top_Tank.LBA_TYPE.SelectedItem = reader("LBA_TYPE").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub
    Public Sub Load_TVTermination_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TVTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                Top_Tank.TV_Termination.SelectedItem = reader("TV_Termination").ToString()
                Top_Tank.TV_Termination_Mounting.SelectedItem = reader("TV_Termination_Mounting").ToString()

                'U
                Top_Tank.TV_bush_rating_U.SelectedItem = reader("TV_bush_rating_U").ToString()
                Top_Tank.TV_bush_BCT_U.SelectedItem = reader("TV_bush_BCT_U").ToString()
                Top_Tank.TV_turret_type_U.SelectedItem = reader("TV_turret_type_U").ToString()
                Top_Tank.TV_turret_dia_U.SelectedItem = reader("TV_turret_dia_U").ToString()
                Top_Tank.TV_turret_ht_U.SelectedItem = reader("TV_turret_ht_U").ToString()

                Top_Tank.TV_ht_CT_MP_U.Text = reader("TV_ht_CT_MP_U").ToString()
                Top_Tank.TV_ht_CT_box_U.Text = reader("TV_ht_CT_box_U").ToString()
                Top_Tank.TV_ID_CT_MP_U.Text = reader("TV_ID_CT_MP_U").ToString()

                Top_Tank.TV_bush_MTG_U.SelectedItem = reader("TV_bush_MTG_U").ToString()
                Top_Tank.TV_CT_terminal_box_U.SelectedItem = reader("TV_CT_terminal_box_U").ToString()

                Top_Tank.TV_CT_terminal_box_ang_U.Text = reader("TV_CT_terminal_box_ang_U").ToString()

                Top_Tank.TV_inspection_cover_U.SelectedItem = reader("TV_inspection_cover_U").ToString()

                Top_Tank.TV_inspection_cover_ang_U.Text = reader("TV_inspection_cover_ang_U").ToString()
                Top_Tank.TV_equalizing_cover_U.Text = reader("TV_equalizing_cover_U").ToString()
                Top_Tank.TV_welded_trunk_ht_U.Text = reader("TV_welded_trunk_ht_U").ToString()

                Top_Tank.TV_ht_turre_body_U.SelectedItem = reader("TV_ht_turre_body_U").ToString()

                Top_Tank.TV_ht_small_turret_body_U.Text = reader("TV_ht_small_turret_body_U").ToString()

                Top_Tank.TV_turret_angle_U.Text = reader("TV_turret_angle_U").ToString()

                'V1
                Top_Tank.TV_bush_rating_V1.SelectedItem = reader("TV_bush_rating_V1").ToString()
                Top_Tank.TV_bush_BCT_V1.SelectedItem = reader("TV_bush_BCT_V1").ToString()
                Top_Tank.TV_turret_type_V1.SelectedItem = reader("TV_turret_type_V1").ToString()
                Top_Tank.TV_turret_dia_V1.SelectedItem = reader("TV_turret_dia_V1").ToString()
                Top_Tank.TV_turret_ht_V1.SelectedItem = reader("TV_turret_ht_V1").ToString()

                Top_Tank.TV_ht_CT_MP_V1.Text = reader("TV_ht_CT_MP_V1").ToString()
                Top_Tank.TV_ht_CT_box_V1.Text = reader("TV_ht_CT_box_V1").ToString()
                Top_Tank.TV_ID_CT_MP_V1.Text = reader("TV_ID_CT_MP_V1").ToString()

                Top_Tank.TV_bush_MTG_V1.Text = reader("TV_bush_MTG_V1").ToString()
                Top_Tank.TV_CT_terminal_box_V1.SelectedItem = reader("TV_CT_terminal_box_V1").ToString()

                Top_Tank.TV_CT_terminal_box_ang_V1.Text = reader("TV_CT_terminal_box_ang_V1").ToString()

                Top_Tank.TV_inspection_cover_V1.Text = reader("TV_inspection_cover_V1").ToString()

                Top_Tank.TV_inspection_cover_ang_V1.Text = reader("TV_inspection_cover_ang_V1").ToString()
                Top_Tank.TV_equalizing_cover_V1.Text = reader("TV_equalizing_cover_V1").ToString()
                Top_Tank.TV_welded_trunk_ht_V1.Text = reader("TV_welded_trunk_ht_V1").ToString()

                Top_Tank.TV_ht_turre_body_V1.SelectedItem = reader("TV_ht_turre_body_V1").ToString()

                Top_Tank.TV_ht_small_turret_body_V1.Text = reader("TV_ht_small_turret_body_V1").ToString()

                Top_Tank.TV_turret_angle_V1.Text = reader("TV_turret_angle_V1").ToString()

                'V2
                Top_Tank.TV_bush_rating_V2.SelectedItem = reader("TV_bush_rating_V2").ToString()
                Top_Tank.TV_bush_BCT_V2.SelectedItem = reader("TV_bush_BCT_V2").ToString()
                Top_Tank.TV_turret_type_V2.SelectedItem = reader("TV_turret_type_V2").ToString()
                Top_Tank.TV_turret_dia_V2.SelectedItem = reader("TV_turret_dia_V2").ToString()
                Top_Tank.TV_turret_ht_V2.SelectedItem = reader("TV_turret_ht_V2").ToString()

                Top_Tank.TV_ht_CT_MP_V2.Text = reader("TV_ht_CT_MP_V2").ToString()
                Top_Tank.TV_ht_CT_box_V2.Text = reader("TV_ht_CT_box_V2").ToString()
                Top_Tank.TV_ID_CT_MP_V2.Text = reader("TV_ID_CT_MP_V2").ToString()

                Top_Tank.TV_bush_MTG_V2.Text = reader("TV_bush_MTG_V2").ToString()
                Top_Tank.TV_CT_terminal_box_V2.SelectedItem = reader("TV_CT_terminal_box_V2").ToString()

                Top_Tank.TV_CT_terminal_box_ang_V2.Text = reader("TV_CT_terminal_box_ang_V2").ToString()

                Top_Tank.TV_inspection_cover_V2.SelectedItem = reader("TV_inspection_cover_V2").ToString()

                Top_Tank.TV_inspection_cover_ang_V2.Text = reader("TV_inspection_cover_ang_V2").ToString()
                Top_Tank.TV_equalizing_cover_V2.Text = reader("TV_equalizing_cover_V2").ToString()
                Top_Tank.TV_welded_trunk_ht_V2.Text = reader("TV_welded_trunk_ht_V2").ToString()

                Top_Tank.TV_ht_turre_body_V2.SelectedItem = reader("TV_ht_turre_body_V2").ToString()

                Top_Tank.TV_ht_small_turret_body_V2.Text = reader("TV_ht_small_turret_body_V2").ToString()

                Top_Tank.TV_turret_angle_V2.Text = reader("TV_turret_angle_V2").ToString()


                'V1
                Top_Tank.TV_bush_rating_W.SelectedItem = reader("TV_bush_rating_W").ToString()
                Top_Tank.TV_bush_BCT_W.SelectedItem = reader("TV_bush_BCT_W").ToString()
                Top_Tank.TV_turret_type_W.SelectedItem = reader("TV_turret_type_W").ToString()
                Top_Tank.TV_turret_dia_W.SelectedItem = reader("TV_turret_dia_W").ToString()
                Top_Tank.TV_turret_ht_W.SelectedItem = reader("TV_turret_ht_W").ToString()

                Top_Tank.TV_ht_CT_MP_W.Text = reader("TV_ht_CT_MP_W").ToString()
                Top_Tank.TV_ht_CT_box_W.Text = reader("TV_ht_CT_box_W").ToString()
                Top_Tank.TV_ID_CT_MP_W.Text = reader("TV_ID_CT_MP_W").ToString()

                Top_Tank.TV_bush_MTG_W.Text = reader("TV_bush_MTG_W").ToString()
                Top_Tank.TV_CT_terminal_box_W.SelectedItem = reader("TV_CT_terminal_box_W").ToString()

                Top_Tank.TV_CT_terminal_box_ang_W.Text = reader("TV_CT_terminal_box_ang_W").ToString()

                Top_Tank.TV_inspection_cover_W.SelectedItem = reader("TV_inspection_cover_W").ToString()

                Top_Tank.TV_inspection_cover_ang_W.Text = reader("TV_inspection_cover_ang_W").ToString()
                Top_Tank.TV_equalizing_cover_W.Text = reader("TV_equalizing_cover_W").ToString()
                Top_Tank.TV_welded_trunk_ht_W.Text = reader("TV_welded_trunk_ht_W").ToString()

                Top_Tank.TV_ht_turre_body_W.SelectedItem = reader("TV_ht_turre_body_W").ToString()

                Top_Tank.TV_ht_small_turret_body_W.Text = reader("TV_ht_small_turret_body_W").ToString()

                Top_Tank.TV_turret_angle_W.Text = reader("TV_turret_angle_W").ToString()
            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_CoolerBank_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TVTermination_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.CB_MOUNT.SelectedItem = reader("CB_MOUNT").ToString()
                Cooler_Bank.cooler_bnk.SelectedItem = reader("cooler_bnk").ToString()
                Cooler_Bank.cooler_bnk_pos1.SelectedItem = reader("cooler_bnk_pos1").ToString()
                Cooler_Bank.cooler_bnk_pos2.SelectedItem = reader("cooler_bnk_pos2").ToString()
                Cooler_Bank.cooler_bnk_pos3.SelectedItem = reader("cooler_bnk_pos3").ToString()
                Cooler_Bank.typ_of_cooling.SelectedItem = reader("typ_of_cooling").ToString()
                Cooler_Bank.header_pipe_size_bot.SelectedItem = reader("header_pipe_size_bot").ToString()

                'ONAN
                Cooler_Bank.ONAN_butterfly_valve_size.SelectedItem = reader("header_pipe_size_bot").ToString()
                'ONAF
                Cooler_Bank.BF_VL_ONAF.SelectedItem = reader("header_pipe_size_bot").ToString()
                'ODAF
                Cooler_Bank.ODAF_butterfly_valve_size.SelectedItem = reader("header_pipe_size_bot").ToString()

                Cooler_Bank.Radiator_size_C_C.Text = reader("Radiator_size_C_C").ToString()
                Cooler_Bank.Radiator_size_FINS.Text = reader("Radiator_size_FINS").ToString()
                Cooler_Bank.Radiator_size_W.Text = reader("Radiator_size_W").ToString()
                Cooler_Bank.Radiator_size_QTY.Text = reader("Radiator_size_QTY").ToString()

                Cooler_Bank.Radiator_pipe_size.SelectedItem = reader("Radiator_pipe_size").ToString()
                Cooler_Bank.Ground_Support_Reg.SelectedItem = reader("Ground_Support_Reg").ToString()
                Cooler_Bank.Radiator_drain1.SelectedItem = reader("Radiator_drain1").ToString()
                Cooler_Bank.Radiator_drain2.SelectedItem = reader("Radiator_drain2").ToString()

                Cooler_Bank.CB_TopCP_POS.SelectedItem = reader("CB_TopCP_POS").ToString()
                Cooler_Bank.CB_BottomCP_POS.SelectedItem = reader("CB_BottomCP_POS").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_TopCooler1_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopCooler1_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.cooler_pipe_size.SelectedItem = reader("cooler_pipe_size").ToString()
                Cooler_Bank.butterfly_valve_size.SelectedItem = reader("butterfly_valve_size").ToString()
                Cooler_Bank.expansion_bellows.SelectedItem = reader("expansion_bellows").ToString()
                Cooler_Bank.expansion_bellows_size.SelectedItem = reader("expansion_bellows_size").ToString()

                Cooler_Bank.no_of_port.SelectedItem = reader("no_of_port").ToString()
                Cooler_Bank.TCP_HV_X1.Text = reader("TCP_HV_X1").ToString()
                Cooler_Bank.TCP_HV_X2.Text = reader("TCP_HV_X2").ToString()
                Cooler_Bank.TCP_HV_X3.Text = reader("TCP_HV_X3").ToString()
                Cooler_Bank.TCP_HV_Y1.Text = reader("TCP_HV_Y1").ToString()
                Cooler_Bank.TCP_HV_Y2.Text = reader("TCP_HV_Y2").ToString()
                Cooler_Bank.TCP_HV_Y3.Text = reader("TCP_HV_Y3").ToString()
                Cooler_Bank.Z_1.Text = reader("Z_1").ToString()
                Cooler_Bank.Z_2.Text = reader("Z_2").ToString()
                Cooler_Bank.Z_3.Text = reader("Z_3").ToString()
                Cooler_Bank.Z_4.Text = reader("Z_4").ToString()


                Cooler_Bank.TCP_LV_X1.Text = reader("TCP_LV_X1").ToString()
                Cooler_Bank.TCP_LV_X2.Text = reader("TCP_LV_X2").ToString()
                Cooler_Bank.TCP_LV_Y1.Text = reader("TCP_LV_Y1").ToString()


                Cooler_Bank.NO_SUP.SelectedItem = reader("NO_SUP").ToString()
                Cooler_Bank.TCP_LV_X3.Text = reader("TCP_LV_X3").ToString()
                Cooler_Bank.TCP_LV_Y3.Text = reader("TCP_LV_Y3").ToString()
                Cooler_Bank.TCP_LV_Y4.Text = reader("TCP_LV_Y4").ToString()
                Cooler_Bank.TCP_LV_Y5.Text = reader("TCP_LV_Y5").ToString()
                Cooler_Bank.TCP_LV_Y6.Text = reader("TCP_LV_Y6").ToString()
                Cooler_Bank.V_1.Text = reader("V_1").ToString()
                Cooler_Bank.H_U.Text = reader("H_U").ToString()
                Cooler_Bank.H_V.Text = reader("H_V").ToString()
                Cooler_Bank.RCC.Text = reader("RCC").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_TopCooler2_TT()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopCooler2_TT WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.B2_CPSIZE.SelectedItem = reader("B2_CPSIZE").ToString()
                Cooler_Bank.B2_BVS.SelectedItem = reader("B2_BVS").ToString()
                Cooler_Bank.B2_EXPBLW.SelectedItem = reader("B2_EXPBLW").ToString()
                Cooler_Bank.B2_EXPSIZE.SelectedItem = reader("B2_EXPSIZE").ToString()

                Cooler_Bank.B2_NO_P.SelectedItem = reader("B2_NO_P").ToString()
                Cooler_Bank.B2_X1.Text = reader("B2_X1").ToString()
                Cooler_Bank.B2_X2.Text = reader("B2_X2").ToString()
                Cooler_Bank.B2_X3.Text = reader("B2_X3").ToString()
                Cooler_Bank.B2_Y1.Text = reader("B2_Y1").ToString()
                Cooler_Bank.B2_Y2.Text = reader("B2_Y2").ToString()
                Cooler_Bank.B2_Y3.Text = reader("B2_Y3").ToString()
                Cooler_Bank.B2_Z1.Text = reader("B2_Z1").ToString()
                Cooler_Bank.B2_Z2.Text = reader("B2_Z2").ToString()
                Cooler_Bank.B2_Z3.Text = reader("B2_Z3").ToString()
                Cooler_Bank.B2_Z4.Text = reader("B2_Z4").ToString()


                Cooler_Bank.B2_BX1.Text = reader("B2_BX1").ToString()
                Cooler_Bank.B2_BX2.Text = reader("B2_BX2").ToString()
                Cooler_Bank.B2_BY.Text = reader("B2_BY").ToString()


                Cooler_Bank.B2_NO_GSUP.SelectedItem = reader("B2_NO_GSUP").ToString()
                Cooler_Bank.B2_D1.Text = reader("B2_D1").ToString()
                Cooler_Bank.B2_D2.Text = reader("B2_D2").ToString()
                Cooler_Bank.B2_D3.Text = reader("B2_D3").ToString()
                Cooler_Bank.B2_D4.Text = reader("B2_D4").ToString()
                Cooler_Bank.B2_Y.Text = reader("B2_Y").ToString()
                Cooler_Bank.B2_V1.Text = reader("B2_V1").ToString()
                Cooler_Bank.B2_U.Text = reader("B2_U").ToString()
                Cooler_Bank.B2_V.Text = reader("B2_V").ToString()
                Cooler_Bank.B2_R_C_C.Text = reader("B2_R_C_C").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_TopCooler3_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopCooler3_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.B3_CPSIZE.SelectedItem = reader("B3_CPSIZE").ToString()
                Cooler_Bank.B3_BVS.SelectedItem = reader("B3_BVS").ToString()
                Cooler_Bank.B3_EXPBLW.SelectedItem = reader("B3_EXPBLW").ToString()
                Cooler_Bank.B3_EXPSIZE.SelectedItem = reader("B3_EXPSIZE").ToString()

                Cooler_Bank.B3_NO_P.SelectedItem = reader("B3_NO_P").ToString()
                Cooler_Bank.B3_X1.Text = reader("B3_X1").ToString()
                Cooler_Bank.B3_X2.Text = reader("B3_X2").ToString()
                Cooler_Bank.B3_X3.Text = reader("B3_X3").ToString()
                Cooler_Bank.B3_Y1.Text = reader("B3_Y1").ToString()
                Cooler_Bank.B3_Y2.Text = reader("B3_Y2").ToString()
                Cooler_Bank.B3_Y3.Text = reader("B3_Y3").ToString()
                Cooler_Bank.B3_Z1.Text = reader("B3_Z1").ToString()
                Cooler_Bank.B3_Z2.Text = reader("B3_Z2").ToString()
                Cooler_Bank.B3_Z3.Text = reader("B3_Z3").ToString()
                Cooler_Bank.B3_Z4.Text = reader("B3_Z4").ToString()


                Cooler_Bank.B3_BX1.Text = reader("B3_BX1").ToString()
                Cooler_Bank.B3_BX2.Text = reader("B3_BX2").ToString()
                Cooler_Bank.B3_BY.Text = reader("B3_BY").ToString()


                Cooler_Bank.B3_NO_GSUP.SelectedItem = reader("B3_NO_GSUP").ToString()
                Cooler_Bank.B3_D1.Text = reader("B3_D1").ToString()
                Cooler_Bank.B3_D2.Text = reader("B3_D2").ToString()
                Cooler_Bank.B3_D3.Text = reader("B3_D3").ToString()
                Cooler_Bank.B3_D4.Text = reader("B3_D4").ToString()
                Cooler_Bank.B3_Y.Text = reader("B3_Y").ToString()
                Cooler_Bank.B3_V1.Text = reader("B3_V1").ToString()
                Cooler_Bank.B3_U.Text = reader("B3_U").ToString()
                Cooler_Bank.B3_V.Text = reader("B3_V").ToString()
                Cooler_Bank.B3_R_C_C.Text = reader("B3_R_C_C").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_TopBotHeader_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopBotHeader_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.HP_SIZE_HV.SelectedItem = reader("HP_SIZE_HV").ToString()
                Cooler_Bank.RAD_CC_HV.Text = reader("RAD_CC_HV").ToString()
                Cooler_Bank.RAD_P_SIZE_HV.SelectedItem = reader("RAD_P_SIZE_HV").ToString()
                Cooler_Bank.N0_RAD_HV.SelectedItem = reader("N0_RAD_HV").ToString()
                Cooler_Bank.AIR_R_PG_HV.SelectedItem = reader("AIR_R_PG_HV").ToString()
                Cooler_Bank.T_P_HV.SelectedItem = reader("T_P_HV").ToString()
                Cooler_Bank.FCD_HV.SelectedItem = reader("FCD_HV").ToString()
                Cooler_Bank.AIR_PLUG_HV.Text = reader("AIR_PLUG_HV").ToString()
                Cooler_Bank.TM_POC_VAL_HV.Text = reader("TM_POC_VAL_HV").ToString()
                Cooler_Bank.FCD_VAL_HV.Text = reader("FCD_VAL_HV").ToString()
                Cooler_Bank.CP_LEN_HV.Text = reader("CP_LEN_HV").ToString()
                Cooler_Bank.HP_LEN_HV.Text = reader("HP_LEN_HV").ToString()

                Cooler_Bank.HP_SIZE_LV.SelectedItem = reader("HP_SIZE_LV").ToString()
                Cooler_Bank.RAD_CC_LV.Text = reader("RAD_CC_LV").ToString()
                Cooler_Bank.RAD_P_SIZE_LV.SelectedItem = reader("RAD_P_SIZE_LV").ToString()
                Cooler_Bank.N0_RAD_LV.SelectedItem = reader("N0_RAD_LV").ToString()
                Cooler_Bank.AIR_R_PG_LV.SelectedItem = reader("AIR_R_PG_LV").ToString()
                Cooler_Bank.T_P_LV.SelectedItem = reader("T_P_LV").ToString()
                Cooler_Bank.FCD_LV.SelectedItem = reader("FCD_LV").ToString()
                Cooler_Bank.AIR_PLUG_LV.Text = reader("AIR_PLUG_LV").ToString()
                Cooler_Bank.TM_POC_VAL_LV.Text = reader("TM_POC_VAL_LV").ToString()
                Cooler_Bank.FCD_VAL_LV.Text = reader("FCD_VAL_LV").ToString()
                Cooler_Bank.CP_LEN_LV.Text = reader("CP_LEN_LV").ToString()
                Cooler_Bank.HP_LEN_LV.Text = reader("HP_LEN_LV").ToString()

                Cooler_Bank.HP_SIZE_OOS.SelectedItem = reader("HP_SIZE_OOS").ToString()
                Cooler_Bank.RAD_CC_OOS.Text = reader("RAD_CC_OOS").ToString()
                Cooler_Bank.RAD_P_SIZE_OOS.SelectedItem = reader("RAD_P_SIZE_OOS").ToString()
                Cooler_Bank.N0_RAD_OOS.SelectedItem = reader("N0_RAD_OOS").ToString()
                Cooler_Bank.AIR_R_PG_OOS.SelectedItem = reader("AIR_R_PG_OOS").ToString()
                Cooler_Bank.T_P_OOS.SelectedItem = reader("T_P_OOS").ToString()
                Cooler_Bank.FCD_OOS.SelectedItem = reader("FCD_OOS").ToString()
                Cooler_Bank.AIR_PLUG_OOS.Text = reader("AIR_PLUG_OOS").ToString()
                Cooler_Bank.TM_POC_VAL_OOS.Text = reader("TM_POC_VAL_OOS").ToString()
                Cooler_Bank.FCD_VAL_OOS.Text = reader("FCD_VAL_OOS").ToString()
                Cooler_Bank.CP_LEN_OOS.Text = reader("CP_LEN_OOS").ToString()
                Cooler_Bank.HP_LEN_OOS.Text = reader("HP_LEN_OOS").ToString()

                Cooler_Bank.HP_SIZE_OS.SelectedItem = reader("HP_SIZE_OS").ToString()
                Cooler_Bank.RAD_CC_OS.Text = reader("RAD_CC_OS").ToString()
                Cooler_Bank.RAD_P_SIZE_OS.SelectedItem = reader("RAD_P_SIZE_OS").ToString()
                Cooler_Bank.N0_RAD_OS.SelectedItem = reader("N0_RAD_OS").ToString()
                Cooler_Bank.AIR_R_PG_OS.SelectedItem = reader("AIR_R_PG_OS").ToString()
                Cooler_Bank.T_P_OS.SelectedItem = reader("T_P_OS").ToString()
                Cooler_Bank.FCD_OS.SelectedItem = reader("FCD_OS").ToString()
                Cooler_Bank.AIR_PLUG_OS.Text = reader("AIR_PLUG_OS").ToString()
                Cooler_Bank.TM_POC_VAL_OS.Text = reader("TM_POC_VAL_OS").ToString()
                Cooler_Bank.FCD_VAL_OS.Text = reader("FCD_VAL_OS").ToString()
                Cooler_Bank.CP_LEN_OS.Text = reader("CP_LEN_OS").ToString()
                Cooler_Bank.HP_LEN_OS.Text = reader("HP_LEN_OS").ToString()
            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_ONAN_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from ONAN_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.Cooler_Pipe_size_ONAN.SelectedItem = reader("Cooler_Pipe_size_ONAN").ToString()
                Cooler_Bank.typ_of_cooling_cb1.SelectedItem = reader("typ_of_cooling_cb1").ToString()
                'Cooler_Bank.ONAN_no_port.SelectedItem = Bottom_Tank.NO_OF_PORT_BOT.SelectedItem
                'Cooler_Bank.ONAN_no_port.SelectedItem = reader("ONAN_no_port").ToString()
                'Cooler_Bank.ONAN_butterfly_valve_size.SelectedItem = Cooler_Bank.header_pipe_size_bot.SelectedItem
                'Cooler_Bank.ONAN_butterfly_valve_size.SelectedItem = reader("ONAN_butterfly_valve_size").ToString()
                Cooler_Bank.ONAN_expansion_bellows.SelectedItem = reader("ONAN_expansion_bellows").ToString()
                Cooler_Bank.ONAN_expansion_bellows_size.SelectedItem = reader("ONAN_expansion_bellows_size").ToString()
                Cooler_Bank.ONAN_cooler_pipe_supp_qty.SelectedItem = reader("ONAN_cooler_pipe_supp_qty").ToString()

                Cooler_Bank.P1L_LV.Text = reader("P1L_LV").ToString()
                Cooler_Bank.L1_LV.Text = reader("L1_LV").ToString()
                Cooler_Bank.P2L_LV.Text = reader("P2L_LV").ToString()
                Cooler_Bank.P3L_LV.Text = reader("P3L_LV").ToString()
                Cooler_Bank.L2_LV.Text = reader("L2_LV").ToString()
                Cooler_Bank.P4L_LV.Text = reader("P4L_LV").ToString()
                Cooler_Bank.L4_LV.Text = reader("L4_LV").ToString()
                Cooler_Bank.P5L1_LV.Text = reader("P5L1_LV").ToString()
                Cooler_Bank.P1L_HV.Text = reader("P1L_HV").ToString()
                Cooler_Bank.L1_HV.Text = reader("L1_HV").ToString()
                Cooler_Bank.P2L_HV.Text = reader("P2L_HV").ToString()
                Cooler_Bank.P3L_HV.Text = reader("P3L_HV").ToString()
                Cooler_Bank.L2_HV.Text = reader("L2_HV").ToString()
                Cooler_Bank.P4L_HV.Text = reader("P4L_HV").ToString()
                Cooler_Bank.L4_HV.Text = reader("L4_HV").ToString()
                Cooler_Bank.P5L1_HV.Text = reader("P5L1_HV").ToString()

                Cooler_Bank.BL1_LV.Text = reader("BL1_LV").ToString()
                Cooler_Bank.BL1_HV.Text = reader("BL1_HV").ToString()

                Cooler_Bank.X1LV.Text = reader("X1LV").ToString()
                Cooler_Bank.X2LV.Text = reader("X2LV").ToString()
                Cooler_Bank.X3LV.Text = reader("X3LV").ToString()
                Cooler_Bank.Y1LV.Text = reader("Y1LV").ToString()
                Cooler_Bank.Y2LV.Text = reader("Y2LV").ToString()
                Cooler_Bank.X1HV.Text = reader("X1HV").ToString()
                Cooler_Bank.X2HV.Text = reader("X2HV").ToString()
                Cooler_Bank.X3HV.Text = reader("X3HV").ToString()
                Cooler_Bank.Y1HV.Text = reader("Y1HV").ToString()
                Cooler_Bank.Y2HV.Text = reader("Y2HV").ToString()


                Cooler_Bank.ID1.Text = reader("ID1").ToString()
                Cooler_Bank.ID2.Text = reader("ID2").ToString()
                Cooler_Bank.BC_H1.Text = reader("BC_H1").ToString()
                Cooler_Bank.BC_H2.Text = reader("BC_H2").ToString()
                Cooler_Bank.BC_H3.Text = reader("BC_H3").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_ONAF_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from ONAF_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()
                Cooler_Bank.bot_Pipe_size_ONAF.SelectedItem = reader("bot_Pipe_size_ONAF").ToString()
                Cooler_Bank.TYP_COOLING_ONAF.SelectedItem = reader("TYP_COOLING_ONAF").ToString()
                'Cooler_Bank.NO_OF_PORT_ONAF.SelectedItem = reader("NO_OF_PORT_ONAF").ToString()
                'Cooler_Bank.NO_OF_PORT_ONAF.SelectedItem = Bottom_Tank.NO_OF_PORT_BOT.SelectedItem
                'Cooler_Bank.BF_VL_ONAF.SelectedItem = Cooler_Bank.header_pipe_size_bot.SelectedItem
                'Cooler_Bank.BF_VL_ONAF.SelectedItem = reader("BF_VL_ONAF").ToString()
                Cooler_Bank.EXP_BELOW_ONAF.SelectedItem = reader("EXP_BELOW_ONAF").ToString()
                Cooler_Bank.EXP_BLW_SIZE_ONAF.SelectedItem = reader("EXP_BLW_SIZE_ONAF").ToString()
                Cooler_Bank.IFRAME_ONAF.SelectedItem = reader("IFRAME_ONAF").ToString()

                Cooler_Bank.P1L_LVF.Text = reader("P1L_LVF").ToString()
                Cooler_Bank.L1_LVF.Text = reader("L1_LVF").ToString()
                Cooler_Bank.P2L_LVF.Text = reader("P2L_LVF").ToString()
                Cooler_Bank.P3L_LVF.Text = reader("P3L_LVF").ToString()
                Cooler_Bank.L2_LVF.Text = reader("L2_LVF").ToString()
                Cooler_Bank.P4L_LVF.Text = reader("P4L_LVF").ToString()
                Cooler_Bank.L4_LVF.Text = reader("L4_LVF").ToString()
                Cooler_Bank.P5L1_LVF.Text = reader("P5L1_LVF").ToString()
                Cooler_Bank.P1L_HVF.Text = reader("P1L_HVF").ToString()
                Cooler_Bank.L1_HVF.Text = reader("L1_HVF").ToString()
                Cooler_Bank.P2L_HVF.Text = reader("P2L_HVF").ToString()
                Cooler_Bank.P3L_HVF.Text = reader("P3L_HVF").ToString()
                Cooler_Bank.L2_HVF.Text = reader("L2_HVF").ToString()
                Cooler_Bank.P4L_HVF.Text = reader("P4L_HVF").ToString()
                Cooler_Bank.L4_HVF.Text = reader("L4_HVF").ToString()
                Cooler_Bank.P5L1_HVF.Text = reader("P5L1_HVF").ToString()

                Cooler_Bank.BL1_LVF.Text = reader("BL1_LVF").ToString()
                Cooler_Bank.BL1_HVF.Text = reader("BL1_HVF").ToString()

                Cooler_Bank.X1LVF.Text = reader("X1LVF").ToString()
                Cooler_Bank.X2LVF.Text = reader("X2LVF").ToString()
                Cooler_Bank.X3LVF.Text = reader("X3LVF").ToString()
                Cooler_Bank.Y1LVF.Text = reader("Y1LVF").ToString()
                Cooler_Bank.Y2LVF.Text = reader("Y2LVF").ToString()
                Cooler_Bank.X1HVF.Text = reader("X1HVF").ToString()
                Cooler_Bank.X2HVF.Text = reader("X2HVF").ToString()
                Cooler_Bank.X3HVF.Text = reader("X3HVF").ToString()
                Cooler_Bank.Y1HVF.Text = reader("Y1HVF").ToString()
                Cooler_Bank.Y2HVF.Text = reader("Y2HVF").ToString()


                Cooler_Bank.ID1F.Text = reader("ID1F").ToString()
                Cooler_Bank.ID2F.Text = reader("ID2F").ToString()
                Cooler_Bank.BC_H1F.Text = reader("BC_H1F").ToString()
                Cooler_Bank.BC_H2F.Text = reader("BC_H2F").ToString()
                Cooler_Bank.BC_H3F.Text = reader("BC_H3F").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_TopCoolerPipe_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from TopCoolerPipe_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Cooler_Bank.ODAF_cooler_pipe_size.SelectedItem = reader("ODAF_cooler_pipe_size").ToString()
                'Cooler_Bank.ODAF_no_of_port.SelectedItem = reader("ODAF_no_of_port").ToString()
                'Cooler_Bank.ODAF_no_of_port.SelectedItem = Top_Tank.NO_OF_PORT_TOP_HV.SelectedItem
                'Cooler_Bank.ODAF_no_of_port1.SelectedItem = Top_Tank.NO_OF_PORT_TOP_LV.SelectedItem
                Cooler_Bank.ODAF_butterfly_valve_size.SelectedItem = Cooler_Bank.header_pipe_size_bot.SelectedItem
                'Cooler_Bank.ODAF_butterfly_valve_size.SelectedItem = reader("ODAF_butterfly_valve_size").ToString()
                Cooler_Bank.ODAF_expansion_bellows.SelectedItem = reader("ODAF_expansion_bellows").ToString()
                Cooler_Bank.ODAF_expansion_bellows_size.SelectedItem = reader("ODAF_expansion_bellows_size").ToString()


                Cooler_Bank.P1L_LVT.Text = reader("P1L_LVT").ToString()
                Cooler_Bank.P2L_LVT.Text = reader("P2L_LVT").ToString()
                Cooler_Bank.P3L_LVT.Text = reader("P3L_LVT").ToString()
                Cooler_Bank.P4L_LVT.Text = reader("P4L_LVT").ToString()
                Cooler_Bank.P5L_LVT.Text = reader("P5L_LVT").ToString()
                Cooler_Bank.P1L_HVT.Text = reader("P1L_HVT").ToString()
                Cooler_Bank.P2L_HVT.Text = reader("P2L_HVT").ToString()
                Cooler_Bank.P3L_HVT.Text = reader("P3L_HVT").ToString()
                Cooler_Bank.P4L_HVT.Text = reader("P4L_HVT").ToString()
                Cooler_Bank.P5L_HVT.Text = reader("P5L_HVT").ToString()

                Cooler_Bank.X1LVT.Text = reader("X1LVT").ToString()
                Cooler_Bank.X2LVT.Text = reader("X2LVT").ToString()
                Cooler_Bank.X3LVT.Text = reader("X3LVT").ToString()
                Cooler_Bank.Y1LVT.Text = reader("Y1LVT").ToString()
                Cooler_Bank.X1HVT.Text = reader("X1HVT").ToString()
                Cooler_Bank.X2HVT.Text = reader("X2HVT").ToString()
                Cooler_Bank.X3HVT.Text = reader("X3HVT").ToString()
                Cooler_Bank.Y1HVT.Text = reader("Y1HVT").ToString()

                Cooler_Bank.L1_FVT.Text = reader("L1_FVT").ToString()
                Cooler_Bank.H1_FVT.Text = reader("H1_FVT").ToString()
                Cooler_Bank.H2_FVT.Text = reader("H2_FVT").ToString()
                Cooler_Bank.HL1_FVT.Text = reader("HL1_FVT").ToString()

                Cooler_Bank.P6L_LVT.Text = reader("P6L_LVT").ToString()
                Cooler_Bank.P6L_HVT.Text = reader("P6L_HVT").ToString()
                Cooler_Bank.Y2LVT.Text = reader("Y2LVT").ToString()
                Cooler_Bank.Y2HVT.Text = reader("Y2HVT").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_PipeArrangement200()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from PipeArrangement200 WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Cooler_Bank.D1A.Text = reader("D1A").ToString()
                Cooler_Bank.D2A.Text = reader("D2A").ToString()
                Cooler_Bank.D3A.Text = reader("D3A").ToString()
                Cooler_Bank.D4A.Text = reader("D4A").ToString()
                Cooler_Bank.D5A.Text = reader("D5A").ToString()
                Cooler_Bank.D6A.Text = reader("D6A").ToString()
                Cooler_Bank.D7A.Text = reader("D7A").ToString()
                Cooler_Bank.D8A.Text = reader("D8A").ToString()
                Cooler_Bank.APY1.Text = reader("APY1").ToString()

                Cooler_Bank.H12X100.Text = reader("H12X100").ToString()
                Cooler_Bank.U12X100.Text = reader("U12X100").ToString()
                Cooler_Bank.U22X100.Text = reader("U22X100").ToString()
                Cooler_Bank.U32X100.Text = reader("U32X100").ToString()

                Cooler_Bank.Coolingfan_POS.SelectedItem = reader("Coolingfan_POS").ToString()
                Cooler_Bank.FanLeftQTY.SelectedItem = reader("FanLeftQTY").ToString()
                Cooler_Bank.FanRightQTY.SelectedItem = reader("FanRightQTY").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_FrameProposal_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from FrameProposal_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Cooler_Bank.ISMC_W.Text = reader("ISMC_W").ToString()
                Cooler_Bank.ISMC_H.Text = reader("ISMC_H").ToString()
                Cooler_Bank.ISMC_T.Text = reader("ISMC_T").ToString()
                Cooler_Bank.ISMC_L.Text = reader("ISMC_L").ToString()

                Cooler_Bank.H1.Text = reader("H1").ToString()
                Cooler_Bank.H2.Text = reader("H2").ToString()
                Cooler_Bank.H3.Text = reader("H3").ToString()
                Cooler_Bank.H4.Text = reader("H4").ToString()
                Cooler_Bank.H5.Text = reader("H5").ToString()
                Cooler_Bank.H6.Text = reader("H6").ToString()
                Cooler_Bank.L1.Text = reader("L1").ToString()
                Cooler_Bank.ANG.Text = reader("ANG").ToString()
                Cooler_Bank.A1.Text = reader("A1").ToString()
                Cooler_Bank.A.Text = reader("A").ToString()
                Cooler_Bank.D.Text = reader("D").ToString()
                Cooler_Bank.DG.Text = reader("DG").ToString()
                Cooler_Bank.DL.Text = reader("DL").ToString()
                Cooler_Bank.A_D1.Text = reader("A_D1").ToString()
                Cooler_Bank.A_D2.Text = reader("A_D2").ToString()


            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Public Sub Load_PipeArrangement100()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from PipeArrangement100 WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Cooler_Bank.X100.Text = reader("X100").ToString()
                Cooler_Bank.H1100.Text = reader("H1100").ToString()
                Cooler_Bank.U1100.Text = reader("U1100").ToString()
                Cooler_Bank.U2100.Text = reader("U2100").ToString()
                Cooler_Bank.UT100.Text = reader("UT100").ToString()
                Cooler_Bank.D1100.Text = reader("D1100").ToString()
                Cooler_Bank.D2100.Text = reader("D2100").ToString()

                Cooler_Bank.Coolingfan_POS1.SelectedItem = reader("Coolingfan_POS1").ToString()
                Cooler_Bank.FanLeftQTY1.SelectedItem = reader("FanLeftQTY1").ToString()
                Cooler_Bank.FanRightQTY1.SelectedItem = reader("FanRightQTY1").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub


    Public Sub Load_Conservator_CB()

        Try

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to edit.")
                Exit Sub
            End If

            ' Get selected Work Order No
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            'Dim connectionString As String = "Server=HP-GET-07\SQLEXPRESS;Database=Indo_tech;User Id=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Conservator_CB WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()

                Cooler_Bank.Conservator_type2.SelectedItem = reader("Conservator_type2").ToString()
                Cooler_Bank.Conservator_type3.SelectedItem = reader("Conservator_type3").ToString()

                Cooler_Bank.Conservator_X_Distance.Text = reader("Conservator_X_Distance").ToString()

                Cooler_Bank.Conservator_dia.Text = reader("Conservator_dia").ToString()
                Cooler_Bank.Conservator_len.Text = reader("Conservator_len").ToString()

                Cooler_Bank.Conservator_Pos1.SelectedItem = reader("Conservator_Pos1").ToString()
                Cooler_Bank.Conservator_Pos2.SelectedItem = reader("Conservator_Pos2").ToString()
                Cooler_Bank.Conservator_Bracket.SelectedItem = reader("Conservator_Bracket").ToString()

                Cooler_Bank.Conservator_supp_dis_B1.Text = reader("Conservator_supp_dis_B1").ToString()
                Cooler_Bank.Conservator_supp_dis_B2.Text = reader("Conservator_supp_dis_B2").ToString()
                Cooler_Bank.B_R_pipe_center.Text = reader("B_R_pipe_center").ToString()

                Cooler_Bank.Conservator_drain_valve.SelectedItem = reader("Conservator_drain_valve").ToString()

                Cooler_Bank.Conservator_mounting.Text = reader("Conservator_mounting").ToString()

                Cooler_Bank.Aircel_rapture_relay.SelectedItem = reader("Aircel_rapture_relay").ToString()
                Cooler_Bank.MOLG_SIZE.SelectedItem = reader("MOLG_SIZE").ToString()
                Cooler_Bank.IITL_Monogram_plate.SelectedItem = reader("IITL_Monogram_plate").ToString()
                Cooler_Bank.B_R_Pipe_Connection.SelectedItem = reader("B_R_Pipe_Connection").ToString()
                Cooler_Bank.Breather_pipe_connection.SelectedItem = reader("Breather_pipe_connection").ToString()
                Cooler_Bank.Air_release_pipe_connection.SelectedItem = reader("Air_release_pipe_connection").ToString()
                Cooler_Bank.OLTC_Conservator_type1.SelectedItem = reader("OLTC_Conservator_type1").ToString()
                Cooler_Bank.OLTC_Conservator_type2.SelectedItem = reader("OLTC_Conservator_type2").ToString()

                Cooler_Bank.OLTC_conservator_dia.Text = reader("OLTC_conservator_dia").ToString()
                Cooler_Bank.OLTC_conservator_pos.Text = reader("OLTC_conservator_pos").ToString()

                Cooler_Bank.OLTC_cons_position.SelectedItem = reader("OLTC_cons_position").ToString()
                Cooler_Bank.OLTC_Conservator_supp.SelectedItem = reader("OLTC_Conservator_supp").ToString()

                Cooler_Bank.OLTC_Conservator_supp_dist_X.Text = reader("OLTC_Conservator_supp_dist_X").ToString()
                Cooler_Bank.OLTC_Conservator_supp_dist_Y.Text = reader("OLTC_Conservator_supp_dist_Y").ToString()
                Cooler_Bank.OSR_pipe_ceneter.Text = reader("OSR_pipe_ceneter").ToString()

                Cooler_Bank.OLTC_Conservator_drain_valve.SelectedItem = reader("OLTC_Conservator_drain_valve").ToString()
                Cooler_Bank.OLTC_MOLG_size.SelectedItem = reader("OLTC_MOLG_size").ToString()
                Cooler_Bank.OLTC_Breather_pipe_connection.SelectedItem = reader("OLTC_Breather_pipe_connection").ToString()
                Cooler_Bank.OLTC_suction_pipe_connection.SelectedItem = reader("OLTC_suction_pipe_connection").ToString()
                Cooler_Bank.OLTC_equalizer_pipe_connection.SelectedItem = reader("OLTC_equalizer_pipe_connection").ToString()
                Cooler_Bank.conservator_breather_cmb.SelectedItem = reader("conservator_breather_cmb").ToString()
                Cooler_Bank.conservator_breather_type.SelectedItem = reader("conservator_breather_type").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Create_Project.Show()
        Me.Close()
        Main_Panel.Hide()

        wk_odr_No = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

        Load_CreateProject()
        Load_BasicInputs()
        Load_BottomTank_BT()
        Load_UnderBase_BT()
        Load_FlatStiffner_BT()
        Load_JackingPad_HV_BT()
        Load_JakingPad_LV_BT()
        Load_Accessories_BT()
        Load_Tank_TT()
        Load_HVTermination_TT()
        Load_HVNTermination_TT()
        Load_LVTermination_TT()
        Load_BareBushing_TT()
        Load_CableBox_TT()
        Load_LV2CableBox_TT()
        Load_CB_Term_TT()
        Load_Trunk_DC_TT()
        Load_LVNTermination_TT()
        Load_Stiffner1_TT()
        Load_OLTC_Stiffner1_TT()
        Load_Lashing_leg_TT()
        Load_GCD_PARTS_TT()
        Load_Stiffner2_TT()
        Load_TopCover_Stiffner_TT()
        Load_Accessories1_TT()
        Load_Accessories2_TT()
        Load_Accessories3_TT()
        Accessories4_TT()
        Accessories5_TT()
        Load_TVTermination_TT()
        Load_CoolerBank_CB()
        Load_TopCooler1_CB()
        Load_TopCooler2_TT()
        Load_TopCooler3_CB()
        Load_TopBotHeader_CB()
        Load_TopBotHeader_CB()
        Load_ONAN_CB()
        Load_ONAF_CB()
        Load_TopCoolerPipe_CB()
        Load_PipeArrangement200()
        Load_FrameProposal_CB()
        Load_PipeArrangement100()
        Load_Conservator_CB()

    End Sub

    'Public Sub DeleteWorkOrderData()

    '    If DataGridView1.SelectedRows.Count = 0 Then
    '        MessageBox.Show("Please select a row to delete.")
    '        Exit Sub
    '    End If

    '    ' Get selected Work Order No from DataGridView
    '    Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()


    '    ' Your DB connection
    '    Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"

    '    Using conn As New SqlConnection(connectionString)
    '        conn.Open()

    '        Dim transaction = conn.BeginTransaction()

    '        Try
    '            ' List of all table names to delete from
    '            Dim tables As String() = {
    '            "Create_Project",
    '            "Basic_Input",
    '            "Bottom_Tank_BT",
    '            "UnderBase_BT",
    '            "CurbDetails_BT",
    '            "FlatStiffner_BT",
    '            "JackingPad_HV_BT",
    '            "JakingPad_LV_BT",
    '            "Accessories_BT",
    '            "Tank_TT",
    '            "HVTermination_TT",
    '            "HVNTermination_TT",
    '            "LVTermination_TT",
    '            "BareBushing_TT",
    '            "CableBox_TT",
    '            "LV2CableBox_TT",
    '            "LVNTermination_TT",
    '            "TVTermination_TT",
    '            "Stiffner1_TT",
    '            "OLTC_Stiffner1_TT",
    '            "Stiffner2_TT",
    '            "TopCover_Stiffner_TT",
    '            "Accessories1_TT",
    '            "Accessories2_TT",
    '            "Accessories3_TT",
    '            "Accessories4_TT",
    '            "CoolerBank_CB",
    '            "TopCooler1_CB",
    '            "TopCooler2_TT",
    '            "TopCooler3_CB",
    '            "TopBotHeader_CB",
    '            "ONAN_CB",
    '            "ONAF_CB",
    '            "TopCoolerPipe_CB",
    '            "PipeArrangement200",
    '            "FrameProposal_CB",
    '            "PipeArrangement100",
    '            "Conservator_CB"
    '                        }

    '            ' Execute delete for each table
    '            For Each tableName As String In tables
    '                Dim deleteQuery As String = $"DELETE FROM {tableName} WHERE wk_odr_No = @wk_odr_No"
    '                Using cmd As New SqlCommand(deleteQuery, conn, transaction)
    '                    cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
    '                    cmd.ExecuteNonQuery()
    '                End Using
    '            Next

    '            ' Commit all deletes
    '            transaction.Commit()


    '        Catch ex As Exception
    '            transaction.Rollback()
    '            MessageBox.Show("Error while deleting: " & ex.Message)
    '        End Try

    '        conn.Close()
    '    End Using

    'End Sub


    Public Sub DeleteWorkOrderData()

        Dim result As DialogResult = MessageBox.Show(
                "Are you sure want to delete?",
                "Confirm Update",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

        If result = DialogResult.Yes Then

            If DataGridView1.SelectedRows.Count = 0 Then
                MessageBox.Show("Please select a row to delete.")
                Exit Sub
            End If

            ' Get selected Work Order No from DataGridView
            Dim wk_odr_No As String = DataGridView1.SelectedRows(0).Cells("Column2").Value.ToString()

            ' Prevent deletion of Work Order No "0000"
            If wk_odr_No = "0000" Then
                MessageBox.Show("Work Order No '0000' cannot be deleted.", "IndoTech Transformer")
                Exit Sub
            End If

            ' Your DB connection
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"

            Using conn As New SqlConnection(Main_Panel.connectionString)
                conn.Open()
                Dim transaction = conn.BeginTransaction()

                Try
                    ' List of all table names to delete from
                    Dim tables As String() = {
                    "Create_Project", "Basic_Input", "Bottom_Tank_BT", "UnderBase_BT", "CurbDetails_BT",
                    "FlatStiffner_BT", "JackingPad_HV_BT", "JakingPad_LV_BT", "Accessories_BT",
                    "Tank_TT", "HVTermination_TT", "HVNTermination_TT", "LVTermination_TT",
                    "BareBushing_TT", "CableBox_TT", "LV2CableBox_TT", "LVNTermination_TT",
                    "TVTermination_TT", "Stiffner1_TT", "OLTC_Stiffner1_TT", "Stiffner2_TT",
                    "TopCover_Stiffner_TT", "Accessories1_TT", "Accessories2_TT", "Accessories3_TT", "Accessories4_TT",
                    "CoolerBank_CB", "TopCooler1_CB", "TopCooler2_TT", "TopCooler3_CB",
                    "TopBotHeader_CB", "ONAN_CB", "ONAF_CB", "TopCoolerPipe_CB",
                    "PipeArrangement200", "FrameProposal_CB", "PipeArrangement100", "Conservator_CB", "Lashing_Leg_TT", "GCD_PARTS",
                    "Accesories5_TT", "CB_Term_TT", "Trunk_DC_TT"
                }

                    ' Execute delete for each table
                    For Each tableName As String In tables
                        Dim deleteQuery As String = $"DELETE FROM {tableName} WHERE wk_odr_No = @wk_odr_No"
                        Using cmd As New SqlCommand(deleteQuery, conn, transaction)
                            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No)
                            cmd.ExecuteNonQuery()
                        End Using
                    Next

                    ' Commit all deletes
                    transaction.Commit()
                    MessageBox.Show($"Work Order No '{wk_odr_No}' deleted successfully.")

                Catch ex As Exception
                    transaction.Rollback()
                    MessageBox.Show("Error while deleting: " & ex.Message)
                End Try

                conn.Close()

            End Using
        Else
            MsgBox("Failed to Delete!!!")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DeleteWorkOrderData()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.ReadOnly = True
        load_data()
    End Sub

    Private Sub SearchWorkOrder_Prefix()
        Try
            Dim searchText As String = TextBox1.Text.Trim()

            If String.IsNullOrEmpty(searchText) Then
                MessageBox.Show("Please enter a Work Order No to search.")
                Exit Sub
            End If


            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"

            Using connection As New SqlConnection(Main_Panel.connectionString)
                connection.Open()

                Dim query As String = "SELECT * FROM Create_Project WHERE LTRIM(RTRIM(wk_odr_No)) LIKE @prefix"

                Using adapter As New SqlDataAdapter(query, connection)
                    adapter.SelectCommand.Parameters.AddWithValue("@prefix", searchText & "%")

                    Dim table As New DataTable()
                    adapter.Fill(table)

                    If table.Rows.Count = 0 Then
                        MessageBox.Show("No matching records found.")
                    Else

                        Dim resultMessage As String = "Results Found:" & vbCrLf & vbCrLf

                        For Each row As DataRow In table.Rows
                            resultMessage &= "Work Order No: " & row("wk_odr_No").ToString() & vbCrLf

                        Next


                        DataGridView1.AutoGenerateColumns = True
                        DataGridView1.DataSource = Nothing
                        DataGridView1.Rows.Clear()
                        DataGridView1.Columns.Clear()
                        DataGridView1.DataSource = table


                        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                    End If
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Error during search: " & ex.Message)
        End Try
    End Sub


    Private Sub Seacrh_Click(sender As Object, e As EventArgs) Handles Seacrh.Click
        SearchWorkOrder_Prefix()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            SearchWorkOrder_Prefix()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            ' Step 1: Get username and work order no from selected row
            Dim username As String = DataGridView1.CurrentRow.Cells("Column5").Value.ToString()
            Dim workOrderNo As String = DataGridView1.CurrentRow.Cells("Column2").Value.ToString()

            ' Step 2: Build path
            Dim rootPath As String = "D:\Cadopt_Tool_Output"
            Dim userFolder As String = System.IO.Path.Combine(rootPath, username)
            Dim workOrderFolder As String = System.IO.Path.Combine(userFolder, workOrderNo)

            ' ✅ Updated file name: [WorkOrderNo]_OGA.iam
            Dim modelFileName As String = workOrderNo & "_OGA.iam"
            Dim modelPath As String = System.IO.Path.Combine(workOrderFolder, modelFileName)

            ' Step 3: Check if file exists
            If System.IO.File.Exists(modelPath) Then
                ' Step 4: Launch Inventor and open model
                Dim inventorApp As Inventor.Application

                Try
                    inventorApp = Marshal.GetActiveObject("Inventor.Application")
                Catch ex As Exception
                    ' If not running, start it
                    Dim invType As Type = Type.GetTypeFromProgID("Inventor.Application")
                    inventorApp = CType(Activator.CreateInstance(invType), Inventor.Application)
                    inventorApp.Visible = True
                End Try

                inventorApp.Documents.Open(modelPath)
                MessageBox.Show("Model opened successfully.")
            Else
                MessageBox.Show("Model file not found: " & modelPath)
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try

    End Sub
End Class
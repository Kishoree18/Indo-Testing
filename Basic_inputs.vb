Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Inventor
Imports file = Inventor.File

Public Class Basic_inputs

    Public Shared ThisApplication As Inventor.Application
    Public Shared transform_ft_view_val As String


    'Public Property CurrentWorkOrderNo As String = "0000"  ' Default to "0000"



    Private Sub core_limp_SelectedIndexChanged(sender As Object, e As EventArgs) Handles core_limp.SelectedIndexChanged

        MS_T.Hide()
        WH.Hide()
        MAXOD.Hide()
        MS_B.Hide()
        ST.Hide()
        LC1_L.Hide()
        LC1R.Hide()
        FV.Hide()


        LC2_5L1.Hide()
        LC2_5L2.Hide()
        LC1_5L1.Hide()
        LC1_5L2.Hide()
        WH_5L.Hide()
        MS_5L.Hide()
        ELH1_5L.Hide()
        ELH2_5L.Hide()
        ST_5L.Hide()
        MAXOD_5L.Hide()

        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources._3_limb



            end_limb_1.Enabled = False
            Tp_btm_el_1.Enabled = False
            leg_centre_2.Enabled = False
        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then




            Basic_input_picbox.Image = My.Resources._5_limb
            end_limb_1.Enabled = True
            Tp_btm_el_1.Enabled = True
            leg_centre_2.Enabled = True
        End If


    End Sub

    Public Sub LoadData_BI()


        Try
            If Main_Panel.userTxt.Text = "" Or Create_Project.wk_ord_No = "" Then
                MsgBox("Enter username and work order number to proceed")
            Else
                'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
                'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
                Using connection As New SqlConnection(Main_Panel.connectionString)
                    connection.Open()

                    ' Step 1: Check if work order number exists
                    Dim checkQuery As String = "SELECT COUNT(*) FROM Basic_Input WHERE wk_odr_No = @wk_odr_No"
                    Using checkCmd As New SqlCommand(checkQuery, connection)
                        checkCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

                        Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                        If count > 0 Then
                            ' 🔁 Update existing row
                            Dim updateQuery As String = "
                        UPDATE Basic_Input SET 
                            Username=@Username,
                            core_limp=@core_limp,
                            core_dia=@core_dia,
                            stack_thk=@stack_thk,
                            max_step=@max_step,
                            leg_centre_1=@leg_centre_1,
                            leg_centre_2=@leg_centre_2,
                            Max_winding_od=@Max_winding_od,
                            window_ht=@window_ht,
                            end_limb_1=@end_limb_1,
                            Tp_btm_el_1=@Tp_btm_el_1,
                            vector_group=@vector_group,
                            type_of_cooling=@type_of_cooling,
                            transform_ft_view=@transform_ft_view,
                            OLTC_pkt_position=@OLTC_pkt_position
                            WHERE wk_odr_No=@wk_odr_No"

                            Using updateCmd As New SqlCommand(updateQuery, connection)
                                updateCmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                                updateCmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)
                                updateCmd.Parameters.AddWithValue("@core_limp", core_limp.SelectedItem)
                                updateCmd.Parameters.AddWithValue("@core_dia", core_dia.Text)
                                updateCmd.Parameters.AddWithValue("@stack_thk", stack_thk.Text)
                                updateCmd.Parameters.AddWithValue("@max_step", max_step.Text)
                                updateCmd.Parameters.AddWithValue("@leg_centre_1", leg_centre_1.Text)
                                updateCmd.Parameters.AddWithValue("@leg_centre_2", leg_centre_2.Text)
                                updateCmd.Parameters.AddWithValue("@Max_winding_od", Max_winding_od.Text)
                                updateCmd.Parameters.AddWithValue("@window_ht", window_ht.Text)
                                updateCmd.Parameters.AddWithValue("@end_limb_1", end_limb_1.Text)
                                updateCmd.Parameters.AddWithValue("@Tp_btm_el_1", Tp_btm_el_1.Text)
                                updateCmd.Parameters.AddWithValue("@vector_group", vector_group.Text)
                                updateCmd.Parameters.AddWithValue("@type_of_cooling", type_of_cooling.SelectedItem)
                                updateCmd.Parameters.AddWithValue("@transform_ft_view", transform_ft_view.SelectedItem)
                                updateCmd.Parameters.AddWithValue("@OLTC_pkt_position", OLTC_pkt_position.SelectedItem)

                                updateCmd.ExecuteNonQuery()
                                MsgBox("Data updated successfully!")
                            End Using
                        Else
                            '    ' ➕ Insert new row
                            '    Dim insertQuery As String = "
                            'INSERT INTO Basic_Input 
                            '    (Username, wk_odr_No, core_limp, core_dia, stack_thk, max_step, leg_centre_1, leg_centre_2, Max_winding_od,
                            '     window_ht, end_limb_1, Tp_btm_el_1, vector_group, type_of_cooling, transform_ft_view, OLTC_pkt_position, Gasket_type, Flange_type)
                            'VALUES 
                            '    (@Username, @wk_odr_No, @core_limp, @core_dia, @stack_thk, @max_step, @leg_centre_1, @leg_centre_2, @Max_winding_od,
                            '     @window_ht, @end_limb_1, @Tp_btm_el_1, @vector_group, @type_of_cooling, @transform_ft_view, @OLTC_pkt_position, @Gasket_type, @Flange_type)"

                            '    Using insertCmd As New SqlCommand(insertQuery, connection)
                            '        insertCmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
                            '        insertCmd.Parameters.AddWithValue("@wk_odr_No", Edit_Project.wk_odr_No)
                            '        insertCmd.Parameters.AddWithValue("@core_limp", core_limp.SelectedItem)
                            '        insertCmd.Parameters.AddWithValue("@core_dia", core_dia.Text)
                            '        insertCmd.Parameters.AddWithValue("@stack_thk", stack_thk.Text)
                            '        insertCmd.Parameters.AddWithValue("@max_step", max_step.Text)
                            '        insertCmd.Parameters.AddWithValue("@leg_centre_1", leg_centre_1.Text)
                            '        insertCmd.Parameters.AddWithValue("@leg_centre_2", leg_centre_2.Text)
                            '        insertCmd.Parameters.AddWithValue("@Max_winding_od", Max_winding_od.Text)
                            '        insertCmd.Parameters.AddWithValue("@window_ht", window_ht.Text)
                            '        insertCmd.Parameters.AddWithValue("@end_limb_1", end_limb_1.Text)
                            '        insertCmd.Parameters.AddWithValue("@Tp_btm_el_1", Tp_btm_el_1.Text)
                            '        insertCmd.Parameters.AddWithValue("@vector_group", vector_group.Text)
                            '        insertCmd.Parameters.AddWithValue("@type_of_cooling", type_of_cooling.SelectedItem)
                            '        insertCmd.Parameters.AddWithValue("@transform_ft_view", transform_ft_view.SelectedItem)
                            '        insertCmd.Parameters.AddWithValue("@OLTC_pkt_position", OLTC_pkt_position.SelectedItem)
                            '        insertCmd.Parameters.AddWithValue("@Gasket_type", Gasket_type.SelectedItem)
                            '        insertCmd.Parameters.AddWithValue("@Flange_type", Flange_type.SelectedItem)

                            '        insertCmd.ExecuteNonQuery()
                            '        MsgBox("Data inserted successfully!")
                            '    End Using
                        End If
                    End Using
                End Using
            End If
        Catch ex As Exception
            MsgBox("Unable to insert/update record: " & ex.Message)
        End Try


        'Try

        '    If Main_Panel.userTxt.Text = Nothing Or Create_Project.wk_odr_No.Text = Nothing Then
        '        MsgBox("Enter username and password to proceed")
        '    Else

        '        'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '        Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '        Dim connection As New SqlConnection(connectionString)

        '        connection.Open()
        '        Dim query As String = "INSERT INTO Basic_Input (Username,wk_odr_No,core_limp,core_dia,stack_thk,max_step,leg_centre_1,leg_centre_2,Max_winding_od,window_ht,end_limb_1,Tp_btm_el_1,vector_group,type_of_cooling,transform_ft_view,OLTC_pkt_position,Gasket_type,Flange_type) 
        '                           VALUES (@Username,@wk_odr_No,@core_limp,@core_dia,@stack_thk,@max_step,@leg_centre_1,@leg_centre_2,@Max_winding_od,@window_ht,@end_limb_1,@Tp_btm_el_1,@vector_group,@type_of_cooling,@transform_ft_view,@OLTC_pkt_position,@Gasket_type,@Flange_type)"

        '        Dim cmd As New SqlCommand(query, connection)
        '        cmd.Parameters.AddWithValue("@Username", Main_Panel.userTxt.Text)
        '        cmd.Parameters.AddWithValue("@wk_odr_No", Create_Project.wk_ord_No)

        '        cmd.Parameters.AddWithValue("@core_limp", core_limp.SelectedItem)
        '        cmd.Parameters.AddWithValue("@core_dia", core_dia.Text)
        '        cmd.Parameters.AddWithValue("@stack_thk", stack_thk.Text)
        '        cmd.Parameters.AddWithValue("@max_step", max_step.Text)
        '        cmd.Parameters.AddWithValue("@leg_centre_1", leg_centre_1.Text)
        '        cmd.Parameters.AddWithValue("@leg_centre_2", leg_centre_2.Text)
        '        cmd.Parameters.AddWithValue("@Max_winding_od", Max_winding_od.Text)
        '        cmd.Parameters.AddWithValue("@window_ht", window_ht.Text)
        '        cmd.Parameters.AddWithValue("@end_limb_1", end_limb_1.Text)
        '        cmd.Parameters.AddWithValue("@Tp_btm_el_1", Tp_btm_el_1.Text)
        '        cmd.Parameters.AddWithValue("@vector_group", vector_group.Text)
        '        cmd.Parameters.AddWithValue("@type_of_cooling", type_of_cooling.SelectedItem)
        '        cmd.Parameters.AddWithValue("@transform_ft_view", transform_ft_view.SelectedItem)
        '        cmd.Parameters.AddWithValue("@OLTC_pkt_position", OLTC_pkt_position.SelectedItem)
        '        cmd.Parameters.AddWithValue("@Gasket_type", Gasket_type.SelectedItem)
        '        cmd.Parameters.AddWithValue("@Flange_type", Flange_type.SelectedItem)

        '        Dim result As Integer = cmd.ExecuteNonQuery()
        '        connection.Close()

        '        If result > 0 Then
        '            MsgBox("Data inserted successfully!!!")
        '        Else
        '            MsgBox("Data not inserted!!!")
        '        End If

        '    End If
        'Catch ex As Exception
        '    MsgBox("Unable to insert record" & ex.ToString)
        'End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'LoadData_CP()
        LoadData_BI()
    End Sub





    Private Sub transform_ft_view_SelectedIndexChanged(sender As Object, e As EventArgs) Handles transform_ft_view.SelectedIndexChanged
        If transform_ft_view.SelectedItem IsNot Nothing Then
            transform_ft_view_val = transform_ft_view.SelectedItem.ToString()

        End If

        If transform_ft_view.SelectedItem = "HV" Then

            Basic_input_picbox.Image = My.Resources.OLTC_RIGHT_FVLV

        Else

            Basic_input_picbox.Image = My.Resources.OLTC_RIGHT_FVLV


        End If

    End Sub

    Public Sub Load_Create_project_inputs()

        core_limp.SelectedItem = Main_Panel.core_limp
        core_dia.Text = Main_Panel.core_dia
        stack_thk.Text = Main_Panel.stack_thk
        max_step.Text = Main_Panel.max_step
        leg_centre_1.Text = Main_Panel.leg_centre_1
        leg_centre_2.Text = Main_Panel.leg_centre_2
        end_limb_1.Text = Main_Panel.end_limb_1
        Tp_btm_el_1.Text = Main_Panel.Tp_btm_el_1
        window_ht.Text = Main_Panel.window_ht
        Max_winding_od.Text = Main_Panel.Max_winding_od
        vector_group.Text = Main_Panel.vector_group

        type_of_cooling.SelectedItem = Main_Panel.type_of_cooling
        transform_ft_view.SelectedItem = Main_Panel.transform_ft_view

        OLTC_pkt_position.SelectedItem = Main_Panel.OLTC_pkt_position

    End Sub

    Public Sub Save_Basicinput_project_inputs()

        Main_Panel.core_limp = core_limp.SelectedItem
        Main_Panel.core_dia = core_dia.Text
        Main_Panel.stack_thk = stack_thk.Text
        Main_Panel.max_step = max_step.Text
        Main_Panel.leg_centre_1 = leg_centre_1.Text
        Main_Panel.leg_centre_2 = leg_centre_2.Text
        Main_Panel.end_limb_1 = end_limb_1.Text
        Main_Panel.Tp_btm_el_1 = Tp_btm_el_1.Text
        Main_Panel.window_ht = window_ht.Text
        Main_Panel.Max_winding_od = Max_winding_od.Text
        Main_Panel.vector_group = vector_group.Text

        Main_Panel.type_of_cooling = type_of_cooling.SelectedItem
        Main_Panel.transform_ft_view = transform_ft_view.SelectedItem
        ' Main_Panel.OLTC_pkt_position = OLTC_pkt_position.SelectedItem
        Main_Panel.OLTC_pkt_position = OLTC_pkt_position.SelectedItem


    End Sub

    'Public Sub Load_BasicInputs(wk_odr_No As String)

    '    Try
    '        MsgBox("wk_odr_No " & wk_odr_No)

    '        Dim wk_odr_No1 As String = "0000"


    '        Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
    '        Dim connection As New SqlConnection(connectionString)

    '        connection.Open()
    '        Dim query As String = "SELECT * from Basic_Input WHERE wk_odr_No = @wk_odr_No"

    '        Dim cmd As New SqlCommand(query, connection)
    '        cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No.Trim())
    '        Dim reader As SqlDataReader = cmd.ExecuteReader()

    '        While reader.Read()
    '            'Main_Panel.userTxt.Text = reader("Username").ToString()

    '            core_limp.SelectedItem = reader("core_limp").ToString()
    '            core_dia.Text = reader("core_dia").ToString()
    '            MsgBox(core_dia.Text)
    '            stack_thk.Text = reader("stack_thk").ToString()
    '            max_step.Text = reader("max_step").ToString()
    '            leg_centre_1.Text = reader("leg_centre_1").ToString()
    '            leg_centre_2.Text = reader("leg_centre_2").ToString()

    '            Max_winding_od.Text = reader("Max_winding_od").ToString()
    '            window_ht.Text = reader("window_ht").ToString()
    '            end_limb_1.Text = reader("end_limb_1").ToString()
    '            Tp_btm_el_1.Text = reader("Tp_btm_el_1").ToString()
    '            vector_group.Text = reader("vector_group").ToString()
    '            type_of_cooling.SelectedItem = reader("type_of_cooling").ToString()
    '            transform_ft_view.SelectedItem = reader("transform_ft_view").ToString()
    '            OLTC_pkt_position.SelectedItem = reader("OLTC_pkt_position").ToString()
    '            Gasket_type.SelectedItem = reader("Gasket_type").ToString()
    '            Flange_type.SelectedItem = reader("Flange_type").ToString()

    '        End While

    '        connection.Close()

    '    Catch ex As Exception
    '        MsgBox("Error" & ex.ToString)
    '    End Try

    'End Sub
    Public Sub Load_BasicInputs()

        Try


            Dim wk_odr_No1 As String = Create_Project.wk_ord_No


            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from Basic_Input WHERE wk_odr_No = @wk_odr_No"

            Dim cmd As New SqlCommand(query, connection)
            cmd.Parameters.AddWithValue("@wk_odr_No", wk_odr_No1)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            While reader.Read()
                'Main_Panel.userTxt.Text = reader("Username").ToString()

                core_limp.SelectedItem = reader("core_limp").ToString()
                core_dia.Text = reader("core_dia").ToString()
                stack_thk.Text = reader("stack_thk").ToString()
                max_step.Text = reader("max_step").ToString()
                leg_centre_1.Text = reader("leg_centre_1").ToString()
                leg_centre_2.Text = reader("leg_centre_2").ToString()

                Max_winding_od.Text = reader("Max_winding_od").ToString()
                window_ht.Text = reader("window_ht").ToString()
                end_limb_1.Text = reader("end_limb_1").ToString()
                Tp_btm_el_1.Text = reader("Tp_btm_el_1").ToString()
                vector_group.Text = reader("vector_group").ToString()
                type_of_cooling.SelectedItem = reader("type_of_cooling").ToString()
                transform_ft_view.SelectedItem = reader("transform_ft_view").ToString()
                OLTC_pkt_position.SelectedItem = reader("OLTC_pkt_position").ToString()

            End While

            connection.Close()

        Catch ex As Exception
            MsgBox("Error" & ex.ToString)
        End Try

    End Sub

    Private Sub Basic_inputs_Load(sender As Object, e As EventArgs) Handles MyBase.Load




        If Main_Panel.core_limp = Nothing Then
            core_limp.SelectedItem = "3 Limbs"
            type_of_cooling.SelectedItem = "ONAN"
            transform_ft_view.SelectedItem = "HV"
            OLTC_pkt_position.SelectedItem = "LV"
            'OLTC_pkt_position.SelectedItem = "Left"
            ' OLTC_pkt_position.SelectedItem = "HV"



        Else
            Load_Create_project_inputs()


        End If

        transform_ft_view.Enabled = False

        FV.Hide()
        Basic_input_picbox.Image = My.Resources._3_limb



        Load_BasicInputs()


    End Sub






    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Save_Basicinput_project_inputs()
        Create_Project.Show()
        Me.Close()


    End Sub

    Private Sub core_dia_MouseClick(sender As Object, e As MouseEventArgs) Handles core_dia.MouseClick


        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources._3limb_coredia
        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then
            Basic_input_picbox.Image = My.Resources._5limb_coredia
        End If

        MS_T.Hide()
        WH.Hide()
        MAXOD.Hide()
        MS_B.Hide()
        ST.Hide()
        LC1R.Hide()
        LC1_L.Hide()


        LC2_5L1.Hide()
        LC2_5L2.Hide()
        LC1_5L1.Hide()
        LC1_5L2.Hide()
        WH_5L.Hide()
        MS_5L.Hide()
        ELH1_5L.Hide()
        ELH2_5L.Hide()
        ST_5L.Hide()
        MAXOD_5L.Hide()
        FV.Hide()

    End Sub

    Private Sub stack_thk_MouseClick(sender As Object, e As MouseEventArgs) Handles stack_thk.MouseClick


        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.ST
            ST.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MAXOD.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_T.Show()
            WH.Show()
            MAXOD.Show()
            MS_B.Show()
            ST.Show()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()


            FV.Hide()
        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.ST_5L
            ST_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MAXOD.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            ST_5L.Show()
            MAXOD_5L.Show()

            FV.Hide()




        End If





    End Sub



    Private Sub Max_winding_od_MouseClick(sender As Object, e As MouseEventArgs) Handles Max_winding_od.MouseClick



        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.ST


            ' Basic_input_picbox.Image = My.Resources.ST
            ST.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MAXOD.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            MS_T.Show()
            WH.Show()
            MAXOD.Show()
            MS_B.Show()
            ST.Show()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.ST_5L



            ST_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MAXOD_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            ST_5L.Show()
            MAXOD_5L.Show()

            FV.Hide()





        End If



    End Sub

    Private Sub max_step_MouseClick(sender As Object, e As MouseEventArgs) Handles max_step.MouseClick




        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.ST


            ST.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MAXOD.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_B.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_T.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)

            MS_T.Show()
            WH.Show()
            MAXOD.Show()
            MS_B.Show()
            ST.Show()
            LC1R.Hide()
            LC1_L.Hide()


            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC2_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC2_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()
            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()

        End If




    End Sub


    Private Sub window_ht_MouseClick_1(sender As Object, e As MouseEventArgs) Handles window_ht.MouseClick




        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.ST


            ST.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MAXOD.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Show()
            WH.Show()
            MAXOD.Show()
            MS_B.Show()
            ST.Show()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC2_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC2_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()

            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()
        End If



    End Sub

    Private Sub leg_centre_1_MouseClick(sender As Object, e As MouseEventArgs) Handles leg_centre_1.MouseClick





        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.LC_3LIMB_1

            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Show()
            LC1_L.Show()

            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()

            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC2_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC2_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()

            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()

        End If




    End Sub

    Private Sub leg_centre_2_MouseClick(sender As Object, e As MouseEventArgs) Handles leg_centre_2.MouseClick


        If core_limp.SelectedItem.ToString = "3 Limbs" Then
            Basic_input_picbox.Image = My.Resources.LC_3LIMB_1



            LC1R.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC1_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Show()
            LC1_L.Show()
            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()
            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC2_5L1.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            LC2_5L2.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()

            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()
        End If




    End Sub

    Private Sub end_limb_1_MouseClick(sender As Object, e As MouseEventArgs) Handles end_limb_1.MouseClick


        If core_limp.SelectedItem.ToString = "3 Limbs" Then


            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Show()
            LC1_L.Show()
            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()
            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()

            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()
        End If




    End Sub

    Private Sub Tp_btm_el_1_MouseClick(sender As Object, e As MouseEventArgs) Handles Tp_btm_el_1.MouseClick

        If core_limp.SelectedItem.ToString = "3 Limbs" Then


            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Show()
            LC1_L.Show()
            LC2_5L2.Hide()
            LC2_5L2.Hide()
            LC1_5L1.Hide()
            LC1_5L2.Hide()
            WH_5L.Hide()
            MS_5L.Hide()
            ELH1_5L.Hide()
            ELH2_5L.Hide()
            FV.Hide()

        ElseIf core_limp.SelectedItem.ToString = "5 Limbs" Then


            Basic_input_picbox.Image = My.Resources.LC_5LIMBS


            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            ELH1_5L.BackColor = System.Drawing.Color.FromArgb(120, 250, 60)
            WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
            LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)

            MS_T.Hide()
            WH.Hide()
            MAXOD.Hide()
            MS_B.Hide()
            ST.Hide()
            LC1R.Hide()
            LC1_L.Hide()

            LC2_5L1.Show()
            LC2_5L2.Show()
            LC1_5L1.Show()
            LC1_5L2.Show()
            WH_5L.Show()
            MS_5L.Show()
            ELH1_5L.Show()
            ELH2_5L.Show()

            ST_5L.Hide()
            MAXOD_5L.Hide()
            FV.Hide()

        End If





    End Sub

    Private Sub OLTC_pkt_position_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OLTC_pkt_position.SelectedIndexChanged


        MS_T.Hide()
        WH.Hide()
        MAXOD.Hide()
        MS_B.Hide()
        ST.Hide()
        LC1R.Hide()
        LC1_L.Hide()

        LC2_5L1.Hide()
        LC2_5L2.Hide()
        LC1_5L1.Hide()
        LC1_5L2.Hide()
        WH_5L.Hide()
        MS_5L.Hide()
        ELH1_5L.Hide()
        ELH2_5L.Hide()

        ST_5L.Hide()
        MAXOD_5L.Hide()
        FV.Show()

        'If OLTC_pkt_position.SelectedItem = "Left" And OLTC_pkt_position.SelectedItem = "HV" Then

        '    Basic_input_picbox.Image = My.Resources.OLTC_LEFT_FVLV

        'ElseIf OLTC_pkt_position.SelectedItem = "Left" And OLTC_pkt_position.SelectedItem = "LV" Then

        '    Basic_input_picbox.Image = My.Resources.OLTC_LEFT_FVHV


        'ElseIf OLTC_pkt_position.SelectedItem = "Right" And OLTC_pkt_position.SelectedItem = "HV" Then

        '    Basic_input_picbox.Image = My.Resources.OLTC_RIGHT_FVLV

        'ElseIf OLTC_pkt_position.SelectedItem = "Right" And OLTC_pkt_position.SelectedItem = "LV" Then

        '    Basic_input_picbox.Image = My.Resources.OLTC_RIGHT_FVHV


        'End If


    End Sub



    Private Sub vector_group_Click(sender As Object, e As EventArgs) Handles vector_group.Click


        LC1_5L1.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        LC1_5L2.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        ELH2_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        ELH1_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        WH_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        MS_5L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)


        ST.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        WH.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        LC1R.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        LC1_L.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        MAXOD.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        MS_B.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
        MS_T.BackColor = System.Drawing.Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Basic_inputs_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Save_Basicinput_project_inputs()
        Create_Project.Show()
    End Sub

    Private Sub vector_group_MouseClick(sender As Object, e As MouseEventArgs) Handles vector_group.MouseClick

    End Sub

    Private Sub OLTC_pkt_position_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles OLTC_pkt_position.SelectedIndexChanged

    End Sub
End Class
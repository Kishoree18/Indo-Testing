Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Diagnostics.Eventing
Imports MySql.Data.MySqlClient
Imports System.IO

Public Class Create_New_User

    Public Shared userdata As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles create_btn.Click

        Try
            If userTxt.Text = Nothing Or passTxt.Text = Nothing Then
                MsgBox("Enter username and password to proceed")
            Else
                Using connection As New SqlConnection(Main_Panel.connectionString)
                    connection.Open()

                    ' Step 1: Check if same username/password/credential already exists
                    Dim checkQuery As String = "SELECT COUNT(*) FROM log WHERE Username = @username OR Password = @password"
                    Using checkCmd As New SqlCommand(checkQuery, connection)
                        checkCmd.Parameters.AddWithValue("@username", userTxt.Text)
                        checkCmd.Parameters.AddWithValue("@password", passTxt.Text)

                        Dim exists As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())

                        If exists > 0 Then
                            MsgBox("This Username with the same password and role already exists.")
                            Exit Sub
                        End If
                    End Using

                    ' Step 2: Insert new user record
                    Dim insertQuery As String = "INSERT INTO log (Username, Password, Credential, CreatedBy) VALUES (@username, @password, @credential, @CreatedBy)"
                    Using cmd As New SqlCommand(insertQuery, connection)
                        cmd.Parameters.AddWithValue("@username", userTxt.Text)
                        cmd.Parameters.AddWithValue("@password", passTxt.Text)
                        cmd.Parameters.AddWithValue("@credential", rightsCombo.SelectedItem)
                        cmd.Parameters.AddWithValue("@CreatedBy", Main_Panel.userTxt.Text)

                        cmd.ExecuteNonQuery()
                    End Using
                End Using

                ' Step 3: Show message based on role
                If rightsCombo.SelectedItem = "Admin" Then
                    MsgBox("Admin credentials created successfully")
                Else
                    MsgBox("User credentials created successfully")
                End If
            End If
        Catch ex As Exception
            MsgBox("Unable to insert record: " & ex.Message)
        End Try


        'Try

        '    Try
        '        If userTxt.Text = Nothing Or passTxt.Text = Nothing Then
        '            MsgBox("Enter username and password to proceed")
        '        Else
        '            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
        '            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
        '            Dim connection As New SqlConnection(Main_Panel.connectionString)

        '            connection.Open()
        '            Dim query As String = "INSERT INTO log (Username, Password, Credential) VALUES (@username, @password, @credential)"
        '            Dim cmd As New SqlCommand(query, connection)
        '            cmd.Parameters.AddWithValue("@username", userTxt.Text)
        '            cmd.Parameters.AddWithValue("@password", passTxt.Text)
        '            cmd.Parameters.AddWithValue("@credential", rightsCombo.SelectedItem)
        '            Dim result As Integer = cmd.ExecuteNonQuery()
        '            connection.Close()

        '        End If

        '        If rightsCombo.SelectedItem = "Admin" Then
        '            MsgBox("Admin credentials created successfully")
        '        Else
        '            MsgBox("User credentials created successfully")
        '        End If

        '    Catch ex As Exception
        '        MsgBox("Unable to insert record" & ex.ToString)
        '    End Try
        'Catch ex As Exception

        'End Try

        'Try
        '    Dim username As String = userTxt.Text.Trim()
        '    Dim mainFolder As String = "D:\Cadopt_Tool_Output"
        '    Dim userFolder As String = Path.Combine(mainFolder, username)

        '    ' 🔍 Step 1: Check and create main folder if it doesn't exist
        '    If Not Directory.Exists(mainFolder) Then
        '        Directory.CreateDirectory(mainFolder)
        '        'MessageBox.Show("Main folder 'Cadopt_Tool_Output' was missing. It has now been created.", "Folder Created", MessageBoxButtons.OK, MessageBoxIcon.Information)
        '    End If

        '    ' 📁 Step 2: Now create user folder inside main folder
        '    If Not Directory.Exists(userFolder) Then
        '        Directory.CreateDirectory(userFolder)
        '        'MessageBox.Show("User created and folder created at: " & userFolder)
        '    Else
        '        'MessageBox.Show("User created. Folder already exists.")
        '    End If

        'Catch ex As Exception
        '    MessageBox.Show("An error occurred while creating folder: " & ex.Message)
        'End Try

        Try
            Dim username As String = userTxt.Text.Trim()
            Dim serverIP As String = "192.168.7.70"
            Dim basePath As String = ""

            ' 🔍 Check if D drive exists on server
            If Directory.Exists("\\192.168.7.70\D$\Cadopt_Tool_Output") Then
                basePath = "\\192.168.7.70\D$\Cadopt_Tool_Output"
            ElseIf Directory.Exists("\\192.168.7.70\C$\Cadopt_Tool_Output") Then
                basePath = "\\192.168.7.70\C$\Cadopt_Tool_Output"
            Else
                MessageBox.Show("Neither D nor C drive folder found on server.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ' 📁 Create user folder path
            Dim userFolder As String = System.IO.Path.Combine(basePath, username)

            ' 📁 Create if not exists
            If Not System.IO.Directory.Exists(userFolder) Then
                System.IO.Directory.CreateDirectory(userFolder)
            End If

        Catch ex As Exception
            MessageBox.Show("An error occurred while creating folder on server: " & ex.Message)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles back_btn.Click

        Main_Panel.Show()
        Me.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles reset_btn.Click

        userTxt.Clear()
        passTxt.Clear()
    End Sub

    Private Sub rightsCombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rightsCombo.SelectedIndexChanged

        rightsCombo.DropDownStyle = ComboBoxStyle.DropDownList
    End Sub

    Private Sub LoadUserData()

        Try
            'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
            'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
            Dim connection As New SqlConnection(Main_Panel.connectionString)

            connection.Open()
            Dim query As String = "SELECT * from log"
            Dim cmd As New SqlCommand(query, connection)
            Dim reader As SqlDataReader = cmd.ExecuteReader()

            DataGridView1.Rows.Clear()

            Dim sno As Integer = 1

            While reader.Read()
                Dim username As String = reader("Username").ToString()
                Dim password As String = reader("Password").ToString()
                Dim credential As String = reader("Credential").ToString()
                Dim createdBy As String = reader("CreatedBy").ToString()
                Dim dateCreated As String = DateTime.Now.ToString("yyyy-MM-dd")
                Dim rights As String = credential

                DataGridView1.Rows.Add(sno, username, password, createdBy, dateCreated, rights)
                sno += 1
            End While

        Catch ex As Exception
            MsgBox("err " & ex.ToString())
        End Try

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadUserData()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles edit_btn.Click
        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                selectedRow.Tag = selectedRow.Cells(1).Value.ToString()
                DataGridView1.ReadOnly = False
                DataGridView1.CurrentCell = selectedRow.Cells(1)
            Else
                MsgBox("Please select a row to edit.")
            End If
        Catch ex As Exception
            MsgBox("err" & ex.ToString())
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles update_btn.Click

        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                If selectedRow.Tag IsNot Nothing Then
                    Dim oldUsername As String = selectedRow.Tag.ToString()
                    Dim newUsername As String = selectedRow.Cells(1).Value.ToString()
                    Dim password As String = selectedRow.Cells(2).Value.ToString()
                    'Dim credential As String = selectedRow.Cells(3).Value.ToString()
                    'Dim createdBy As String = selectedRow.Cells(4).Value.ToString() 
                    Dim dateCreated As String = selectedRow.Cells(5).Value.ToString()
                    'Dim rights As String = credential 

                    'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
                    'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
                    Dim connection As New SqlConnection(Main_Panel.connectionString)

                    connection.Open()
                    Dim query As String = "UPDATE log SET Username=@newUsername, Password=@password WHERE Username=@oldUsername"
                    Dim cmd As New SqlCommand(query, connection)
                    cmd.Parameters.AddWithValue("@newUsername", newUsername)
                    cmd.Parameters.AddWithValue("@password", password)
                    'cmd.Parameters.AddWithValue("@credential", credential)
                    cmd.Parameters.AddWithValue("@oldUsername", oldUsername)

                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    connection.Close()

                    If rowsAffected > 0 Then
                        MsgBox("Data updated successfully!")
                        DataGridView1.ReadOnly = True

                        selectedRow.Cells(5).Value = dateCreated

                        ' Refresh the DataGridView to reflect changes
                        LoadUserData()
                    Else
                        MsgBox("Update failed. No matching record found.")
                    End If
                Else
                    MsgBox("Original username not found. Please select a row and click edit again.")
                End If
            Else
                MsgBox("Please select a row to update.")
            End If


            LoadUserData()

        Catch ex As Exception
            MsgBox("err" & ex.ToString())
        End Try
    End Sub

    Private Sub delete_btn_Click(sender As Object, e As EventArgs) Handles delete_btn.Click

        Try
            If DataGridView1.SelectedRows.Count > 0 Then
                Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
                Dim username As String = selectedRow.Cells(1).Value.ToString()

                'Dim connectionString As String = "Server=CTPL-HW-198\SQLEXPRESS;Initial Catalog=Indo_tech;User ID=sqladmin;Password=Cadopt@123"
                'Dim connectionString As String = "Server=192.168.7.70,1433;Database=Indo_tech;User ID=sa;Password=Indo@123"
                Dim connection As New SqlConnection(Main_Panel.connectionString)

                connection.Open()
                Dim query As String = "DELETE FROM log WHERE Username=@username"
                Dim cmd As New SqlCommand(query, connection)
                cmd.Parameters.AddWithValue("@username", username)

                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                connection.Close()

                If rowsAffected > 0 Then
                    MsgBox("Data deleted successfully!")

                    DataGridView1.Rows.Remove(selectedRow)
                Else
                    MsgBox("Delete failed. No matching record found.")
                End If
            Else
                MsgBox("Please select a row to delete.")
            End If
        Catch ex As Exception
            MsgBox("err" & ex.ToString())
        End Try
    End Sub

    Private Sub refresh_btn_Click(sender As Object, e As EventArgs) Handles refresh_btn.Click
        DataGridView1.ReadOnly = True
        LoadUserData()
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Main_Panel.Show()
    End Sub

End Class
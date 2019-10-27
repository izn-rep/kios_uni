Imports Npgsql

Public Class User

    Public LastSelected As Object
    Public str1 As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
    Dim conn As New NpgsqlConnection
    Dim cmd As New NpgsqlCommand
    Dim DataSetPostgre As New DataSet
    Dim AdapterPostgre As New NpgsqlDataAdapter
    Dim myData As NpgsqlDataReader
    Dim query As String

    Public Function koneksi() As Boolean
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If
        conn.ConnectionString = str1
        Try
            conn.Open()
            cmd.Connection = conn
            Return True
        Catch ex As Exception
            conn.Close()
            Return False
        End Try
    End Function

    Public Function CmdSQL(ByVal perintah As String) As Boolean
        If koneksi() = True Then
            Try
                cmd.CommandText = perintah
                cmd.ExecuteNonQuery()
                AdapterPostgre = New NpgsqlDataAdapter(cmd)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End If
        Return Nothing
    End Function

    Public Sub ViewTabel(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
        If conn.State = ConnectionState.Open Then
            conn.Close()
        End If

        conn.ConnectionString = str1

        Try
            conn.Open()
            cmd.Connection = conn
            cmd.CommandText = query
            DataSetPostgre = New DataSet("namadataset")
            AdapterPostgre = New NpgsqlDataAdapter(cmd)
            AdapterPostgre.Fill(DataSetPostgre, datatable)
            namadg.DataSource = DataSetPostgre.Tables(datatable)
            namadg.AutoGenerateColumns = True
            conn.Close()
        Catch ex As Exception
            conn.Close()
        End Try

        'Set Header
        DataGridView1.Columns(0).HeaderText = "Nama Pengguna"
        DataGridView1.Columns(1).HeaderText = "Otoritas"


    End Sub

    Sub bukatabel()
        ViewTabel("DataTable", DataGridView1, "select nama_pengguna, otoritas from pengguna")
    End Sub

    Private Sub btnButton50_Click(sender As Object, e As EventArgs) Handles btnButton50.Click
        Try
            If TextBox1.Text = "" Or TextBox1.TextLength < 1 Then
                MsgBox("Nama Pengguna Wajib Diisi!")
                conn.Close()
            ElseIf ComboBox1.Text = "-Pilih Otoritas-" Then
                MsgBox("Pilih Otoritas")
                conn.Close()
            ElseIf TextBox2.Text = "" Or TextBox2.TextLength < 1 Then
                MsgBox("Kata Sandi Wajib Diisi!")
                conn.Close()
            ElseIf TextBox2.Text <> TextBox3.Text Then
                TextBox2.Text = ""
                TextBox3.Text = ""
                MsgBox("Kata Sandi Tidak Cocok/Sama!")
                conn.Close()
            Else
                If CmdSQL("insert into pengguna (nama_pengguna, kata_sandi, otoritas) values ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "')") = False Then
                    MsgBox("Error Saat Membuat User Baru!")
                    reset()
                    conn.Close()
                Else
                    MsgBox("User Telah Dibuat")
                    bukatabel()
                    reset()
                    conn.Close()
                End If
            End If
        Catch ex As NpgsqlException
            MsgBox(ex.Message)
            conn.Close()
        End Try

    End Sub

    Private Sub User_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        bukatabel()
        ComboBox1.Text = "-Pilih Otoritas-"
        LastSelected = Me.TextBox1
        TextBox1.Focus()
    End Sub

    Sub reset()
        TextBox1.Text = ""
        ComboBox1.Text = "-Pilih Otoritas-"
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub btnButton51_Click(sender As Object, e As EventArgs) Handles btnButton51.Click
        Me.Close()
        Admin.Show()
    End Sub

    Private Sub btnButton49_Click(sender As Object, e As EventArgs) Handles btnButton49.Click
        If TextBox1.Text = "" Or TextBox1.TextLength < 0 Then
            MsgBox("Isi Nama Pengguna, Terlebih Dahulu!")
            TextBox1.Focus()
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus User Ini?", "Kios UNI V.3", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then

                If CmdSQL("delete from pengguna where nama_pengguna ='" & TextBox1.Text & "'") = False Then
                    MsgBox("Error Saat Menghapus User!")
                    reset()
                Else
                    MsgBox("Pengguna Telah Dihapus!")
                    bukatabel()
                    reset()
                End If
            End If
        End If
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button48.Click, Button47.Click, Button46.Click, Button45.Click, Button44.Click, Button43.Click, Button42.Click, Button41.Click, Button40.Click, Button4.Click, Button39.Click, Button38.Click, Button37.Click, Button36.Click, Button35.Click, Button34.Click, Button33.Click, Button32.Click, Button31.Click, Button30.Click, Button3.Click, Button29.Click, Button28.Click, Button26.Click, Button25.Click, Button24.Click, Button23.Click, Button22.Click, Button21.Click, Button20.Click, Button2.Click, Button19.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
        If ShiftR.FlatStyle = FlatStyle.Flat Then
            LastSelected.Focus()
            ShiftR.PerformClick()
            SendKeys.Send("+" & sender.text)
        Else
            LastSelected.Focus()
            SendKeys.Send(sender.text)
        End If
    End Sub

    Private Sub Back_Click(sender As Object, e As EventArgs) Handles Back.Click
        LastSelected.Focus()
        SendKeys.Send("{BACKSPACE}")
    End Sub

    Private Sub ShiftR_Click(sender As Object, e As EventArgs) Handles ShiftL.Click, ShiftR.Click
        If ShiftR.FlatStyle = FlatStyle.Flat Then
            ShiftR.FlatStyle = FlatStyle.Standard
            ShiftL.FlatStyle = FlatStyle.Standard
            For Each ctl As Control In Me.Controls
                If (ctl.Name.StartsWith("Button")) Then
                    Dim btn As Button = DirectCast(ctl, Button)
                    btn.Text = btn.Text.ToLower
                    Button1.Text = "1"
                    Button2.Text = "2"
                    Button3.Text = "3"
                    Button4.Text = "4"
                    Button5.Text = "5"
                    Button6.Text = "6"
                    Button7.Text = "7"
                    Button8.Text = "8"
                    Button9.Text = "9"
                    Button10.Text = "0"
                    Button11.Text = "-"
                    Button12.Text = "="
                    Button13.Text = "`"
                    Button14.Text = "]"
                    Button15.Text = "\"
                    Button16.Text = "["
                    Button29.Text = "'"
                    Button30.Text = ";"
                    Button28.Text = "/"
                    Button40.Text = "."
                    Button41.Text = ","
                    Button27.Text = "Enter"
                End If
            Next
        ElseIf ShiftR.FlatStyle = FlatStyle.Standard Then
            ShiftL.FlatStyle = FlatStyle.Flat
            ShiftR.FlatStyle = FlatStyle.Flat
            For Each ctl As Control In Me.Controls
                If (ctl.Name.StartsWith("Button")) Then
                    Dim btn As Button = DirectCast(ctl, Button)
                    btn.Text = btn.Text.ToUpper
                    Button1.Text = "!"
                    Button2.Text = "@"
                    Button3.Text = "#"
                    Button4.Text = "$"
                    Button5.Text = "%"
                    Button6.Text = "^"
                    Button7.Text = "&&"
                    Button8.Text = "*"
                    Button9.Text = "("
                    Button10.Text = ")"
                    Button11.Text = "_"
                    Button12.Text = "+"
                    Button13.Text = "~"
                    Button14.Text = "}"
                    Button15.Text = "|"
                    Button16.Text = "{"
                    Button29.Text = """"
                    Button30.Text = ":"
                    Button28.Text = "?"
                    Button40.Text = ">"
                    Button41.Text = "<"
                    Button27.Text = "Enter"
                End If
            Next
        End If
    End Sub

    Private Sub Caps_Click(sender As Object, e As EventArgs) Handles Caps.Click
        If Caps.FlatStyle = FlatStyle.Flat Then
            Caps.FlatStyle = FlatStyle.Standard
            Caps.BackColor = Color.FromKnownColor(KnownColor.Control)
            For Each ctl As Control In Me.Controls
                If (ctl.Name.StartsWith("Button")) Then
                    Dim btn As Button = DirectCast(ctl, Button)
                    btn.Text = btn.Text.ToLower
                    Button1.Text = "1"
                    Button2.Text = "2"
                    Button3.Text = "3"
                    Button4.Text = "4"
                    Button5.Text = "5"
                    Button6.Text = "6"
                    Button7.Text = "7"
                    Button8.Text = "8"
                    Button9.Text = "9"
                    Button10.Text = "0"
                    Button11.Text = "-"
                    Button12.Text = "="
                    Button13.Text = "`"
                    Button14.Text = "]"
                    Button15.Text = "\"
                    Button16.Text = "["
                    Button29.Text = "'"
                    Button30.Text = ";"
                    Button28.Text = "/"
                    Button40.Text = "."
                    Button41.Text = ","
                    Button27.Text = "Enter"
                End If
            Next
        ElseIf Caps.FlatStyle = FlatStyle.Standard Then
            Caps.FlatStyle = FlatStyle.Flat
            Caps.BackColor = Color.LightGreen
            For Each ctl As Control In Me.Controls
                If (ctl.Name.StartsWith("Button")) Then
                    Dim btn As Button = DirectCast(ctl, Button)
                    btn.Text = btn.Text.ToUpper
                    Button1.Text = "1"
                    Button2.Text = "2"
                    Button3.Text = "3"
                    Button4.Text = "4"
                    Button5.Text = "5"
                    Button6.Text = "6"
                    Button7.Text = "7"
                    Button8.Text = "8"
                    Button9.Text = "9"
                    Button10.Text = "0"
                    Button11.Text = "-"
                    Button12.Text = "="
                    Button13.Text = "`"
                    Button14.Text = "]"
                    Button15.Text = "\"
                    Button16.Text = "["
                    Button29.Text = "'"
                    Button30.Text = ";"
                    Button28.Text = "/"
                    Button40.Text = "."
                    Button41.Text = ","
                    Button27.Text = "Enter"
                End If
            Next
        End If
        Beep()
    End Sub

    Private Sub Space_Click(sender As Object, e As EventArgs) Handles Space.Click
        LastSelected.Focus()
        SendKeys.Send(" ")
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click, TextBox2.Click, TextBox3.Click
        LastSelected = sender
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        LastSelected.Focus()
        SendKeys.Send("{ENTER}")
    End Sub

    Private Sub Tab_Click(sender As Object, e As EventArgs) Handles Tab.Click
        LastSelected.Focus()
        SendKeys.Send("^{TAB}")
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Bantuan.Show()
    End Sub
End Class
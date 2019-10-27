Imports Npgsql

Public Class Home

    Public LastSelected As Object
    Dim val As Decimal
    Public str1 As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
    Dim conn As New NpgsqlConnection
    Dim cmd As New NpgsqlCommand
    Dim DataSetPostgre As New DataSet
    Dim AdapterPostgre As New NpgsqlDataAdapter
    Dim myData As NpgsqlDataReader
    Dim query As String
    Dim total As Double = 0

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

    Sub reset()
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        Label11.Text = "0"
        DateTimePicker1.ResetText()
    End Sub

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
        'Menyembunyikan Kolom ID
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(8).Visible = False

        'Set Header
        DataGridView1.Columns(1).HeaderText = "Nama"
        DataGridView1.Columns(2).HeaderText = "Keterangan"
        DataGridView1.Columns(3).HeaderText = "Hari"
        DataGridView1.Columns(4).HeaderText = "Tanggal"
        DataGridView1.Columns(5).HeaderText = "Jam"
        DataGridView1.Columns(6).HeaderText = "Nominal"
        DataGridView1.Columns(7).HeaderText = "Kasir"

        Me.DataGridView1.Font = New Font("Cambria", 14)

    End Sub

    Sub bukatabel()
        ViewTabel("DataTable", DataGridView1, "select * from harian where tanggal = '" & DateTimePicker1.Text & "' and status = '1' order by no desc")
        Me.DataGridView1.Columns(6).DefaultCellStyle.Format = "C"
        Me.DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        If DataGridView1.RowCount > 0 Then
            Dim nom As Double = 0
            For index As Integer = 0 To DataGridView1.RowCount - 1
                nom += Convert.ToDouble(DataGridView1.Rows(index).Cells(6).Value)
            Next
            'Label16.Text = nom
            Label16.Text = Format(nom, "##,##0")
        Else
            Label16.Text = "0"
        End If
    End Sub
    Private Sub Home_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LastSelected = Me.TextBox2
        DateTimePicker1.ResetText()
        bukatabel()
        TextBox2.Focus()
        Try
            conn.Open()
            query = ("select * from pengguna where nama_pengguna = '" & LogIn.TextBox1.Text & "' and otoritas = 'Administrator'")
            cmd.Connection = conn
            cmd.CommandText = query
            AdapterPostgre.SelectCommand = cmd
            Try
                myData = cmd.ExecuteReader()
                If myData.Read() Then
                    tmAdmin.Visible = True
                    'MsgBox("welome admin")
                    conn.Close()
                Else
                    tmAdmin.Visible = False
                    'MsgBox("welcom user")
                    conn.Close()
                End If
            Catch ex As NpgsqlException

                conn.Close()
            End Try
        Catch ex As Exception

            conn.Close()
        End Try
        Dim Hari As String

        Hari = DateTime.Now.DayOfWeek
        If Hari = "1" Then
            Label10.Text = "Senin"
        ElseIf Hari = "2" Then
            Label10.Text = "Selasa"
        ElseIf Hari = "3" Then
            Label10.Text = "Rabu"
        ElseIf Hari = "4" Then
            Label10.Text = "Kamis"
        ElseIf Hari = "5" Then
            Label10.Text = "Jumat"
        ElseIf Hari = "6" Then
            Label10.Text = "Sabtu"
        ElseIf Hari = "0" Then
            Label10.Text = "Minggu"
        Else
            Label10.Text = "Libur"
        End If
        bukatabel()
    End Sub

    Private Sub tmLogOut_Click(sender As Object, e As EventArgs) Handles tmLogOut.Click
        Label1.Text = ""
        Me.Close()
        LogIn.Show()
    End Sub

    Private Sub tmUbahSandi_Click(sender As Object, e As EventArgs) Handles tmUbahSandi.Click
        Me.Hide()
        UbahSandi.Show()
    End Sub

    Private Sub tmAdmin_Click(sender As Object, e As EventArgs) Handles tmAdmin.Click
        Admin.Show()
        Me.Hide()
    End Sub

    Private Sub tmSimpan_Click(sender As Object, e As EventArgs) Handles tmSimpan.Click
        Dim hari As String
        If (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Monday) Then
            hari = "Senin"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Tuesday) Then
            hari = "Selasa"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Wednesday) Then
            hari = "Rabu"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Thursday) Then
            hari = "Kamis"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Friday) Then
            hari = "Jumat"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Saturday) Then
            hari = "Sabtu"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Sunday) Then
            hari = "Minggu"
        Else
            hari = "Libur"
        End If

        Try
            If TextBox2.Text = "" Or TextBox2.TextLength < 1 Then
                MsgBox("Nama Pelanggan Wajib Diisi!")
                TextBox2.Focus()
            ElseIf TextBox3.Text = "" Or TextBox3.TextLength < 1 Then
                MsgBox("Keterangan Wajib Diisi!")
                TextBox3.Focus()
            Else
                If CmdSQL("insert into harian (nama, keterangan, hari, tanggal, jam, nominal, pengguna,status) values ('" & TextBox2.Text & "','" & TextBox3.Text & "','" & hari & "','" & DateTimePicker1.Text & "','" & Label4.Text & "','" & TextBox4.Text & "','" & Label1.Text & "','1')") = False Then
                    MsgBox("Error Saat Proses Transaksi, Hubungi Administrator!")
                Else
                    bukatabel()
                    reset()
                End If
            End If
        Catch ex As NpgsqlException
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        TextBox2.Text = DataGridView1.SelectedCells.Item(1).Value.ToString()
        TextBox3.Text = DataGridView1.SelectedCells.Item(2).Value.ToString()
        TextBox4.Text = DataGridView1.SelectedCells.Item(6).Value.ToString()
        Label11.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
        DateTimePicker1.Text = DataGridView1.SelectedCells.Item(4).Value.ToString()
        Label12.Text = DataGridView1.SelectedCells.Item(4).Value.ToString()
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        TextBox2.Text = DataGridView1.SelectedCells.Item(1).Value.ToString()
        TextBox3.Text = DataGridView1.SelectedCells.Item(2).Value.ToString()
        TextBox4.Text = DataGridView1.SelectedCells.Item(6).Value.ToString()
        Label11.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
        DateTimePicker1.Text = DataGridView1.SelectedCells.Item(4).Value.ToString()
        Label12.Text = DataGridView1.SelectedCells.Item(4).Value.ToString()
    End Sub

    Private Sub tmHapus_Click(sender As Object, e As EventArgs) Handles tmHapus.Click
        If Label11.Text = "0" Then
            MsgBox("Pilih Transaksi Yang Akan Dihapus Terlebih Dahulu!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Data Ini?", "Kios UNI V.3", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("UPDATE harian SET status ='2' where no ='" & Label11.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Menghapus Data, Hubungi Administrator!")
                Else
                    bukatabel()
                    reset()
                    Label11.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub tmCari_Click(sender As Object, e As EventArgs) Handles tmCari.Click
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""

        Dim mode As String = ""

        If ComboBox1.Text = "Tanggal" Then
            mode = "tanggal"
        ElseIf ComboBox1.Text = "Nama" Then
            mode = "nama"
        ElseIf ComboBox1.Text = "Keterangan" Then
            mode = "keterangan"
        ElseIf ComboBox1.Text = "Nominal" Then
            mode = "nominal"
        ElseIf ComboBox1.Text = "Kasir" Then
            mode = "pengguna"
        End If

        ViewTabel("DataTabel", DataGridView1, "select * from harian where " & mode & " like '%" & TextBox5.Text & "%' and status ='1' order by no desc")

        If DataGridView1.RowCount > 0 Then
            Dim nom As Double = 0
            For index As Integer = 0 To DataGridView1.RowCount - 1
                nom += Convert.ToDouble(DataGridView1.Rows(index).Cells(6).Value)
            Next
            'Label16.Text = nom
            Label16.Text = Format(nom, "##,##0")
        Else
            Label16.Text = "0"
        End If
    End Sub

    Private Sub tmBuku_Click(sender As Object, e As EventArgs) Handles tmBuku.Click
        Buku.Show()
        Me.Hide()
    End Sub

    Private Sub tmReset_Click(sender As Object, e As EventArgs) Handles tmReset.Click
        bukatabel()
        reset()
    End Sub

    Private Sub tmExit_Click(sender As Object, e As EventArgs) Handles tmExit.Click
        Dim result = MessageBox.Show("Anda Yakin Ingin Keluar?", "Kios UNI V.3", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then

        ElseIf result = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub tmUbah_Click(sender As Object, e As EventArgs) Handles tmUbah.Click
        Dim hari As String
        If (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Monday) Then
            hari = "Senin"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Tuesday) Then
            hari = "Selasa"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Wednesday) Then
            hari = "Rabu"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Thursday) Then
            hari = "Kamis"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Friday) Then
            hari = "Jumat"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Saturday) Then
            hari = "Sabtu"
        ElseIf (DateTimePicker1.Value.DayOfWeek = DayOfWeek.Sunday) Then
            hari = "Minggu"
        Else
            hari = "Libur"
        End If

        If Label11.Text = "0" Then
            MsgBox("Pilih Transaksi Yang Akan Diubah Terlebih Dahulu!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Mengubah Data Ini?", "Kios UNI V.3", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("UPDATE harian SET nama ='" & TextBox2.Text & "',keterangan = '" & TextBox3.Text & "',hari='" & hari & "',tanggal='" & DateTimePicker1.Text & "',nominal = '" & TextBox4.Text & "' WHERE no = '" & Label11.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Mengubah Data, Hubungi Administrator!")
                Else
                    bukatabel()
                    Label11.Text = "0"
                    reset()
                End If
            End If
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        'val = TextBox4.Text
        'TextBox4.Text = Format(val, "##,##0")
        'TextBox4.SelectionStart = Len(TextBox4.Text)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button5.Click, _
        Button48.Click, Button47.Click, Button46.Click, Button45.Click, Button44.Click, Button43.Click, Button42.Click, Button41.Click, Button40.Click, _
        Button4.Click, Button39.Click, Button38.Click, Button37.Click, Button36.Click, Button35.Click, Button34.Click, Button33.Click, Button32.Click, _
        Button31.Click, Button30.Click, Button3.Click, Button29.Click, Button28.Click, Button26.Click, Button25.Click, Button24.Click, Button23.Click, _
        Button22.Click, Button21.Click, Button20.Click, Button2.Click, Button19.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, _
        Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
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

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click, TextBox4.Click, TextBox3.Click, TextBox5.Click
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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Label4.Text = TimeString
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Bantuan.Show()
    End Sub

    Private valHolder1 As Double 'variable to hold operands
    Private valHolder2 As Double
    Private tmpValue As Double 'variable temporary values
    Private hasDecimal As Boolean
    Private inputStatus As Boolean
    Private clearText As Boolean
    Private calcFunc As String

    Private Sub CalculateTotals()
        valHolder2 = CDbl(txtInput.Text)
        Select Case calcFunc
            Case "Add"
                valHolder1 = valHolder1 + valHolder2
            Case "Subtract"
                valHolder1 = valHolder1 - valHolder2
            Case "Divide"
                valHolder1 = valHolder1 / valHolder2
            Case "Multiply"
                valHolder1 = valHolder1 * valHolder2
        End Select
        txtInput.Text = CStr(valHolder1)
        inputStatus = True
    End Sub

    Private Sub cmd9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd9.Click
        If inputStatus = False Then
            txtInput.Text += cmd9.Text
        Else
            txtInput.Text = cmd9.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd8.Click
        If inputStatus = False Then
            txtInput.Text += cmd8.Text
        Else
            txtInput.Text = cmd8.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd7.Click
        If inputStatus = False Then
            txtInput.Text += cmd7.Text
        Else
            txtInput.Text = cmd7.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd6.Click
        If inputStatus = False Then
            txtInput.Text += cmd6.Text
        Else
            txtInput.Text = cmd6.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd5.Click
        If inputStatus = False Then
            txtInput.Text += cmd5.Text
        Else
            txtInput.Text = cmd5.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd4.Click
        If inputStatus = False Then
            txtInput.Text += cmd4.Text
        Else
            txtInput.Text = cmd4.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd3.Click
        If inputStatus = False Then
            txtInput.Text += cmd3.Text
        Else
            txtInput.Text = cmd3.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd2.Click
        If inputStatus = False Then
            txtInput.Text += cmd2.Text
        Else
            txtInput.Text = cmd2.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd1.Click
        If inputStatus = False Then
            txtInput.Text += cmd1.Text
        Else
            txtInput.Text = cmd1.Text
            inputStatus = False
        End If
    End Sub

    Private Sub cmd0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd0.Click
        If inputStatus = False Then
            If txtInput.Text.Length >= 1 Then
                txtInput.Text += cmd0.Text
            End If
        End If
    End Sub

    Private Sub cmdDecimal_Click(sender As Object, e As EventArgs) Handles cmdDecimal.Click
        'Check for input status (we want flase)
        If Not inputStatus Then
            'Check if it already has a decimal (if it does then do nothing)
            If Not hasDecimal Then
                'Check to make sure the length is > than 1
                'Dont want user to add decimal as first character
                If txtInput.Text.Length > 0 Then
                    'Make sure 0 isnt the first number
                    If Not txtInput.Text = "0" Then
                        'It met all our requirements so add the zero
                        txtInput.Text += ","
                        'Toggle the flag to true (only 1 decimal per calculation)
                        hasDecimal = True
                    End If
                Else
                    txtInput.Text = "0,"
                End If
            End If
        End If
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Add"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdSubtract_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubtract.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Subtract"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdDivide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDivide.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Divide"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdMultiply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMultiply.Click
        If txtInput.Text.Length <> 0 Then
            If calcFunc = String.Empty Then
                valHolder1 = CDbl(txtInput.Text)
                txtInput.Text = String.Empty
            Else
                CalculateTotals()
            End If
            calcFunc = "Multiply"
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdEqual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEqual.Click
        If txtInput.Text.Length <> 0 AndAlso valHolder1 <> 0 Then
            CalculateTotals()
            calcFunc = ""
            hasDecimal = False
        End If
    End Sub

    Private Sub cmdClear_Click(sender As Object, e As EventArgs) Handles cmdClear.Click
        txtInput.Text = String.Empty
        valHolder1 = 0
        valHolder2 = 0
        calcFunc = String.Empty
        hasDecimal = False
    End Sub

    Private Sub cmdDel_Click(sender As Object, e As EventArgs) Handles cmdDel.Click
        Dim str As String
        Dim loc As Integer
        If txtInput.Text.Length > 0 Then
            str = txtInput.Text.Chars(txtInput.Text.Length - 1)
            If str = "." Then
                hasDecimal = False
            End If
            loc = txtInput.Text.Length
            txtInput.Text = txtInput.Text.Remove(loc - 1, 1)
        End If
    End Sub

    'Public Value1 As Long
    'Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
    '   Value1 = txtInput.Text
    '  txtInput.Text = Format(Value1, "##,##0")
    '   txtInput.SelectionStart = Len(txtInput.Text)
    ' End Sub
End Class
Imports Npgsql

Public Class Buku

    Public LastSelected As Object

    Public str1 As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
    Dim conn As New NpgsqlConnection
    Dim cmd As New NpgsqlCommand
    Dim DataSetPostgre As New DataSet
    Dim AdapterPostgre As New NpgsqlDataAdapter
    Dim myData As NpgsqlDataReader
    Dim query As String
    Dim total As Double = 0


    Private da As NpgsqlDataAdapter
    Private ds As DataSet
    Private dtSource As DataTable
    Private PageCount As Integer
    Private maxRec As Integer
    Private pageSize As Integer
    Private currentPage As Integer
    Private recNo As Integer

    Dim sSql As String

    Sub LoadDS(ByVal sSQL As String)
        Try
            Dim cnString As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
            Dim conn As NpgsqlConnection = New NpgsqlConnection(cnString)

            da = New NpgsqlDataAdapter(sSQL, conn)
            ds = New DataSet()

            da.Fill(ds, "Items")
            dtSource = ds.Tables("Items")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub DisplayPageInfo()
        Label17.Text = currentPage.ToString & "/ " & PageCount.ToString
    End Sub

    Sub LoadPage()
        Dim i As Integer
        Dim startRec As Integer
        Dim endRec As Integer
        Dim dtTemp As DataTable

        dtTemp = dtSource.Clone

        If currentPage = PageCount Then
            endRec = maxRec
        Else
            endRec = pageSize * currentPage
        End If

        startRec = recNo

        If dtSource.Rows.Count > 0 Then
            For i = startRec To endRec - 1
                dtTemp.ImportRow(dtSource.Rows(i))
                recNo = recNo + 1
            Next
        End If

        DataGridView1.DataSource = dtTemp
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "Nama"
        DataGridView1.Columns(2).HeaderText = "Alamat"
        DataGridView1.Columns(3).HeaderText = "No Telepon"
        DataGridView1.Columns(4).Visible = False

        Me.DataGridView1.Font = New Font("Cambria", 14)

        DisplayPageInfo()
    End Sub

    Sub FillGrid()
        pageSize = 5
        maxRec = dtSource.Rows.Count
        PageCount = maxRec \ pageSize

        If (maxRec Mod pageSize) > 0 Then
            PageCount = PageCount + 1
        End If

        currentPage = 1
        recNo = 0

        LoadPage()
    End Sub

    Private Function CheckFillButton() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            CheckFillButton = False
        Else
            CheckFillButton = True
        End If
    End Function


    Private da2 As NpgsqlDataAdapter
    Private ds2 As DataSet
    Private dtSource2 As DataTable
    Private PageCount2 As Integer
    Private maxRec2 As Integer
    Private pageSize2 As Integer
    Private currentPage2 As Integer
    Private recNo2 As Integer

    Dim sSql2 As String

    Sub LoadDS2(ByVal sSQL As String)
        Try
            Dim cnString As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
            Dim conn As NpgsqlConnection = New NpgsqlConnection(cnString)

            da2 = New NpgsqlDataAdapter(sSQL, conn)
            ds2 = New DataSet()

            da2.Fill(ds2, "Items")
            dtSource2 = ds2.Tables("Items")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub DisplayPageInfo2()
        Label18.Text = currentPage2.ToString & "/ " & PageCount2.ToString
    End Sub

    Sub LoadPage2()
        Dim i As Integer
        Dim startRec2 As Integer
        Dim endRec2 As Integer
        Dim dtTemp2 As DataTable

        dtTemp2 = dtSource2.Clone

        If currentPage2 = PageCount2 Then
            endRec2 = maxRec2
        Else
            endRec2 = pageSize2 * currentPage2
        End If

        startRec2 = recNo2

        If dtSource2.Rows.Count > 0 Then
            For i = startRec2 To endRec2 - 1
                dtTemp2.ImportRow(dtSource2.Rows(i))
                recNo2 = recNo2 + 1
            Next
        End If

        DataGridView2.DataSource = dtTemp2
        DataGridView2.Columns(0).Visible = False
        DataGridView2.Columns(1).HeaderText = "Keterangan"
        DataGridView2.Columns(2).HeaderText = "Hari"
        DataGridView2.Columns(3).HeaderText = "Tanggal"
        DataGridView2.Columns(4).HeaderText = "Nominal"
        DataGridView2.Columns(5).HeaderText = "Sisa"

        Me.DataGridView2.Font = New Font("Cambria", 14)
        Me.DataGridView2.Columns(4).DefaultCellStyle.Format = "c"
        Me.DataGridView2.Columns(5).DefaultCellStyle.Format = "c"
        Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        DisplayPageInfo2()
    End Sub

    Sub FillGrid2()
        pageSize2 = 6
        maxRec2 = dtSource2.Rows.Count
        PageCount2 = maxRec2 \ pageSize2

        If (maxRec2 Mod pageSize2) > 0 Then
            PageCount2 = PageCount2 + 1
        End If

        currentPage2 = 1
        recNo2 = 0

        LoadPage2()
    End Sub

    Private Function CheckFillButton2() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize2 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            CheckFillButton2 = False
        Else
            CheckFillButton2 = True
        End If
    End Function

    Sub table1()
        sSql = "Select * from pelanggan where status='1' and nama like '%" & TextBox6.Text & "%' order by nama"
        LoadDS(sSql)
        FillGrid()
    End Sub

    Sub table2()
        sSql2 = "Select no, keterangan, hari, tanggal, nominal, sisa from buku where id = '" & DataGridView1.SelectedCells.Item(0).Value.ToString() & "' and status = '1' order by no desc"
        Label12.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
        TextBox5.Text = DataGridView1.SelectedCells.Item(1).Value.ToString()
        TextBox1.Text = DataGridView1.SelectedCells.Item(2).Value.ToString()
        TextBox2.Text = DataGridView1.SelectedCells.Item(3).Value.ToString()
        LoadDS2(sSql2)
        FillGrid2()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        'If the user did not click the "Fill Grid" button then Return
        If Not CheckFillButton() Then Return
        'Check if the user clicked the "Fill Grid" button.
        If pageSize = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            Return
        End If

        currentPage = currentPage + 1

        If currentPage > PageCount Then
            currentPage = PageCount

            'Check if you are already at the last page.
            If recNo = maxRec Then
                MessageBox.Show("You are at the Last Page!")
                Return
            End If
        End If

        LoadPage()
    End Sub


    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click

        If Not CheckFillButton() Then Return

        currentPage = currentPage - 1

        'Check if you are already at the first page.
        If currentPage < 1 Then
            MessageBox.Show("You are at the First Page!")
            currentPage = 1
            Return
        Else
            recNo = pageSize * (currentPage - 1)
        End If

        LoadPage()
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click

        If Not CheckFillButton() Then Return

        ' Check if you are already at the first page.
        If currentPage = 1 Then
            MessageBox.Show("You are at the First Page!")
            Return
        End If

        currentPage = 1
        recNo = 0

        LoadPage()
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click

        If Not CheckFillButton() Then Return

        ' Check if you are already at the last page.
        If recNo = maxRec Then
            MessageBox.Show("You are at the Last Page!")
            Return
        End If

        currentPage = PageCount

        recNo = pageSize * (currentPage - 1)

        LoadPage()
    End Sub

    Private Sub btnFirst2_Click(sender As Object, e As EventArgs) Handles btnFirst2.Click

        If Not CheckFillButton2() Then Return

        ' Check if you are already at the first page.
        If currentPage2 = 1 Then
            MessageBox.Show("You are at the First Page!")
            Return
        End If

        currentPage2 = 1
        recNo2 = 0

        LoadPage2()
    End Sub

    Private Sub btnPrev2_Click(sender As Object, e As EventArgs) Handles btnPrev2.Click

        If Not CheckFillButton2() Then Return

        currentPage2 = currentPage2 - 1

        'Check if you are already at the first page.
        If currentPage2 < 1 Then
            MessageBox.Show("You are at the First Page!")
            currentPage2 = 1
            Return
        Else
            recNo2 = pageSize2 * (currentPage2 - 1)
        End If

        LoadPage2()
    End Sub

    Private Sub btnNext2_Click(sender As Object, e As EventArgs) Handles btnNext2.Click
        'If the user did not click the "Fill Grid" button then Return
        If Not CheckFillButton2() Then Return
        'Check if the user clicked the "Fill Grid" button.
        If pageSize2 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            Return
        End If

        currentPage2 = currentPage2 + 1

        If currentPage2 > PageCount2 Then
            currentPage2 = PageCount2

            'Check if you are already at the last page.
            If recNo2 = maxRec2 Then
                MessageBox.Show("You are at the Last Page!")
                Return
            End If
        End If

        LoadPage2()
    End Sub

    Private Sub btnLast2_Click(sender As Object, e As EventArgs) Handles btnLast2.Click

        If Not CheckFillButton2() Then Return

        ' Check if you are already at the last page.
        If recNo2 = maxRec2 Then
            MessageBox.Show("You are at the Last Page!")
            Return
        End If

        currentPage2 = PageCount2

        recNo2 = pageSize2 * (currentPage2 - 1)

        LoadPage2()
    End Sub


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
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""

    End Sub

    ' Public Sub ViewTabel1(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
    '    If conn.State = ConnectionState.Open Then
    '       conn.Close()
    '   End If

    '   conn.ConnectionString = str1

    '    Try
    '        conn.Open()
    '         cmd.Connection = conn
    '         cmd.CommandText = query
    '        DataSetPostgre = New DataSet("namadataset")
    '         AdapterPostgre = New NpgsqlDataAdapter(cmd)
    '       AdapterPostgre.Fill(DataSetPostgre, datatable)
    '        namadg.DataSource = DataSetPostgre.Tables(datatable)
    '        namadg.AutoGenerateColumns = True
    '        conn.Close()
    '   Catch ex As Exception
    '       conn.Close()
    '    End Try
    '   DataGridView1.Columns(0).Visible = False
    '   DataGridView1.Columns(1).HeaderText = "Nama"
    '  DataGridView1.Columns(2).HeaderText = "Alamat"
    '   DataGridView1.Columns(3).HeaderText = "No Telepon"
    '   DataGridView1.Columns(4).Visible = False

    '   Me.DataGridView1.Font = New Font("Cambria", 14)

    ' End Sub

    ' Public Sub ViewTabel2(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
    '    If conn.State = ConnectionState.Open Then
    '        conn.Close()
    '    End If

    '   conn.ConnectionString = str1
    '
    '      Try
    '       conn.Open()
    '        cmd.Connection = conn
    '       cmd.CommandText = query
    '       DataSetPostgre = New DataSet("namadataset")
    '      AdapterPostgre = New NpgsqlDataAdapter(cmd)
    '      AdapterPostgre.Fill(DataSetPostgre, datatable)
    '      namadg.DataSource = DataSetPostgre.Tables(datatable)
    '      namadg.AutoGenerateColumns = True
    '      conn.Close()
    '    Catch ex As Exception
    '       conn.Close()
    '    End Try

    '    DataGridView2.Columns(0).Visible = False
    '    DataGridView2.Columns(1).HeaderText = "Keterangan"
    '    DataGridView2.Columns(2).HeaderText = "Hari"
    '   DataGridView2.Columns(3).HeaderText = "Tanggal"
    ''   DataGridView2.Columns(4).HeaderText = "Nominal"
    '  DataGridView2.Columns(5).HeaderText = "Sisa"

    '    Me.DataGridView2.Font = New Font("Cambria", 14)
    '   Me.DataGridView2.Columns(4).DefaultCellStyle.Format = "c"
    '  Me.DataGridView2.Columns(5).DefaultCellStyle.Format = "c"
    '  Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '   Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '  End Sub

    ' Sub bukatable1()
    '     ViewTabel1("DataTabel", DataGridView1, "Select * from pelanggan where status='1' and nama like '%" & TextBox6.Text & "%' order by nama")
    'TextBox6.Text = ""
    ' End Sub

    '  Public Sub ViewTabelbuat(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
    '  If conn.State = ConnectionState.Open Then
    '     conn.Close()
    ' End If

    ' conn.ConnectionString = str1

    '  Try
    '     conn.Open()
    '     cmd.Connection = conn
    '    cmd.CommandText = query
    'DataSetPostgre = New DataSet("namadataset")
    '    AdapterPostgre = New NpgsqlDataAdapter(cmd)
    ' AdapterPostgre.Fill(DataSetPostgre, datatable)
    ''    namadg.DataSource = DataSetPostgre.Tables(datatable)
    '    namadg.AutoGenerateColumns = True
    '     conn.Close()
    '  Catch ex As Exception
    '    conn.Close()
    '  End Try
    '
    '  DataGridView2.Columns(0).Visible = False
    '  DataGridView2.Columns(1).HeaderText = "Keterangan"
    '  DataGridView2.Columns(2).HeaderText = "Hari"
    '  DataGridView2.Columns(3).HeaderText = "Tanggal"
    '  DataGridView2.Columns(4).HeaderText = "Nominal"
    '   DataGridView2.Columns(5).HeaderText = "Sisa"

    '   Me.DataGridView2.Font = New Font("Arial", 18)
    '   Me.DataGridView2.Columns(4).DefaultCellStyle.Format = "c"
    '   Me.DataGridView2.Columns(5).DefaultCellStyle.Format = "c"
    '    Me.DataGridView2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '     Me.DataGridView2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    '  End Sub

    ' Sub bukatablebuat()
    '     ViewTabel1("DataTabel", DataGridView1, "Select * from pelanggan where status='1' and nama like '%" & Buat.TextBox1.Text & "%' order by nama")
    '     TextBox6.Text = ""
    '  End Sub

    '  Sub bukatable2()
    '      ViewTabel2("DataTabel", DataGridView2, "Select no, keterangan, hari, tanggal, nominal, sisa from buku where id = '" & DataGridView1.SelectedCells.Item(0).Value.ToString() & "' and status = '1' order by no desc")
    '      Label12.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
    '     TextBox5.Text = DataGridView1.SelectedCells.Item(1).Value.ToString()
    '     TextBox1.Text = DataGridView1.SelectedCells.Item(2).Value.ToString()
    '     TextBox2.Text = DataGridView1.SelectedCells.Item(3).Value.ToString()
    ' End Sub

    Private Sub btnbkReset_Click(sender As Object, e As EventArgs) Handles btnbkReset.Click
        DataGridView1.DataSource = Nothing
        DataGridView2.DataSource = Nothing
        TextBox6.Text = ""
        TextBox5.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DateTimePicker1.ResetText()
        Label12.Text = "0"
        Label9.Text = "0"
        Label7.Text = "0"
        TextBox6.Focus()
        btnbkCari.Enabled = False
    End Sub

    Sub resetcari()
        Label7.Text = "0"
        Label9.Text = "0"
        Label12.Text = "0"
        TextBox5.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        DateTimePicker1.ResetText()
        DataGridView2.DataSource = Nothing
    End Sub


    Private Sub btnbkCari_Click(sender As Object, e As EventArgs) Handles btnbkCari.Click
        'bukatable1()
        resetcari()
        table1()
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Label9.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
        'bukatable2()
        table2()
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Label7.Text = DataGridView2.SelectedCells.Item(0).Value.ToString()
        TextBox3.Text = DataGridView2.SelectedCells.Item(1).Value.ToString()
        DateTimePicker1.Text = DataGridView2.SelectedCells.Item(3).Value.ToString()
    End Sub

    Private Sub btnbkKurang_Click(sender As Object, e As EventArgs) Handles btnbkKurang.Click
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

        Label10.Text = hari
        Dim sisa As Double
        If currentPage2 = "1" Then
            If TextBox5.Text = "" Then
                MsgBox("Klik Nama Pelanggan Terlebih Dahulu!")
            Else
                If DataGridView2.RowCount > 0 Then
                    sisa = DataGridView2.Rows(0).Cells(5).Value
                    Label4.Text = sisa
                ElseIf DataGridView2.RowCount <= 0 Then
                    sisa = "0"
                End If
                Try
                    If TextBox3.Text = "" Or TextBox3.TextLength < 1 Then
                        MsgBox("Keterangan Wajib Diisi!")
                        TextBox3.Focus()
                    ElseIf TextBox4.Text = "" Or TextBox4.TextLength < 1 Then
                        MsgBox("Nominal Wajib Diisi!")
                        TextBox4.Focus()
                    Else
                        Dim kurang As Double
                        kurang = sisa - TextBox4.Text
                        If CmdSQL("insert into buku (id,keterangan,hari,tanggal,nominal,sisa,status) values ('" & Label12.Text & "','" & TextBox3.Text & "','" & hari & "','" & DateTimePicker1.Text & "','" & TextBox4.Text & "','" & kurang & "','1')") = False Then
                            MsgBox("Terjadi Kesalahan Saat Simpan Data, Hubungi Administrator!")
                        Else
                            'bukatable2()
                            table2()
                            Dim result = MessageBox.Show("Berhasil Disimpan." & vbCrLf & "Simpan Juga ke Transaksi Harian ?", "Kios UNI V.3", MessageBoxButtons.YesNo)
                            If result = DialogResult.No Then
                                TextBox4.Text = ""
                                TextBox3.Text = ""
                            ElseIf result = DialogResult.Yes Then
                                CmdSQL("insert into harian (nama, keterangan, hari, tanggal, jam, nominal, pengguna, status) values ('" & TextBox5.Text & "','" & TextBox3.Text & "','" & hari & "','" & DateTimePicker1.Text & "','" & Home.Label4.Text & "','" & TextBox4.Text & "','" & Home.Label1.Text & "','1')")
                                TextBox4.Text = ""
                                TextBox3.Text = ""
                                DateTimePicker1.ResetText()
                                Home.bukatabel()
                            End If
                        End If
                    End If
                Catch ex As NpgsqlException
                    MsgBox(ex.Message)
                End Try
            End If
        Else
            MsgBox("Transaksi Harus Pada Halaman 1")
        End If

    End Sub

    Private Sub btnbkTambah_Click(sender As Object, e As EventArgs) Handles btnbkTambah.Click
        Dim hari As String
        Dim sisa As Double
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

        Label10.Text = hari

        If currentPage2 = "1" Then
            If TextBox5.Text = "" Then
                MsgBox("Cari Nama Pelanggan Terlebih Dahulu!")
            Else
                If DataGridView2.RowCount > 0 Then
                    sisa = DataGridView2.Rows(0).Cells(5).Value
                    Label4.Text = sisa
                ElseIf DataGridView2.RowCount <= 0 Then
                    sisa = "0"
                End If
                Try
                    If TextBox3.Text = "" Or TextBox3.TextLength < 1 Then
                        MsgBox("Keterangan Wajib Diisi!")
                        TextBox3.Focus()
                    ElseIf TextBox4.Text = "" Or TextBox4.TextLength < 1 Then
                        MsgBox("Nominal Wajib Diisi!")
                        TextBox4.Focus()
                    Else
                        Dim tambah As Double
                        tambah = TextBox4.Text + sisa
                        If CmdSQL("insert into buku (id,keterangan,hari,tanggal,nominal,sisa,status) values ('" & Label12.Text & "','" & TextBox3.Text & "','" & hari & "','" & DateTimePicker1.Text & "','" & TextBox4.Text & "','" & tambah & "','1')") = False Then
                            MsgBox("Terjadi Kesalahan Saat Simpan Data, Hubungi Administrator!")
                        Else
                            'bukatable2()
                            table2()
                            btnbkKurang.Visible = True
                            TextBox4.Text = ""
                            TextBox3.Text = ""
                            DateTimePicker1.ResetText()
                        End If
                    End If
                Catch ex As Exception
                End Try
            End If
        Else
            MsgBox("Transaksi Harus Pada Halaman 1")
        End If

    End Sub

    Private Sub btnbkHapus_Click(sender As Object, e As EventArgs) Handles btnbkHapus.Click
        Try
            If Label7.Text = "0" Then
                MsgBox("Pilih Transaksi Yang Akan Dihapus Terlebih Dahulu!")
            Else
                Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Data Ini?", "Kios UNI V.3", MessageBoxButtons.YesNo)

                If result = DialogResult.No Then

                ElseIf result = DialogResult.Yes Then
                    If CmdSQL("UPDATE buku SET status ='2' where no ='" & Label7.Text & "'") = False Then
                        MsgBox("Terjadi Kesalahan Saat Menghapus Data, Hubungi Administrator!")
                    Else
                        'bukatable2()
                        table2()
                        Label7.Text = "0"
                    End If
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnbkEditData_Click(sender As Object, e As EventArgs) Handles btnbkEditData.Click
        If Label9.Text = "0" Then
            MsgBox("Pilih Nama Pelanggan Terlebih Dahulu!")
        Else
            Data.Show()
            Me.Hide()
        End If
    End Sub

    Private Sub btnbkBuatBaru_Click(sender As Object, e As EventArgs) Handles btnbkBuatBaru.Click
        Buat.Show()
        Me.Hide()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Me.DateTimePicker1.Font = New Font("Cambria", 14)
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If TextBox6.Text = " " Then
            btnbkCari.Enabled = False
        Else
            btnbkCari.Enabled = True
        End If
    End Sub

    Private Sub btnbkKembali_Click(sender As Object, e As EventArgs) Handles btnbkKembali.Click
        Me.Close()
        Home.Show()
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

    Private Sub TextBox6_Click(sender As Object, e As EventArgs) Handles TextBox6.Click, TextBox3.Click, TextBox4.Click, DateTimePicker1.Click
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

    Private Sub Buku_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LastSelected = Me.TextBox6
        btnbkCari.Enabled = False
        DateTimePicker1.ResetText()
        TextBox6.Focus()
    End Sub

    Private Sub btnHelp_Click(sender As Object, e As EventArgs) Handles btnHelp.Click
        Bantuan.Show()
    End Sub

    Private Sub btnbkubah_Click(sender As Object, e As EventArgs) Handles btnbkubah.Click
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
            If Label7.Text = "0" Then
                MsgBox("Pilih Transaksi Yang Akan Diubah Terlebih Dahulu!")
            Else
                Dim result = MessageBox.Show("Anda Yakin Ingin Mengubah Data Ini?", "Kios UNI V.3", MessageBoxButtons.YesNo)

                If result = DialogResult.No Then

                ElseIf result = DialogResult.Yes Then
                    If CmdSQL("UPDATE buku SET keterangan='" & TextBox3.Text & "',hari='" & hari & "',tanggal='" & DateTimePicker1.Text & "' where no ='" & Label7.Text & "'") = False Then
                        MsgBox("Terjadi Kesalahan Saat Mengubah Data, Hubungi Administrator!")
                    Else
                        'bukatable2()
                        table2()
                        Label7.Text = "0"
                        TextBox3.Text = ""
                        DateTimePicker1.ResetText()
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
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
End Class
Imports Npgsql

Public Class Admin

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

    End Sub

    Sub bukatabel()
        ViewTabel("DataTable", DataGridView1, "select * from harian where status = '2' order by no desc")
    End Sub

    Public Sub ViewTabel2(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
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


    End Sub

    Sub bukatable2()
        ViewTabel2("DataTabel", DataGridView2, "Select * from buku where status = '2' order by no desc")
    End Sub

    Public Sub ViewTabel3(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
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

    End Sub

    Sub bukatabel3()
        ViewTabel3("DataTable", DataGridView3, "select * from pelanggan where status = '2' order by nama")
    End Sub

    Public Sub ViewTabel4(ByVal datatable As String, ByVal namadg As DataGridView, ByVal query As String)
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

    End Sub

    Sub bukatabel4()
        ViewTabel4("DataTable", DataGridView4, "select * from pelanggan where status = '1' order by nama")
    End Sub

    Private Sub DataGridView1_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseClick
        Label3.Text = DataGridView1.SelectedCells.Item(0).Value.ToString()
    End Sub

    Private Sub DataGridView2_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseClick
        Label4.Text = DataGridView2.SelectedCells.Item(0).Value.ToString()
    End Sub

    Private Sub DataGridView3_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.RowHeaderMouseClick
        Label6.Text = DataGridView3.SelectedCells.Item(0).Value.ToString()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        User.Show()
        Me.Hide()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Home.Show()
        Me.Close()
    End Sub

    Private Sub Admin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        table1()
        table2()
        table3()
        table4()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Label3.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Direstore!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Merestore Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("UPDATE harian SET status ='1' where no ='" & Label3.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Merestore Data!")
                Else
                    table3()
                    MsgBox("Transaksi Berhasil Direstore!")
                    Label3.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Label4.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Direstore!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Merestore Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("UPDATE buku SET status ='1' where no ='" & Label4.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Merestore Data!")
                Else
                    table4()
                    MsgBox("Transaksi Berhasil Direstore!")
                    Label4.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Label3.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Dihapus!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("Delete from harian where no ='" & Label3.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Menghapus Data!")
                Else
                    table3()
                    MsgBox("Transaksi Berhasil Dihapus!")
                    Label3.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Label4.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Dihapus!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("Delete from harian where no ='" & Label4.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Menghapus Data!")
                Else
                    table4()
                    MsgBox("Transaksi Berhasil Dihapus!")
                    Label4.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            If DataGridView1.RowCount < 2 Then
                MsgBox("Data Kosong, Tidak Dapat Dihapus!")
            Else
                Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Permanen Semua Data?", "Konfirmasi", MessageBoxButtons.YesNo)

                If result = DialogResult.No Then

                ElseIf result = DialogResult.Yes Then
                    If CmdSQL("Delete from harian where status ='2'") = False Then
                        MsgBox("Terjadi Kesalahan Saat Menghapus Semua Data!")
                    Else
                        table3()
                        MsgBox("Transaksi Berhasil Dihapus Semua!")
                    End If
                End If
            End If
        Catch ex As NpgsqlException
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            If DataGridView2.RowCount < 2 Then
                MsgBox("Data Kosong, Tidak Dapat Dihapus!")
            Else
                Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Permanen Semua Data?", "Konfirmasi", MessageBoxButtons.YesNo)

                If result = DialogResult.No Then

                ElseIf result = DialogResult.Yes Then
                    If CmdSQL("Delete from buku where status ='2'") = False Then
                        MsgBox("Terjadi Kesalahan Saat Menghapus Semua Data!")
                    Else
                        table4()
                        MsgBox("Transaksi Berhasil Dihapus Semua!")
                    End If
                End If
            End If
        Catch ex As NpgsqlException
            MsgBox(ex.Message)
        End Try
    End Sub


    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        table1()
        table2()
        table3()
        table4()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        MsgBox("Maaf, Sementara Fungsi Ini Belum Tersedia" & vbCrLf & "Silahkan Lakukan Backup Manual!")
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If Label6.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Direstore!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Merestore Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("UPDATE pelanggan SET status ='1' where id ='" & Label6.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Merestore Data!")
                Else
                    table2()
                    MsgBox("Transaksi Berhasil Direstore!")
                    Label6.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If Label6.Text = "0" Then
            MsgBox("Pilih Log Yang Akan Dihapus!")
        Else
            Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Data Ini?", "Konfirmasi", MessageBoxButtons.YesNo)

            If result = DialogResult.No Then

            ElseIf result = DialogResult.Yes Then
                If CmdSQL("Delete from pelanggan where id ='" & Label6.Text & "'") = False Then
                    MsgBox("Terjadi Kesalahan Saat Menghapus Data!")
                Else
                    table2()
                    MsgBox("Transaksi Berhasil Dihapus!")
                    Label6.Text = "0"
                End If
            End If
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            If DataGridView3.RowCount < 2 Then
                MsgBox("Data Kosong, Tidak Dapat Dihapus!")
            Else
                Dim result = MessageBox.Show("Anda Yakin Ingin Menghapus Permanen Semua Data?", "Konfirmasi", MessageBoxButtons.YesNo)

                If result = DialogResult.No Then

                ElseIf result = DialogResult.Yes Then
                    If CmdSQL("Delete from pelanggan where status ='2'") = False Then
                        MsgBox("Terjadi Kesalahan Saat Menghapus Semua Data!")
                    Else
                        table2()
                        MsgBox("Transaksi Berhasil Dihapus Semua!")

                    End If
                End If
            End If
        Catch ex As NpgsqlException
            MsgBox(ex.Message)
        End Try
    End Sub


    Private da As NpgsqlDataAdapter
    Private ds As DataSet
    Private dtSource As DataTable
    Private PageCount As Integer
    Private maxRec As Integer
    Private pageSize As Integer
    Private currentPage As Integer
    Private recNo As Integer

    Dim sSql As String

    Private da2 As NpgsqlDataAdapter
    Private ds2 As DataSet
    Private dtSource2 As DataTable
    Private PageCount2 As Integer
    Private maxRec2 As Integer
    Private pageSize2 As Integer
    Private currentPage2 As Integer
    Private recNo2 As Integer

    Dim sSql2 As String

    Private da3 As NpgsqlDataAdapter
    Private ds3 As DataSet
    Private dtSource3 As DataTable
    Private PageCount3 As Integer
    Private maxRec3 As Integer
    Private pageSize3 As Integer
    Private currentPage3 As Integer
    Private recNo3 As Integer

    Dim sSql3 As String

    Private da4 As NpgsqlDataAdapter
    Private ds4 As DataSet
    Private dtSource4 As DataTable
    Private PageCount4 As Integer
    Private maxRec4 As Integer
    Private pageSize4 As Integer
    Private currentPage4 As Integer
    Private recNo4 As Integer

    Dim sSql4 As String


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

    Sub LoadDS3(ByVal sSQL As String)
        Try
            Dim cnString As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
            Dim conn As NpgsqlConnection = New NpgsqlConnection(cnString)

            da3 = New NpgsqlDataAdapter(sSQL, conn)
            ds3 = New DataSet()

            da3.Fill(ds3, "Items")
            dtSource3 = ds3.Tables("Items")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub LoadDS4(ByVal sSQL As String)
        Try
            Dim cnString As String = String.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "localhost", "5432", "postgres", "admin", "Kios_UNI")
            Dim conn As NpgsqlConnection = New NpgsqlConnection(cnString)

            da4 = New NpgsqlDataAdapter(sSQL, conn)
            ds4 = New DataSet()

            da4.Fill(ds4, "Items")
            dtSource4 = ds4.Tables("Items")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub DisplayPageInfo()
        Label17.Text = currentPage.ToString & "/ " & PageCount.ToString
    End Sub

    Sub DisplayPageInfo2()
        Label5.Text = currentPage2.ToString & "/ " & PageCount2.ToString
    End Sub

    Sub DisplayPageInfo3()
        Label1.Text = currentPage3.ToString & "/ " & PageCount3.ToString
    End Sub

    Sub DisplayPageInfo4()
        Label2.Text = currentPage4.ToString & "/ " & PageCount4.ToString
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
        DataGridView4.DataSource = dtTemp
        DisplayPageInfo()
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
        DataGridView3.DataSource = dtTemp2
        DisplayPageInfo2()
    End Sub

    Sub LoadPage3()
        Dim i As Integer
        Dim startRec3 As Integer
        Dim endRec3 As Integer
        Dim dtTemp3 As DataTable

        dtTemp3 = dtSource3.Clone

        If currentPage3 = PageCount3 Then
            endRec3 = maxRec3
        Else
            endRec3 = pageSize3 * currentPage3
        End If

        startRec3 = recNo3

        If dtSource3.Rows.Count > 0 Then
            For i = startRec3 To endRec3 - 1
                dtTemp3.ImportRow(dtSource3.Rows(i))
                recNo3 = recNo3 + 1
            Next
        End If
        DataGridView1.DataSource = dtTemp3
        DisplayPageInfo3()
    End Sub

    Sub LoadPage4()
        Dim i As Integer
        Dim startRec4 As Integer
        Dim endRec4 As Integer
        Dim dtTemp4 As DataTable

        dtTemp4 = dtSource4.Clone

        If currentPage4 = PageCount4 Then
            endRec4 = maxRec4
        Else
            endRec4 = pageSize4 * currentPage4
        End If

        startRec4 = recNo4

        If dtSource4.Rows.Count > 0 Then
            For i = startRec4 To endRec4 - 1
                dtTemp4.ImportRow(dtSource4.Rows(i))
                recNo4 = recNo4 + 1
            Next
        End If
        DataGridView2.DataSource = dtTemp4
        DisplayPageInfo4()
    End Sub

    Sub FillGrid()
        pageSize = 6
        maxRec = dtSource.Rows.Count
        PageCount = maxRec \ pageSize

        If (maxRec Mod pageSize) > 0 Then
            PageCount = PageCount + 1
        End If

        currentPage = 1
        recNo = 0

        LoadPage()
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

    Sub FillGrid3()
        pageSize3 = 6
        maxRec3 = dtSource3.Rows.Count
        PageCount3 = maxRec3 \ pageSize3

        If (maxRec3 Mod pageSize3) > 0 Then
            PageCount3 = PageCount3 + 1
        End If

        currentPage3 = 1
        recNo3 = 0

        LoadPage3()
    End Sub

    Sub FillGrid4()
        pageSize4 = 6
        maxRec4 = dtSource4.Rows.Count
        PageCount4 = maxRec4 \ pageSize4

        If (maxRec4 Mod pageSize4) > 0 Then
            PageCount4 = PageCount4 + 1
        End If

        currentPage4 = 1
        recNo4 = 0

        LoadPage4()
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

    Private Function CheckFillButton2() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize2 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            CheckFillButton2 = False
        Else
            CheckFillButton2 = True
        End If
    End Function

    Private Function CheckFillButton3() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize3 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            CheckFillButton3 = False
        Else
            CheckFillButton3 = True
        End If
    End Function

    Private Function CheckFillButton4() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize4 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            CheckFillButton4 = False
        Else
            CheckFillButton4 = True
        End If
    End Function

    Sub table1()
        sSql = "select * from pelanggan where status = '1' order by nama desc"
        LoadDS(sSql)
        FillGrid()
    End Sub

    Sub table2()
        sSql2 = "select * from pelanggan where status = '2' order by nama desc"
        LoadDS2(sSql2)
        FillGrid2()
    End Sub

    Sub table3()
        sSql3 = "select * from harian where status = '2' order by no desc"
        LoadDS3(sSql3)
        FillGrid3()
    End Sub

    Sub table4()
        sSql4 = "select * from buku where status = '2' order by no desc"
        LoadDS4(sSql4)
        FillGrid4()
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

    Private Sub btnNext3_Click(sender As Object, e As EventArgs) Handles btnNext3.Click
        'If the user did not click the "Fill Grid" button then Return
        If Not CheckFillButton3() Then Return
        'Check if the user clicked the "Fill Grid" button.
        If pageSize3 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            Return
        End If

        currentPage3 = currentPage3 + 1

        If currentPage3 > PageCount3 Then
            currentPage3 = PageCount3

            'Check if you are already at the last page.
            If recNo3 = maxRec3 Then
                MessageBox.Show("You are at the Last Page!")
                Return
            End If
        End If

        LoadPage3()
    End Sub

    Private Sub btnNext4_Click(sender As Object, e As EventArgs) Handles btnNext4.Click
        'If the user did not click the "Fill Grid" button then Return
        If Not CheckFillButton4() Then Return
        'Check if the user clicked the "Fill Grid" button.
        If pageSize4 = 0 Then
            MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            Return
        End If

        currentPage4 = currentPage4 + 1

        If currentPage4 > PageCount4 Then
            currentPage4 = PageCount4

            'Check if you are already at the last page.
            If recNo4 = maxRec4 Then
                MessageBox.Show("You are at the Last Page!")
                Return
            End If
        End If

        LoadPage4()
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

    Private Sub btnPrev3_Click(sender As Object, e As EventArgs) Handles btnPrev3.Click
        If Not CheckFillButton3() Then Return

        currentPage3 = currentPage3 - 1

        'Check if you are already at the first page.
        If currentPage3 < 1 Then
            MessageBox.Show("You are at the First Page!")
            currentPage3 = 1
            Return
        Else
            recNo3 = pageSize3 * (currentPage3 - 1)
        End If

        LoadPage3()
    End Sub

    Private Sub btnPrev4_Click(sender As Object, e As EventArgs) Handles btnPrev4.Click
        If Not CheckFillButton4() Then Return

        currentPage4 = currentPage4 - 1

        'Check if you are already at the first page.
        If currentPage4 < 1 Then
            MessageBox.Show("You are at the First Page!")
            currentPage4 = 1
            Return
        Else
            recNo4 = pageSize4 * (currentPage4 - 1)
        End If

        LoadPage4()
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

    Private Sub btnFirst3_Click(sender As Object, e As EventArgs) Handles btnFirst3.Click
        If Not CheckFillButton3() Then Return

        ' Check if you are already at the first page.
        If currentPage3 = 1 Then
            MessageBox.Show("You are at the First Page!")
            Return
        End If

        currentPage3 = 1
        recNo3 = 0

        LoadPage3()
    End Sub

    Private Sub btnFirst4_Click(sender As Object, e As EventArgs) Handles btnFirst4.Click
        If Not CheckFillButton4() Then Return

        ' Check if you are already at the first page.
        If currentPage4 = 1 Then
            MessageBox.Show("You are at the First Page!")
            Return
        End If

        currentPage4 = 1
        recNo4 = 0

        LoadPage4()
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

    Private Sub btnLast3_Click(sender As Object, e As EventArgs) Handles btnLast3.Click
        If Not CheckFillButton3() Then Return

        ' Check if you are already at the last page.
        If recNo3 = maxRec3 Then
            MessageBox.Show("You are at the Last Page!")
            Return
        End If

        currentPage3 = PageCount3

        recNo3 = pageSize3 * (currentPage3 - 1)

        LoadPage3()
    End Sub

    Private Sub btnLast4_Click(sender As Object, e As EventArgs) Handles btnLast4.Click
        If Not CheckFillButton4() Then Return

        ' Check if you are already at the last page.
        If recNo4 = maxRec4 Then
            MessageBox.Show("You are at the Last Page!")
            Return
        End If

        currentPage4 = PageCount4

        recNo4 = pageSize4 * (currentPage4 - 1)

        LoadPage4()
    End Sub
End Class
Imports System.Data.SQLite
Imports System.IO
Imports Microsoft.VisualBasic.ControlChars
Imports ClosedXML.Excel

Public Class Form1

    ' Database connection
    Private dbPath As String = Application.StartupPath & "\Database\KhenThuong.db"

    Private currentRecordId As Integer = -1

    Private searchResults As List(Of Integer) = New List(Of Integer)()

    Private mainToolTip As New ToolTip()


    ' Initialize form
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            ' Initialize database
            InitializeDatabase()

            ' Load data into comboboxes
            LoadComboBoxData()

            ' Load years
            LoadYears()

            ' Setup ListView columns
            SetupListView()

            ' Setup keyboard shortcuts
            SetupKeyboardShortcuts()

            ' Setup event handlers
            SetupEventHandlers()

            ' Make txtSTT readonly
            txtSTT.ReadOnly = True

            ' Load all records
            LoadAllRecords()

            With mainToolTip

                .AutoPopDelay = 5000
                .InitialDelay = 1000
                .ReshowDelay = 500
                .ShowAlways = True

                ' Tooltip controls
                .SetToolTip(txtHoVaTen, "Nhập họ tên đầy đủ của cá nhân được khen thưởng (vận dụng để ghi tên cho tổ chức được khen thưởng)!")
                .SetToolTip(cbLoaiKhenThuong, "Chọn loại khen thưởng (Chính quyền - Đảng - Đoàn...")
                .SetToolTip(cbNam, "Năm ban hành quyết định")
                .SetToolTip(btnCapNhat, "Lưu thông tin hiện tại vào cơ sở dữ liệu (Ctrl + S)")
                .SetToolTip(btnXuatFileExcel, "Xuất danh sách ra file Excel")
                .SetToolTip(btnNhapExcel, "Nhập danh sách từ file Excel")
                .SetToolTip(btnTimKiem, "Tìm kiếm theo thông tin nhập")
                .SetToolTip(lvKQTK, "Danh sách kết quả")
                .SetToolTip(btnLamMoi, "Làm mới biểu mẫu")
                .SetToolTip(btnXuatDB, "Xuất database ra file .db")
                .SetToolTip(btnNhapDB, "Nhập database từ file .db")
                .SetToolTip(btnXoaBanGhi, "Xóa bản ghi được chọn")
                .SetToolTip(txtGhiChu, "Nhập thông tin bổ sung nằm ngoài mẫu")
                .SetToolTip(txtQueQuan, "Quê quán của cá nhân được khen thưởng")
                .SetToolTip(txtNgaySinh, "Định dạng dd/MM/yyyy")
                .SetToolTip(txtNgayVaoDang, "Định dạng dd/MM/yyyy")
                .SetToolTip(txtCCCD, "Nhập đủ 12 số của TĐV hoặc CCCD")
                .SetToolTip(btnThemMoi, "Thêm mới bản ghi (Ctrl + N)")
                .SetToolTip(cbCapBac, "Cấp bậc của cá nhân được khen thưởng")
                .SetToolTip(cbChucVu, "Chức vụ của cá nhân được khen thưởng")
                .SetToolTip(cbDonVi, "Đơn vị của cá nhân được khen thưởng")
                .SetToolTip(cbHinhThucKhenThuong, "Hình thức khen thưởng mà cá nhân, tổ chức được nhận")
                .SetToolTip(cbLyDoKhenThuong, "Lý do cá nhân, tổ chức được khen thưởng")
                .SetToolTip(cbCapKyQuyetDinh, "Tổ chức nào ký quyết định khen thưởng")
                .SetToolTip(cbNguoiKyQuyetDinh, "Cá nhân nào đại diện cho tổ chức ký quyết định khen thưởng")
                .SetToolTip(txtSoQuyetDinh, "Số quyết định và ngày ký quyết định khen thưởng")

            End With

        Catch ex As Exception

            MessageBox.Show("Lỗi khi khởi tạo form: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

        ' Chữ chạy
        scrollText = lbThongTin.Text & "      "
        lbThongTin.Text = scrollText

        Timer1.Interval = 150
        Timer1.Start()

        ' Chuyển màu cầu vồng
        Timer2.Interval = 35
        Timer2.Start()

    End Sub

    ' Load all records
    Private Sub LoadAllRecords()

        Try

            lvKQTK.Items.Clear()
            searchResults.Clear()

            Using conn As New SQLiteConnection(
            "Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                Dim cmd As New SQLiteCommand("SELECT ID, Nam, LoaiKhenThuong, " &
                "HoVaTen, NgaySinh, " &
                "HinhThucKhenThuong, GhiChu " &
                "FROM KhenThuong " &
                "ORDER BY ID", conn)

                Dim reader As SQLiteDataReader =
                cmd.ExecuteReader()

                Dim stt As Integer = 1

                While reader.Read()

                    Dim item As New ListViewItem(stt.ToString())

                    item.SubItems.Add(reader("Nam").ToString())

                    item.SubItems.Add(reader("LoaiKhenThuong").ToString())

                    item.SubItems.Add(reader("HoVaTen").ToString())

                    item.SubItems.Add(If(IsDBNull(reader("NgaySinh")), "", reader("NgaySinh").ToString()))

                    item.SubItems.Add(reader("HinhThucKhenThuong").ToString())

                    item.SubItems.Add(If(IsDBNull(reader("GhiChu")), "", reader("GhiChu").ToString()))

                    item.Tag = reader("ID")

                    lvKQTK.Items.Add(item)

                    searchResults.Add(
                    CInt(reader("ID")))

                    stt += 1

                End While

                reader.Close()

                lbKQTK.Text =
                lvKQTK.Items.Count.ToString()

                conn.Close()

            End Using

        Catch ex As Exception

            MessageBox.Show("Lỗi khi tải danh sách: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    ' Initialize SQLite database
    Private Sub InitializeDatabase()

        Dim dbFolder = Path.GetDirectoryName(dbPath)

        If Not Directory.Exists(dbFolder) Then
            Directory.CreateDirectory(dbFolder)
        End If

        If Not File.Exists(dbPath) Then
            SQLiteConnection.CreateFile(dbPath)
        End If

        Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

            conn.Open()

            ' Create KhenThuong table
            Dim cmdKhenThuong As New SQLiteCommand("CREATE TABLE IF NOT EXISTS KhenThuong (" & "ID INTEGER PRIMARY KEY AUTOINCREMENT, " & "Nam INTEGER NOT NULL, " &
                "LoaiKhenThuong TEXT NOT NULL, " &
                "HoVaTen TEXT NOT NULL, " &
                "NgaySinh TEXT, " &
                "CapBac TEXT, " &
                "ChucVu TEXT, " &
                "DonVi TEXT, " &
                "HinhThucKhenThuong TEXT, " &
                "LyDoKhenThuong TEXT, " &
                "CapKyQuyetDinh TEXT, " &
                "NguoiKyQuyetDinh TEXT, " &
                "SoQuyetDinh TEXT, " &
                "GhiChu TEXT, " &
                "QueQuan TEXT, " &
                "NgayVaoDang TEXT, " &
                "CCCD TEXT)", conn)

            cmdKhenThuong.ExecuteNonQuery()

            ' Create DanhMuc table
            Dim cmdDanhMuc As New SQLiteCommand(
                "CREATE TABLE IF NOT EXISTS DanhMuc (" &
                "ID INTEGER PRIMARY KEY AUTOINCREMENT, " &
                "Loai TEXT NOT NULL, " &
                "GiaTri TEXT NOT NULL, " &
                "IsDeleted INTEGER DEFAULT 0)", conn)

            cmdDanhMuc.ExecuteNonQuery()

            conn.Close()

        End Using

    End Sub
    ' Setup ListView columns
    Private Sub SetupListView()

        lvKQTK.View = View.Details

        lvKQTK.Columns.Clear()

        lvKQTK.Columns.Add("STT", 50)
        lvKQTK.Columns.Add("Năm khen thưởng", 160)
        lvKQTK.Columns.Add("Loại khen thưởng", 200)
        lvKQTK.Columns.Add("Họ và tên", 180)
        lvKQTK.Columns.Add("Ngày sinh", 100)
        lvKQTK.Columns.Add("Hình thức khen thưởng", 190)
        lvKQTK.Columns.Add("Ghi chú", 200)

        lvKQTK.FullRowSelect = True
        lvKQTK.GridLines = True

    End Sub


    ' Setup keyboard shortcuts
    Private Sub SetupKeyboardShortcuts()

        Me.KeyPreview = True

    End Sub


    ' Setup event handlers
    Private Sub SetupEventHandlers()

        ' DatePicker events
        AddHandler dtpNgaySinh.ValueChanged,
            Sub(sender, e)
                txtNgaySinh.Text =
                    dtpNgaySinh.Value.ToString("dd/MM/yyyy")
            End Sub

        AddHandler dtpNgayVaoDang.ValueChanged,
            Sub(sender, e)
                txtNgayVaoDang.Text =
                    dtpNgayVaoDang.Value.ToString("dd/MM/yyyy")
            End Sub


        ' ListView selection
        AddHandler lvKQTK.SelectedIndexChanged,
            Sub(sender, e)

                If lvKQTK.SelectedItems.Count > 0 Then

                    Dim recordId As Integer =
                        CInt(lvKQTK.SelectedItems(0).Tag)

                    LoadRecordData(recordId)

                End If

            End Sub


        ' Format họ tên
        AddHandler txtHoVaTen.Leave,
            Sub(sender, e)

                txtHoVaTen.Text =
                    FormatName(txtHoVaTen.Text)

            End Sub


        ' Validate CCCD
        AddHandler txtCCCD.Leave,
            Sub(sender, e)

                ValidateCCCD()

            End Sub


        ' ComboBox xóa mềm
        Dim categoryCBs As ComboBox() = {cbLoaiKhenThuong, cbCapBac, cbChucVu, cbDonVi, cbHinhThucKhenThuong, cbLyDoKhenThuong, cbCapKyQuyetDinh, cbNguoiKyQuyetDinh}

        For Each cb In categoryCBs

            AddHandler cb.KeyDown,
                AddressOf CategoryComboBox_KeyDown

        Next


        ' Auto save combobox
        Dim allCBs As ComboBox() = {cbLoaiKhenThuong, cbCapBac, cbChucVu, cbDonVi, cbHinhThucKhenThuong, cbLyDoKhenThuong, cbCapKyQuyetDinh, cbNguoiKyQuyetDinh}

        For Each cb In allCBs

            AddHandler cb.Leave,
                AddressOf AutoSaveComboBox

        Next

    End Sub
    ' Load record data into form
    Private Sub LoadRecordData(recordId As Integer)

        Try

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                Dim cmd As New SQLiteCommand("SELECT * FROM KhenThuong WHERE ID = @ID", conn)

                cmd.Parameters.AddWithValue("@ID", recordId)

                Dim reader As SQLiteDataReader = cmd.ExecuteReader()

                If reader.Read() Then

                    currentRecordId = recordId

                    txtSTT.Text = reader("ID").ToString()
                    cbNam.Text = reader("Nam").ToString()
                    cbLoaiKhenThuong.Text = reader("LoaiKhenThuong").ToString()

                    txtHoVaTen.Text = reader("HoVaTen").ToString()

                    txtNgaySinh.Text = If(IsDBNull(reader("NgaySinh")), "", reader("NgaySinh").ToString())

                    cbCapBac.Text = If(IsDBNull(reader("CapBac")), "", reader("CapBac").ToString())

                    cbChucVu.Text = If(IsDBNull(reader("ChucVu")), "", reader("ChucVu").ToString())

                    cbDonVi.Text = If(IsDBNull(reader("DonVi")), "", reader("DonVi").ToString())

                    cbHinhThucKhenThuong.Text = If(IsDBNull(reader("HinhThucKhenThuong")), "", reader("HinhThucKhenThuong").ToString())

                    cbLyDoKhenThuong.Text = If(IsDBNull(reader("LyDoKhenThuong")), "", reader("LyDoKhenThuong").ToString())

                    cbCapKyQuyetDinh.Text = If(IsDBNull(reader("CapKyQuyetDinh")), "", reader("CapKyQuyetDinh").ToString())

                    cbNguoiKyQuyetDinh.Text = If(IsDBNull(reader("NguoiKyQuyetDinh")), "", reader("NguoiKyQuyetDinh").ToString())

                    txtSoQuyetDinh.Text = If(IsDBNull(reader("SoQuyetDinh")), "", reader("SoQuyetDinh").ToString())

                    txtQueQuan.Text = If(IsDBNull(reader("QueQuan")), "", reader("QueQuan").ToString())

                    txtGhiChu.Text = If(IsDBNull(reader("GhiChu")), "", reader("GhiChu").ToString())

                    txtNgayVaoDang.Text = If(IsDBNull(reader("NgayVaoDang")), "", reader("NgayVaoDang").ToString())

                    txtCCCD.Text = If(IsDBNull(reader("CCCD")), "", reader("CCCD").ToString())

                End If

                reader.Close()

            End Using

        Catch ex As Exception

            MessageBox.Show("Lỗi khi tải dữ liệu bản ghi: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    ' Load years into cbNam
    Private Sub LoadYears()

        cbNam.Items.Clear()

        Dim currentYear As Integer = DateTime.Now.Year - 14

        For iYear As Integer = currentYear To currentYear + 14

            cbNam.Items.Add(iYear)

        Next

        cbNam.SelectedIndex = 4

    End Sub


    ' Load combobox data from database
    Private Sub LoadComboBoxData()

        Try

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                ' Load data
                LoadComboBoxFromDB(conn, cbLoaiKhenThuong, "LoaiKhenThuong")
                LoadComboBoxFromDB(conn, cbCapBac, "CapBac")
                LoadComboBoxFromDB(conn, cbChucVu, "ChucVu")
                LoadComboBoxFromDB(conn, cbDonVi, "DonVi")
                LoadComboBoxFromDB(conn, cbHinhThucKhenThuong, "HinhThucKhenThuong")
                LoadComboBoxFromDB(conn, cbLyDoKhenThuong, "LyDoKhenThuong")
                LoadComboBoxFromDB(conn, cbCapKyQuyetDinh, "CapKyQuyetDinh")
                LoadComboBoxFromDB(conn, cbNguoiKyQuyetDinh, "NguoiKyQuyetDinh")

                conn.Close()

            End Using

        Catch ex As Exception

            MessageBox.Show("Lỗi khi tải dữ liệu danh mục: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub


    ' Load combobox from database
    Private Sub LoadComboBoxFromDB(conn As SQLiteConnection, cb As ComboBox, categoryName As String)

        cb.Items.Clear()

        Dim cmd As New SQLiteCommand("SELECT DISTINCT GiaTri " & "FROM DanhMuc " & "WHERE Loai = @Loai " & "AND IsDeleted = 0 " & "ORDER BY GiaTri", conn)

        cmd.Parameters.AddWithValue("@Loai", categoryName)

        Dim reader As SQLiteDataReader = cmd.ExecuteReader()

        While reader.Read()

            cb.Items.Add(reader("GiaTri").ToString())

        End While

        reader.Close()

    End Sub


    ' Soft delete category
    Public Sub SoftDeleteCategory(cb As ComboBox)

        Dim selectedValue As String = cb.Text

        ' Xác định loại danh mục
        Dim categoryType As String = ""

        Select Case cb.Name

            Case "cbLoaiKhenThuong"
                categoryType = "LoaiKhenThuong"

            Case "cbCapBac"
                categoryType = "CapBac"

            Case "cbChucVu"
                categoryType = "ChucVu"

            Case "cbDonVi"
                categoryType = "DonVi"

            Case "cbHinhThucKhenThuong"
                categoryType = "HinhThucKhenThuong"

            Case "cbLyDoKhenThuong"
                categoryType = "LyDoKhenThuong"

            Case "cbCapKyQuyetDinh"
                categoryType = "CapKyQuyetDinh"

        End Select


        If Not String.IsNullOrEmpty(categoryType) Then

            Dim msg As String = String.Format("Xóa mục '{0}' khỏi danh sách gợi ý?", selectedValue)

            If MessageBox.Show(
                msg,
                "Xác nhận xóa mềm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) = DialogResult.Yes Then

                Try

                    Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                        conn.Open()

                        Using cmd As New SQLiteCommand("UPDATE DanhMuc " & "SET IsDeleted = 1 " & "WHERE Loai = @Loai " & "AND GiaTri = @GiaTri", conn)

                            cmd.Parameters.AddWithValue("@Loai", categoryType)
                            cmd.Parameters.AddWithValue("@GiaTri", selectedValue)

                            cmd.ExecuteNonQuery()

                        End Using

                    End Using


                    ' Reload
                    LoadComboBoxData()

                    cb.SelectedIndex = -1
                    cb.Text = ""

                Catch ex As Exception

                    MessageBox.Show("Lỗi: " & ex.Message)

                End Try

            End If

        End If

    End Sub


    ' Ctrl + -
    Private Sub CategoryComboBox_KeyDown(
        sender As Object,
        e As KeyEventArgs)

        If e.Control AndAlso
           (e.KeyCode = Keys.OemMinus OrElse
            e.KeyCode = Keys.Subtract) Then

            Dim cb As ComboBox =
                DirectCast(sender, ComboBox)

            If Not String.IsNullOrEmpty(cb.Text) Then

                SoftDeleteCategory(cb)

                e.SuppressKeyPress = True

            End If

        End If

    End Sub
    ' Format name
    Private Function FormatName(name As String) As String

        If String.IsNullOrEmpty(name) Then
            Return ""
        End If

        Dim words As String() = name.Split(" "c)

        Dim formattedWords As New List(Of String)

        For Each word In words

            If word.Length > 0 Then

                formattedWords.Add(
                    Char.ToUpper(word(0)) &
                    word.Substring(1).ToLower())

            End If

        Next

        Return String.Join(" ", formattedWords)

    End Function


    ' Xóa dấu tiếng Việt
    Private Function RemoveVietnameseSigns(text As String) As String

        If String.IsNullOrEmpty(text) Then
            Return ""
        End If

        Dim arr1 As String() = {"á", "à", "ả", "ã", "ạ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "đ", "é", "è", "ẻ", "ẽ", "ẹ", "ê", "ế", "ề", "ể", "ễ", "ệ", "í", "ì", "ỉ", "ĩ", "ị", "ó", "ò", "ỏ", "õ", "ọ", "ô", "ố", "ồ", "ổ", "ỗ", "ộ", "ơ", "ớ", "ờ", "ở", "ỡ", "ợ", "ú", "ù", "ủ", "ũ", "ụ", "ư", "ứ", "ừ", "ử", "ữ", "ự", "ý", "ỳ", "ỷ", "ỹ", "ỵ"}

        Dim arr2 As String() = {"a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "d", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "e", "i", "i", "i", "i", "i", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "o", "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", "u", "y", "y", "y", "y", "y"}

        text = text.ToLower()

        For i As Integer = 0 To arr1.Length - 1

            text = text.Replace(arr1(i), arr2(i))

        Next

        Return text

    End Function


    ' Validate CCCD
    Private Sub ValidateCCCD()

        If txtCCCD.Text.Length > 0 AndAlso
           txtCCCD.Text.Length <> 12 Then

            MessageBox.Show(
                "CCCD phải có đúng 12 ký tự!",
                "Cảnh báo",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)

            txtCCCD.Focus()
            txtCCCD.SelectAll()

        End If

    End Sub


    ' Auto save combobox
    Private Sub AutoSaveComboBox(sender As Object, e As EventArgs)

        Dim cb As ComboBox = DirectCast(sender, ComboBox)

        Dim value As String = cb.Text.Trim()

        If value = "" Then Exit Sub

        Dim loai As String = cb.Name.Replace("cb", "")

        AddCategoryIfNotExists(loai, value)

        ' Reload combobox
        LoadComboBoxData()

        ' Giữ giá trị
        cb.Text = value

        ' Reload current record
        If currentRecordId > 0 Then

            LoadRecordData(currentRecordId)

        End If

    End Sub


    ' Add category if not exists
    Private Sub AddCategoryIfNotExists(categoryName As String, categoryValue As String)

        If String.IsNullOrWhiteSpace(categoryValue) Then
            Return
        End If

        Try

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                ' ===== KIỂM TRA ĐÃ TỒN TẠI =====

                Dim checkCmd As New SQLiteCommand("SELECT ID, IsDeleted " & "FROM DanhMuc " & "WHERE Loai = @Loai " & "AND GiaTri = @GiaTri", conn)

                checkCmd.Parameters.AddWithValue("@Loai", categoryName)

                checkCmd.Parameters.AddWithValue("@GiaTri", categoryValue)

                Dim reader As SQLiteDataReader = checkCmd.ExecuteReader()

                If reader.Read() Then

                    Dim id As Integer =
                    Convert.ToInt32(reader("ID"))

                    Dim isDeleted As Integer =
                    Convert.ToInt32(reader("IsDeleted"))

                    reader.Close()

                    ' ===== NẾU ĐÃ XÓA MỀM → KHÔI PHỤC =====

                    If isDeleted = 1 Then

                        Dim restoreCmd As New SQLiteCommand("UPDATE DanhMuc " & "SET IsDeleted = 0 " & "WHERE ID = @ID", conn)

                        restoreCmd.Parameters.AddWithValue("@ID", id)

                        restoreCmd.ExecuteNonQuery()

                        restoreCmd.Dispose()

                    End If

                Else

                    reader.Close()

                    ' ===== THÊM MỚI =====

                    Dim insertCmd As New SQLiteCommand("INSERT INTO DanhMuc " & "(Loai, GiaTri, IsDeleted) " & "VALUES (@Loai, @GiaTri, 0)", conn)

                    insertCmd.Parameters.AddWithValue("@Loai", categoryName)

                    insertCmd.Parameters.AddWithValue("@GiaTri", categoryValue)

                    insertCmd.ExecuteNonQuery()

                    insertCmd.Dispose()

                End If

                checkCmd.Dispose()

                conn.Close()

            End Using

        Catch ex As Exception

            MessageBox.Show(
            "Lỗi khi thêm danh mục: " & ex.Message,
            "Lỗi",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)

        End Try

    End Sub
    ' Button: Xuất DB
    Private Sub btnXuatDB_Click(sender As Object, e As EventArgs) Handles btnXuatDB.Click
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Database Files (*.db)|*.db"
            saveDialog.InitialDirectory = Application.StartupPath

            If saveDialog.ShowDialog() = DialogResult.OK Then
                File.Copy(dbPath, saveDialog.FileName, True)
                MessageBox.Show("Xuất database thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show("Lỗi khi xuất database: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Button: Nhập DB
    Private Sub btnNhapDB_Click(sender As Object, e As EventArgs) Handles btnNhapDB.Click
        Try
            Dim openDialog As New OpenFileDialog()
            openDialog.Filter = "Database Files (*.db)|*.db"
            openDialog.InitialDirectory = Application.StartupPath

            If openDialog.ShowDialog() = DialogResult.OK Then
                ImportFromDB(openDialog.FileName)
            End If

        Catch ex As Exception
            MessageBox.Show("Lỗi khi nhập database: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Import from DB
    Private Sub ImportFromDB(importDbPath As String)

        Try

            Dim duplicates As New List(Of String)

            Dim importedCount As Integer = 0

            Using importConn As New SQLiteConnection("Data Source=" & importDbPath & ";Version=3;")

                importConn.Open()

                Dim importCmd As New SQLiteCommand("SELECT * FROM KhenThuong", importConn)

                Dim reader As SQLiteDataReader = importCmd.ExecuteReader()

                Using mainConn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                    mainConn.Open()

                    While reader.Read()

                        Dim nam As Integer = If(reader("Nam") Is DBNull.Value, 0, Convert.ToInt32(reader("Nam")))

                        Dim loaiKhenThuong As String = If(reader("LoaiKhenThuong") Is DBNull.Value, "", reader("LoaiKhenThuong").ToString())

                        Dim hoVaTen As String = If(reader("HoVaTen") Is DBNull.Value, "", reader("HoVaTen").ToString())

                        Dim hoTenKhongDau As String = RemoveVietnameseSigns(hoVaTen).ToLower()

                        Dim cccd As String = If(reader("CCCD") Is DBNull.Value, "", reader("CCCD").ToString())

                        Dim ngaySinh As String = If(reader("NgaySinh") Is DBNull.Value, "", reader("NgaySinh").ToString())

                        Dim queQuan As String = If(reader("QueQuan") Is DBNull.Value, "", reader("QueQuan").ToString())


                        Dim soQuyetDinh As String = If(reader("SoQuyetDinh") Is DBNull.Value, "", reader("SoQuyetDinh").ToString())

                        ' Kiểm tra trùng
                        Dim checkCmd As New SQLiteCommand("SELECT ID FROM KhenThuong " & "WHERE Nam = @Nam " & "AND LoaiKhenThuong = @LoaiKhenThuong " & "AND HoVaTen = @HoVaTen " & "AND SoQuyetDinh = @SoQuyetDinh " & "AND (CCCD = @CCCD OR NgaySinh = @NgaySinh)", mainConn)

                        checkCmd.Parameters.AddWithValue("@SoQuyetDinh", soQuyetDinh)
                        checkCmd.Parameters.AddWithValue("@Nam", nam)

                        checkCmd.Parameters.AddWithValue("@LoaiKhenThuong", loaiKhenThuong)

                        checkCmd.Parameters.AddWithValue("@HoVaTen", hoVaTen)

                        checkCmd.Parameters.AddWithValue("@CCCD", cccd)

                        checkCmd.Parameters.AddWithValue("@NgaySinh", ngaySinh)

                        Dim existingId As Object = checkCmd.ExecuteScalar()

                        checkCmd.Dispose()

                        If existingId IsNot Nothing AndAlso
                       existingId IsNot DBNull.Value Then

                            duplicates.Add(hoVaTen & " - " & nam.ToString())

                        Else

                            ' Thêm bản ghi mới
                            Dim insertCmd As New SQLiteCommand(
                            "INSERT INTO KhenThuong (" &
                            "Nam, " &
                            "LoaiKhenThuong, " &
                            "HoVaTen, " &
                            "HoVaTenKhongDau, " &
                            "NgaySinh, " &
                            "CapBac, " &
                            "ChucVu, " &
                            "DonVi, " &
                            "HinhThucKhenThuong, " &
                            "LyDoKhenThuong, " &
                            "CapKyQuyetDinh, " &
                            "SoQuyetDinh, " &
                            "NguoiKyQuyetDinh, " &
                            "GhiChu, " &
                            "NgayVaoDang, " &
                            "QueQuan, " &
                            "CCCD) " &
                            "VALUES (" &
                            "@Nam, " &
                            "@LoaiKhenThuong, " &
                            "@HoVaTen, " &
                            "@HoVaTenKhongDau, " &
                            "@NgaySinh, " &
                            "@CapBac, " &
                            "@ChucVu, " &
                            "@DonVi, " &
                            "@HinhThucKhenThuong, " &
                            "@LyDoKhenThuong, " &
                            "@CapKyQuyetDinh, " &
                            "@SoQuyetDinh, " &
                            "@NguoiKyQuyetDinh, " &
                            "@GhiChu, " &
                            "@NgayVaoDang, " &
                            "@QueQuan, " &
                            "@CCCD)", mainConn)

                            insertCmd.Parameters.AddWithValue("@Nam", nam)

                            insertCmd.Parameters.AddWithValue("@LoaiKhenThuong", loaiKhenThuong)

                            insertCmd.Parameters.AddWithValue("@HoVaTen", hoVaTen)

                            insertCmd.Parameters.AddWithValue("@HoVaTenKhongDau", hoTenKhongDau)

                            insertCmd.Parameters.AddWithValue("@NgaySinh", If(String.IsNullOrWhiteSpace(ngaySinh), DBNull.Value, ngaySinh))

                            insertCmd.Parameters.AddWithValue("@CapBac", If(reader("CapBac") Is DBNull.Value, DBNull.Value, reader("CapBac").ToString()))

                            insertCmd.Parameters.AddWithValue("@ChucVu", If(reader("ChucVu") Is DBNull.Value, DBNull.Value, reader("ChucVu").ToString()))

                            insertCmd.Parameters.AddWithValue("@DonVi", If(reader("DonVi") Is DBNull.Value, DBNull.Value, reader("DonVi").ToString()))

                            insertCmd.Parameters.AddWithValue("@HinhThucKhenThuong", If(reader("HinhThucKhenThuong") Is DBNull.Value, DBNull.Value, reader("HinhThucKhenThuong").ToString()))

                            insertCmd.Parameters.AddWithValue("@LyDoKhenThuong", If(reader("LyDoKhenThuong") Is DBNull.Value, DBNull.Value, reader("LyDoKhenThuong").ToString()))

                            insertCmd.Parameters.AddWithValue("@CapKyQuyetDinh", If(reader("CapKyQuyetDinh") Is DBNull.Value, DBNull.Value, reader("CapKyQuyetDinh").ToString()))

                            insertCmd.Parameters.AddWithValue("@NguoiKyQuyetDinh", If(reader("NguoiKyQuyetDinh") Is DBNull.Value, DBNull.Value, reader("NguoiKyQuyetDinh").ToString()))

                            insertCmd.Parameters.AddWithValue("@SoQuyetDinh", If(reader("SoQuyetDinh") Is DBNull.Value, DBNull.Value, reader("SoQuyetDinh").ToString()))

                            insertCmd.Parameters.AddWithValue("@GhiChu", If(reader("GhiChu") Is DBNull.Value, DBNull.Value, reader("GhiChu").ToString()))

                            insertCmd.Parameters.AddWithValue("@NgayVaoDang", If(reader("NgayVaoDang") Is DBNull.Value, DBNull.Value, reader("NgayVaoDang").ToString()))

                            insertCmd.Parameters.AddWithValue("@CCCD", If(String.IsNullOrWhiteSpace(cccd), DBNull.Value, cccd))

                            insertCmd.Parameters.AddWithValue("@QueQuan", If(String.IsNullOrWhiteSpace(queQuan), DBNull.Value, queQuan))

                            insertCmd.ExecuteNonQuery()

                            insertCmd.Dispose()

                            importedCount += 1

                        End If

                    End While

                    mainConn.Close()

                End Using

                reader.Close()

                importCmd.Dispose()

                importConn.Close()

            End Using

            Dim message As String = "Đã nhập thành công " & importedCount & " bản ghi."

            If duplicates.Count > 0 Then

                message &= vbCrLf & "Bỏ qua " & duplicates.Count & " bản ghi trùng lặp."

            End If

            MessageBox.Show(message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

            LoadAllRecords()

            ClearForm()

        Catch ex As Exception

            MessageBox.Show("Lỗi nhập dữ liệu: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    ' Keyboard shortcuts
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean

        ' 1. Xử lý các tổ hợp phím có Ctrl
        If (keyData And Keys.Control) = Keys.Control Then

            Dim keyCode As Keys = keyData And Not Keys.Control

            Select Case keyCode

                Case Keys.N
                    btnThemMoi_Click(Nothing, Nothing)
                    Return True

                Case Keys.S
                    btnCapNhat_Click(Nothing, Nothing)
                    Return True

                Case Keys.OemMinus, Keys.Subtract ' Ctrl + -
                    If TypeOf ActiveControl Is ComboBox Then
                        SoftDeleteCategory(DirectCast(ActiveControl, ComboBox))
                        Return True
                    End If

            End Select

        End If

        ' 2. Xử lý các tổ hợp phím có Alt
        If (keyData And Keys.Alt) = Keys.Alt Then

            Dim keyCode As Keys = keyData And Not Keys.Alt

            If keyCode = Keys.X Then ' Alt + X
                Application.Exit()
                Return True
            End If

        End If

        ' 3. Xử lý các phím đơn lẻ (F1, F3, F5, Enter)
        Select Case keyData

            Case Keys.F1
                ShowHelpInfo()
                Return True

            Case Keys.F5
                btnLamMoi_Click(Nothing, Nothing)
                Return True

            Case Keys.F3, Keys.Enter

                ' Không thực hiện tìm kiếm nếu đang gõ trong ô Ghi chú đa dòng
                If Not (TypeOf ActiveControl Is TextBox AndAlso DirectCast(ActiveControl, TextBox).Multiline) Then
                    btnTimKiem_Click(Nothing, Nothing)
                    Return True
                End If

        End Select

        Return MyBase.ProcessCmdKey(msg, keyData)

    End Function

    ' Hiển thị hướng dẫn
    Private Sub ShowHelpInfo()

        Dim helpMsg As String =
            "PHẦN MỀM THEO DÕI KHEN THƯỞNG" & vbCrLf &
            "-----------------------------" & vbCrLf &
            "- Tác giả: Trần Duy Mạnh - SĐT: 098.110.7698" & vbCrLf &
            "---------------------------------------------------" & vbCrLf &
            "• F1: Xem hướng dẫn sử dụng" & vbCrLf &
            "• F3 hoặc Enter: Thực hiện tìm kiếm" & vbCrLf &
            "• F5: Làm mới form dữ liệu" & vbCrLf &
            "• Ctrl + N: Thêm bản ghi mới" & vbCrLf &
            "• Ctrl + S: Cập nhật dữ liệu" & vbCrLf &
            "• Ctrl + '-': Xóa mục trong danh sách gợi ý" & vbCrLf &
            "• Alt + X: Thoát phần mềm" & vbCrLf & vbCrLf &
            "Phiên bản: 1.0.1"

        MessageBox.Show(helpMsg, "Thông tin & Hướng dẫn", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    ' =========================
    ' HIỆU ỨNG CHẠY CHỮ CẦU VỒNG
    ' =========================

    Private scrollText As String
    Private hueValue As Integer = 0

    ' Chuyển HSV sang RGB
    Private Function HSVtoRGB(h As Integer, s As Double, v As Double) As Color

        Dim hi As Integer = CInt(Math.Floor(h / 60)) Mod 6
        Dim f As Double = h / 60.0 - Math.Floor(h / 60)

        v = v * 255

        Dim p As Integer = CInt(v * (1 - s))
        Dim q As Integer = CInt(v * (1 - f * s))
        Dim t As Integer = CInt(v * (1 - (1 - f) * s))
        Dim vi As Integer = CInt(v)

        Select Case hi

            Case 0
                Return Color.FromArgb(vi, t, p)

            Case 1
                Return Color.FromArgb(q, vi, p)

            Case 2
                Return Color.FromArgb(p, vi, t)

            Case 3
                Return Color.FromArgb(p, q, vi)

            Case 4
                Return Color.FromArgb(t, p, vi)

            Case Else
                Return Color.FromArgb(vi, p, q)

        End Select

    End Function

    ' Timer chạy chữ
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If scrollText.Length > 1 Then

            scrollText = scrollText.Substring(1) & scrollText(0)
            lbThongTin.Text = scrollText

        End If

    End Sub

    ' Timer đổi màu cầu vồng
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        lbThongTin.ForeColor = HSVtoRGB(hueValue, 1, 1)

        hueValue += 1

        If hueValue >= 360 Then
            hueValue = 0
        End If

    End Sub
    ' Clear form
    Private Sub ClearForm()

        txtSTT.Clear()

        cbNam.Text = ""
        cbLoaiKhenThuong.Text = ""

        txtHoVaTen.Clear()

        txtNgaySinh.Clear()

        cbCapBac.Text = ""
        cbChucVu.Text = ""
        cbDonVi.Text = ""

        cbHinhThucKhenThuong.Text = ""
        cbLyDoKhenThuong.Text = ""
        cbCapKyQuyetDinh.Text = ""
        cbNguoiKyQuyetDinh.Text = ""
        txtSoQuyetDinh.Clear()

        txtGhiChu.Clear()
        txtQueQuan.Clear()
        txtNgayVaoDang.Clear()

        txtCCCD.Clear()

    End Sub
    ' Button: Tìm kiếm
    Private Sub btnTimKiem_Click(sender As Object, e As EventArgs) Handles btnTimKiem.Click

        PerformSearch()

    End Sub
    ' Perform search
    Private Sub PerformSearch()

        Try

            lvKQTK.Items.Clear()
            searchResults.Clear()

            Dim whereClause As String = ""
            Dim parameters As New List(Of SQLiteParameter)()

            ' Năm
            If Not String.IsNullOrEmpty(cbNam.Text) Then

                whereClause &= "Nam = @Nam AND "

                parameters.Add(New SQLiteParameter("@Nam", cbNam.Text))

            End If

            ' Loại khen thưởng
            If Not String.IsNullOrEmpty(cbLoaiKhenThuong.Text) Then

                whereClause &= "LoaiKhenThuong LIKE @LoaiKhenThuong AND "

                parameters.Add(New SQLiteParameter("@LoaiKhenThuong", "%" & cbLoaiKhenThuong.Text & "%"))

            End If

            ' Ngày sinh
            If Not String.IsNullOrEmpty(txtNgaySinh.Text) Then

                whereClause &= "NgaySinh LIKE @NgaySinh AND "

                parameters.Add(New SQLiteParameter("@NgaySinh", "%" & txtNgaySinh.Text & "%"))

            End If

            ' Cấp bậc
            If Not String.IsNullOrEmpty(cbCapBac.Text) Then

                whereClause &= "CapBac LIKE @CapBac AND "

                parameters.Add(New SQLiteParameter("@CapBac", "%" & cbCapBac.Text & "%"))

            End If

            ' Chức vụ
            If Not String.IsNullOrEmpty(cbChucVu.Text) Then

                whereClause &= "ChucVu LIKE @ChucVu AND "

                parameters.Add(New SQLiteParameter("@ChucVu", "%" & cbChucVu.Text & "%"))

            End If

            ' Đơn vị
            If Not String.IsNullOrEmpty(cbDonVi.Text) Then

                whereClause &= "DonVi LIKE @DonVi AND "

                parameters.Add(New SQLiteParameter("@DonVi", "%" & cbDonVi.Text & "%"))

            End If

            ' Hình thức khen thưởng
            If Not String.IsNullOrEmpty(cbHinhThucKhenThuong.Text) Then

                whereClause &= "HinhThucKhenThuong LIKE @HinhThucKhenThuong AND "

                parameters.Add(New SQLiteParameter("@HinhThucKhenThuong", "%" & cbHinhThucKhenThuong.Text & "%"))

            End If

            ' Lý do khen thưởng
            If Not String.IsNullOrEmpty(cbLyDoKhenThuong.Text) Then

                whereClause &= "LyDoKhenThuong LIKE @LyDoKhenThuong AND "

                parameters.Add(New SQLiteParameter("@LyDoKhenThuong", "%" & cbLyDoKhenThuong.Text & "%"))

            End If

            ' Cấp ký quyết định
            If Not String.IsNullOrEmpty(cbCapKyQuyetDinh.Text) Then

                whereClause &= "CapKyQuyetDinh LIKE @CapKyQuyetDinh AND "

                parameters.Add(New SQLiteParameter("@CapKyQuyetDinh", "%" & cbCapKyQuyetDinh.Text & "%"))

            End If

            If Not String.IsNullOrEmpty(cbNguoiKyQuyetDinh.Text) Then

                whereClause &= "NguoiKyQuyetDinh LIKE @NguoiKyQuyetDinh AND "

                parameters.Add(New SQLiteParameter("@NguoiKyQuyetDinh", "%" & cbNguoiKyQuyetDinh.Text & "%"))

            End If

            ' Số quyết định
            If Not String.IsNullOrEmpty(txtSoQuyetDinh.Text) Then

                whereClause &= "SoQuyetDinh LIKE @SoQuyetDinh AND "

                parameters.Add(New SQLiteParameter("@SoQuyetDinh", "%" & txtSoQuyetDinh.Text & "%"))

            End If

            If Not String.IsNullOrEmpty(txtQueQuan.Text) Then

                whereClause &= "QueQuan LIKE @QueQuan AND "

                parameters.Add(New SQLiteParameter("@QueQuan", "%" & txtQueQuan.Text & "%"))

            End If

            If Not String.IsNullOrEmpty(txtGhiChu.Text) Then

                whereClause &= "GhiChu LIKE @GhiChu AND "

                parameters.Add(New SQLiteParameter("@GhiChu", "%" & txtGhiChu.Text & "%"))

            End If

            ' CCCD
            If Not String.IsNullOrEmpty(txtCCCD.Text) Then

                whereClause &= "CCCD LIKE @CCCD AND "

                parameters.Add(New SQLiteParameter("@CCCD", "%" & txtCCCD.Text & "%"))

            End If

            ' Xóa AND cuối
            If whereClause.EndsWith(" AND ") Then

                whereClause = whereClause.Substring(0, whereClause.Length - 5)

            End If

            ' Nếu không có điều kiện
            If String.IsNullOrEmpty(whereClause) _
                AndAlso String.IsNullOrEmpty(txtHoVaTen.Text) Then

                LoadAllRecords()
                Return

            End If

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")
                conn.Open()

                Dim query As String = "SELECT ID, Nam, LoaiKhenThuong, " &
                "HoVaTen, NgaySinh, " &
                "HinhThucKhenThuong, GhiChu " &
                "FROM KhenThuong "

                If whereClause <> "" Then
                    query &= "WHERE " & whereClause
                End If

                query &= " ORDER BY ID"

                Dim cmd As New SQLiteCommand(query, conn)

                For Each param In parameters
                    cmd.Parameters.Add(param)
                Next

                Dim reader As SQLiteDataReader = cmd.ExecuteReader()

                Dim stt As Integer = 1

                ' Tìm không dấu
                Dim keywordNoSign As String = RemoveVietnameseSigns(txtHoVaTen.Text.Trim().ToLower())

                While reader.Read()

                    Dim dbHoTen As String = reader("HoVaTen").ToString()

                    Dim dbHoTenNoSign As String = RemoveVietnameseSigns(dbHoTen.ToLower())

                    ' Lọc họ tên không dấu
                    If keywordNoSign <> "" Then

                        If Not dbHoTenNoSign.Contains(keywordNoSign) Then

                            Continue While

                        End If

                    End If

                    Dim item As New ListViewItem(stt.ToString())

                    item.SubItems.Add(reader("Nam").ToString())

                    item.SubItems.Add(reader("LoaiKhenThuong").ToString())

                    item.SubItems.Add(dbHoTen)

                    item.SubItems.Add(If(IsDBNull(reader("NgaySinh")), "", reader("NgaySinh").ToString()))

                    item.SubItems.Add(If(IsDBNull(reader("HinhThucKhenThuong")), "", reader("HinhThucKhenThuong").ToString()))

                    item.SubItems.Add(If(IsDBNull(reader("GhiChu")), "", reader("GhiChu").ToString()))

                    item.Tag = reader("ID")

                    lvKQTK.Items.Add(item)

                    searchResults.Add(CInt(reader("ID")))

                    stt += 1

                End While

                reader.Close()

                lbKQTK.Text = lvKQTK.Items.Count.ToString()

                conn.Close()

            End Using

        Catch ex As Exception

            MessageBox.Show("Lỗi khi tìm kiếm: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub btnLamMoi_Click(sender As Object, e As EventArgs) Handles btnLamMoi.Click

        ClearForm()

    End Sub
    Private Sub btnThemMoi_Click(sender As Object, e As EventArgs) Handles btnThemMoi.Click
        Try
            ' Validate required fields
            If String.IsNullOrEmpty(cbNam.Text) Then
                MessageBox.Show("Vui lòng chọn Năm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cbNam.Focus()
                Return
            End If

            If String.IsNullOrEmpty(cbLoaiKhenThuong.Text) Then
                MessageBox.Show("Vui lòng chọn Loại khen thưởng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cbLoaiKhenThuong.Focus()
                Return
            End If

            If String.IsNullOrEmpty(txtHoVaTen.Text) Then
                MessageBox.Show("Vui lòng nhập Họ và tên!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtHoVaTen.Focus()
                Return
            End If

            ' Validate CCCD
            If txtCCCD.Text.Length > 0 AndAlso txtCCCD.Text.Length <> 12 Then
                MessageBox.Show("CCCD phải có đúng 12 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtCCCD.Focus()
                Return
            End If

            ' Add categories if not exists
            AddCategoryIfNotExists("LoaiKhenThuong", cbLoaiKhenThuong.Text)
            AddCategoryIfNotExists("CapBac", cbCapBac.Text)
            AddCategoryIfNotExists("ChucVu", cbChucVu.Text)
            AddCategoryIfNotExists("DonVi", cbDonVi.Text)
            AddCategoryIfNotExists("HinhThucKhenThuong", cbHinhThucKhenThuong.Text)
            AddCategoryIfNotExists("LyDoKhenThuong", cbLyDoKhenThuong.Text)
            AddCategoryIfNotExists("CapKyQuyetDinh", cbCapKyQuyetDinh.Text)
            AddCategoryIfNotExists("NguoiKyQuyetDinh", cbNguoiKyQuyetDinh.Text)

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")
                conn.Open()

                Dim cmd As New SQLiteCommand(
                    "INSERT INTO KhenThuong (Nam, LoaiKhenThuong, HoVaTen, NgaySinh, CapBac, ChucVu, DonVi, " &
                    "HinhThucKhenThuong, LyDoKhenThuong, CapKyQuyetDinh, SoQuyetDinh, NguoiKyQuyetDinh, GhiChu, NgayVaoDang, QueQuan, CCCD) " &
                    "VALUES (@Nam, @LoaiKhenThuong, @HoVaTen, @NgaySinh, @CapBac, @ChucVu, @DonVi, " &
                    "@HinhThucKhenThuong, @LyDoKhenThuong, @CapKyQuyetDinh, @SoQuyetDinh, @NguoiKyQuyetDinh, @GhiChu, @NgayVaoDang, @QueQuan, @CCCD)", conn)

                cmd.Parameters.AddWithValue("@Nam", cbNam.Text)
                cmd.Parameters.AddWithValue("@LoaiKhenThuong", cbLoaiKhenThuong.Text)
                cmd.Parameters.AddWithValue("@HoVaTen", txtHoVaTen.Text)
                cmd.Parameters.AddWithValue("@NgaySinh", If(String.IsNullOrEmpty(txtNgaySinh.Text), DBNull.Value, txtNgaySinh.Text))
                cmd.Parameters.AddWithValue("@CapBac", If(String.IsNullOrEmpty(cbCapBac.Text), DBNull.Value, cbCapBac.Text))
                cmd.Parameters.AddWithValue("@ChucVu", If(String.IsNullOrEmpty(cbChucVu.Text), DBNull.Value, cbChucVu.Text))
                cmd.Parameters.AddWithValue("@DonVi", If(String.IsNullOrEmpty(cbDonVi.Text), DBNull.Value, cbDonVi.Text))
                cmd.Parameters.AddWithValue("@HinhThucKhenThuong", If(String.IsNullOrEmpty(cbHinhThucKhenThuong.Text), DBNull.Value, cbHinhThucKhenThuong.Text))
                cmd.Parameters.AddWithValue("@LyDoKhenThuong", If(String.IsNullOrEmpty(cbLyDoKhenThuong.Text), DBNull.Value, cbLyDoKhenThuong.Text))
                cmd.Parameters.AddWithValue("@CapKyQuyetDinh", If(String.IsNullOrEmpty(cbCapKyQuyetDinh.Text), DBNull.Value, cbCapKyQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@NguoiKyQuyetDinh", If(String.IsNullOrEmpty(cbNguoiKyQuyetDinh.Text), DBNull.Value, cbNguoiKyQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@SoQuyetDinh", If(String.IsNullOrEmpty(txtSoQuyetDinh.Text), DBNull.Value, txtSoQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@GhiChu", If(String.IsNullOrEmpty(txtGhiChu.Text), DBNull.Value, txtGhiChu.Text))
                cmd.Parameters.AddWithValue("@NgayVaoDang", If(String.IsNullOrEmpty(txtNgayVaoDang.Text), DBNull.Value, txtNgayVaoDang.Text))
                cmd.Parameters.AddWithValue("@CCCD", If(String.IsNullOrEmpty(txtCCCD.Text), DBNull.Value, txtCCCD.Text))
                cmd.Parameters.AddWithValue("@QueQuan", If(String.IsNullOrEmpty(txtQueQuan.Text), DBNull.Value, txtQueQuan.Text))

                cmd.ExecuteNonQuery()
                conn.Close()
            End Using

            MessageBox.Show("Thêm mới bản ghi thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadAllRecords()
            ClearForm()
            currentRecordId = -1

        Catch ex As Exception
            MessageBox.Show("Lỗi khi thêm mới: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnCapNhat_Click(sender As Object, e As EventArgs) Handles btnCapNhat.Click
        Try
            If currentRecordId = -1 Then
                MessageBox.Show("Vui lòng chọn bản ghi để cập nhật!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validate required fields
            If String.IsNullOrEmpty(cbNam.Text) Then
                MessageBox.Show("Vui lòng chọn Năm!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrEmpty(cbLoaiKhenThuong.Text) Then
                MessageBox.Show("Vui lòng chọn Loại khen thưởng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If String.IsNullOrEmpty(txtHoVaTen.Text) Then
                MessageBox.Show("Vui lòng nhập Họ và tên!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validate CCCD
            If txtCCCD.Text.Length > 0 AndAlso txtCCCD.Text.Length <> 12 Then
                MessageBox.Show("CCCD phải có đúng 12 ký tự!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Add categories if not exists
            AddCategoryIfNotExists("LoaiKhenThuong", cbLoaiKhenThuong.Text)
            AddCategoryIfNotExists("CapBac", cbCapBac.Text)
            AddCategoryIfNotExists("ChucVu", cbChucVu.Text)
            AddCategoryIfNotExists("DonVi", cbDonVi.Text)
            AddCategoryIfNotExists("HinhThucKhenThuong", cbHinhThucKhenThuong.Text)
            AddCategoryIfNotExists("LyDoKhenThuong", cbLyDoKhenThuong.Text)
            AddCategoryIfNotExists("CapKyQuyetDinh", cbCapKyQuyetDinh.Text)
            AddCategoryIfNotExists("NguoiKyQuyetDinh", cbNguoiKyQuyetDinh.Text)

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")
                conn.Open()

                Dim cmd As New SQLiteCommand(
                    "UPDATE KhenThuong SET Nam = @Nam, LoaiKhenThuong = @LoaiKhenThuong, HoVaTen = @HoVaTen, " &
                    "NgaySinh = @NgaySinh, CapBac = @CapBac, ChucVu = @ChucVu, DonVi = @DonVi, " &
                    "HinhThucKhenThuong = @HinhThucKhenThuong, LyDoKhenThuong = @LyDoKhenThuong, " &
                    "CapKyQuyetDinh = @CapKyQuyetDinh, SoQuyetDinh = @SoQuyetDinh, NguoiKyQuyetDinh = @NguoiKyQuyetDinh, QueQuan = @QueQuan, GhiChu = @GhiChu, " &
                    "NgayVaoDang = @NgayVaoDang, CCCD = @CCCD WHERE ID = @ID", conn)

                cmd.Parameters.AddWithValue("@Nam", cbNam.Text)
                cmd.Parameters.AddWithValue("@LoaiKhenThuong", cbLoaiKhenThuong.Text)
                cmd.Parameters.AddWithValue("@HoVaTen", txtHoVaTen.Text)
                cmd.Parameters.AddWithValue("@NgaySinh", If(String.IsNullOrEmpty(txtNgaySinh.Text), DBNull.Value, txtNgaySinh.Text))
                cmd.Parameters.AddWithValue("@CapBac", If(String.IsNullOrEmpty(cbCapBac.Text), DBNull.Value, cbCapBac.Text))
                cmd.Parameters.AddWithValue("@ChucVu", If(String.IsNullOrEmpty(cbChucVu.Text), DBNull.Value, cbChucVu.Text))
                cmd.Parameters.AddWithValue("@DonVi", If(String.IsNullOrEmpty(cbDonVi.Text), DBNull.Value, cbDonVi.Text))
                cmd.Parameters.AddWithValue("@HinhThucKhenThuong", If(String.IsNullOrEmpty(cbHinhThucKhenThuong.Text), DBNull.Value, cbHinhThucKhenThuong.Text))
                cmd.Parameters.AddWithValue("@LyDoKhenThuong", If(String.IsNullOrEmpty(cbLyDoKhenThuong.Text), DBNull.Value, cbLyDoKhenThuong.Text))
                cmd.Parameters.AddWithValue("@CapKyQuyetDinh", If(String.IsNullOrEmpty(cbCapKyQuyetDinh.Text), DBNull.Value, cbCapKyQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@NguoiKyQuyetDinh", If(String.IsNullOrEmpty(cbNguoiKyQuyetDinh.Text), DBNull.Value, cbNguoiKyQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@SoQuyetDinh", If(String.IsNullOrEmpty(txtSoQuyetDinh.Text), DBNull.Value, txtSoQuyetDinh.Text))
                cmd.Parameters.AddWithValue("@GhiChu", If(String.IsNullOrEmpty(txtGhiChu.Text), DBNull.Value, txtGhiChu.Text))
                cmd.Parameters.AddWithValue("@NgayVaoDang", If(String.IsNullOrEmpty(txtNgayVaoDang.Text), DBNull.Value, txtNgayVaoDang.Text))
                cmd.Parameters.AddWithValue("@CCCD", If(String.IsNullOrEmpty(txtCCCD.Text), DBNull.Value, txtCCCD.Text))
                cmd.Parameters.AddWithValue("@QueQuan", If(String.IsNullOrEmpty(txtQueQuan.Text), DBNull.Value, txtQueQuan.Text))
                cmd.Parameters.AddWithValue("@ID", currentRecordId)

                cmd.ExecuteNonQuery()
                conn.Close()
            End Using

            MessageBox.Show("Cập nhật bản ghi thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            LoadAllRecords()
            ClearForm()
            currentRecordId = -1

        Catch ex As Exception
            MessageBox.Show("Lỗi khi cập nhật: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnXuatFileExcel_Click(sender As Object, e As EventArgs) Handles btnXuatFileExcel.Click

        Try

            If lvKQTK.Items.Count = 0 Then

                MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Return

            End If

            Dim sfd As New SaveFileDialog()

            sfd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"

            sfd.FileName = "DanhSachKhenThuong.xlsx"

            If sfd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Dim wb As New XLWorkbook()

            Dim ws = wb.Worksheets.Add("KhenThuong")

            ' ===== HEADER =====

            ws.Cell(1, 1).Value = "STT"
            ws.Cell(1, 2).Value = "Năm"
            ws.Cell(1, 3).Value = "Loại khen thưởng"
            ws.Cell(1, 4).Value = "Họ và tên"
            ws.Cell(1, 5).Value = "Ngày sinh"
            ws.Cell(1, 6).Value = "Cấp bậc"
            ws.Cell(1, 7).Value = "Chức vụ"
            ws.Cell(1, 8).Value = "Đơn vị"
            ws.Cell(1, 9).Value = "Hình thức khen thưởng"
            ws.Cell(1, 10).Value = "Lý do khen thưởng"
            ws.Cell(1, 11).Value = "Cấp ký quyết định"
            ws.Cell(1, 12).Value = "Người ký ban hành"
            ws.Cell(1, 13).Value = "Số quyết định"
            ws.Cell(1, 14).Value = "Ngày vào Đảng"
            ws.Cell(1, 15).Value = "CCCD"
            ws.Cell(1, 16).Value = "Quê quán"
            ws.Cell(1, 17).Value = "Ghi chú"

            ' ===== FORMAT =====

            With ws.Range("A1:Q1")

                .Style.Font.Bold = True

                .Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center

                .Style.Alignment.Vertical = XLAlignmentVerticalValues.Center

                .Style.Alignment.WrapText = True

            End With

            ws.Columns().AdjustToContents()

            ' ===== DATA =====

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                Dim cmd As New SQLiteCommand("SELECT * FROM KhenThuong ORDER BY ID", conn)

                Dim reader As SQLiteDataReader = cmd.ExecuteReader()

                Dim row As Integer = 2
                Dim stt As Integer = 1

                While reader.Read()

                    ws.Cell(row, 1).Value = stt

                    ws.Cell(row, 2).Value = reader("Nam").ToString()

                    ws.Cell(row, 3).Value = reader("LoaiKhenThuong").ToString()

                    ws.Cell(row, 4).Value = reader("HoVaTen").ToString()

                    ws.Cell(row, 5).Value = If(reader("NgaySinh") Is DBNull.Value, "", reader("NgaySinh").ToString())

                    ws.Cell(row, 6).Value = If(reader("CapBac") Is DBNull.Value, "", reader("CapBac").ToString())

                    ws.Cell(row, 7).Value = If(reader("ChucVu") Is DBNull.Value, "", reader("ChucVu").ToString())

                    ws.Cell(row, 8).Value = If(reader("DonVi") Is DBNull.Value, "", reader("DonVi").ToString())

                    ws.Cell(row, 9).Value = If(reader("HinhThucKhenThuong") Is DBNull.Value, "", reader("HinhThucKhenThuong").ToString())

                    ws.Cell(row, 10).Value = If(reader("LyDoKhenThuong") Is DBNull.Value, "", reader("LyDoKhenThuong").ToString())

                    ws.Cell(row, 11).Value = If(reader("CapKyQuyetDinh") Is DBNull.Value, "", reader("CapKyQuyetDinh").ToString())

                    ws.Cell(row, 12).Value = If(reader("NguoiKyQuyetDinh") Is DBNull.Value, "", reader("NguoiKyQuyetDinh").ToString())

                    ws.Cell(row, 13).Value = If(reader("SoQuyetDinh") Is DBNull.Value, "", reader("SoQuyetDinh").ToString())

                    ws.Cell(row, 14).Value = If(reader("NgayVaoDang") Is DBNull.Value, "", reader("NgayVaoDang").ToString())

                    ws.Cell(row, 15).Value = If(reader("CCCD") Is DBNull.Value, "", reader("CCCD").ToString())

                    ws.Cell(row, 16).Value = If(reader("QueQuan") Is DBNull.Value, "", reader("QueQuan").ToString())

                    ws.Cell(row, 17).Value = If(reader("GhiChu") Is DBNull.Value, "", reader("GhiChu").ToString())

                    row += 1
                    stt += 1

                End While

                reader.Close()

            End Using

            ' ===== FORMAT =====

            ws.Range("A1:Q1").Style.Font.Bold = True

            ws.Columns().AdjustToContents()

            wb.SaveAs(sfd.FileName)

            MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Lỗi xuất Excel: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub btnNhapExcel_Click(sender As Object, e As EventArgs) Handles btnNhapExcel.Click

        Try

            Dim ofd As New OpenFileDialog()

            ofd.Filter = "Excel Workbook (*.xlsx)|*.xlsx"

            ofd.Title = "Chọn file Excel"

            If ofd.ShowDialog() <> DialogResult.OK Then
                Return
            End If

            Dim duplicates As New List(Of String)

            Dim importedCount As Integer = 0

            Using wb As New XLWorkbook(ofd.FileName)

                Dim ws =
                    wb.Worksheet(1)

                Dim lastRow As Integer = ws.LastRowUsed().RowNumber()

                Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                    conn.Open()

                    For row As Integer = 2 To lastRow

                        Dim nam As String = ws.Cell(row, 2).GetString().Trim()

                        Dim loaiKhenThuong As String = ws.Cell(row, 3).GetString().Trim()

                        Dim hoVaTen As String = ws.Cell(row, 4).GetString().Trim()

                        Dim ngaySinh As String = ws.Cell(row, 5).GetString().Trim()

                        Dim capBac As String = ws.Cell(row, 6).GetString().Trim()

                        Dim chucVu As String = ws.Cell(row, 7).GetString().Trim()

                        Dim donVi As String = ws.Cell(row, 8).GetString().Trim()

                        Dim hinhThucKhenThuong As String = ws.Cell(row, 9).GetString().Trim()

                        Dim lyDoKhenThuong As String = ws.Cell(row, 10).GetString().Trim()

                        Dim capKyQuyetDinh As String = ws.Cell(row, 11).GetString().Trim()

                        Dim nguoiKyQuyetDinh As String = ws.Cell(row, 12).GetString().Trim()

                        Dim soQuyetDinh As String = ws.Cell(row, 13).GetString().Trim()

                        Dim ngayVaoDang As String = ws.Cell(row, 14).GetString().Trim()

                        Dim cccd As String = ws.Cell(row, 15).GetString().Trim()

                        Dim queQuan As String = ws.Cell(row, 16).GetString().Trim()

                        Dim ghiChu As String = ws.Cell(row, 17).GetString().Trim()

                        ' ===== BỎ QUA DÒNG RỖNG =====

                        If String.IsNullOrWhiteSpace(hoVaTen) Then
                            Continue For
                        End If

                        ' ===== KIỂM TRA TRÙNG =====

                        Dim checkCmd As New SQLiteCommand(
                        "SELECT ID FROM KhenThuong " &
                        "WHERE Nam = @Nam " &
                        "AND LoaiKhenThuong = @LoaiKhenThuong " &
                        "AND HoVaTen = @HoVaTen " &
                        "AND SoQuyetDinh = @SoQuyetDinh " &
                        "AND (CCCD = @CCCD OR NgaySinh = @NgaySinh)", conn)

                        checkCmd.Parameters.AddWithValue("@Nam", nam)

                        checkCmd.Parameters.AddWithValue("@SoQuyetDinh", soQuyetDinh)

                        checkCmd.Parameters.AddWithValue("@LoaiKhenThuong", loaiKhenThuong)

                        checkCmd.Parameters.AddWithValue("@HoVaTen", hoVaTen)

                        checkCmd.Parameters.AddWithValue("@CCCD", cccd)

                        checkCmd.Parameters.AddWithValue("@NgaySinh", ngaySinh)

                        Dim existingId As Object =
                            checkCmd.ExecuteScalar()

                        checkCmd.Dispose()

                        If existingId IsNot Nothing AndAlso
                           existingId IsNot DBNull.Value Then

                            duplicates.Add(hoVaTen & " - " & nam)

                        Else

                            ' ===== THÊM DANH MỤC =====

                            AddCategoryIfNotExists("LoaiKhenThuong", loaiKhenThuong)

                            AddCategoryIfNotExists("CapBac", capBac)

                            AddCategoryIfNotExists("ChucVu", chucVu)

                            AddCategoryIfNotExists("DonVi", donVi)

                            AddCategoryIfNotExists("HinhThucKhenThuong", hinhThucKhenThuong)

                            AddCategoryIfNotExists("LyDoKhenThuong", lyDoKhenThuong)

                            AddCategoryIfNotExists("CapKyQuyetDinh", capKyQuyetDinh)

                            ' ===== INSERT =====

                            Dim insertCmd As New SQLiteCommand(
                            "INSERT INTO KhenThuong " &
                            "(Nam, LoaiKhenThuong, HoVaTen, " &
                            "NgaySinh, CapBac, ChucVu, DonVi, " &
                            "HinhThucKhenThuong, LyDoKhenThuong, " &
                            "CapKyQuyetDinh, NguoiKyQuyetDinh, " &
                            "SoQuyetDinh, NgayVaoDang, CCCD, " &
                            "QueQuan, GhiChu) " &
                            "VALUES " &
                            "(@Nam, @LoaiKhenThuong, @HoVaTen, " &
                            "@NgaySinh, @CapBac, @ChucVu, @DonVi, " &
                            "@HinhThucKhenThuong, @LyDoKhenThuong, " &
                            "@CapKyQuyetDinh, @NguoiKyQuyetDinh, " &
                            "@SoQuyetDinh, @NgayVaoDang, @CCCD, " &
                            "@QueQuan, @GhiChu)", conn)

                            insertCmd.Parameters.AddWithValue("@Nam", If(String.IsNullOrWhiteSpace(nam), DBNull.Value, nam))

                            insertCmd.Parameters.AddWithValue("@LoaiKhenThuong", If(String.IsNullOrWhiteSpace(loaiKhenThuong), DBNull.Value, loaiKhenThuong))

                            insertCmd.Parameters.AddWithValue("@HoVaTen", hoVaTen)

                            insertCmd.Parameters.AddWithValue("@NgaySinh", If(String.IsNullOrWhiteSpace(ngaySinh), DBNull.Value, ngaySinh))

                            insertCmd.Parameters.AddWithValue("@CapBac", If(String.IsNullOrWhiteSpace(capBac), DBNull.Value, capBac))

                            insertCmd.Parameters.AddWithValue("@ChucVu", If(String.IsNullOrWhiteSpace(chucVu), DBNull.Value, chucVu))

                            insertCmd.Parameters.AddWithValue("@DonVi", If(String.IsNullOrWhiteSpace(donVi), DBNull.Value, donVi))

                            insertCmd.Parameters.AddWithValue("@HinhThucKhenThuong", If(String.IsNullOrWhiteSpace(hinhThucKhenThuong), DBNull.Value, hinhThucKhenThuong))

                            insertCmd.Parameters.AddWithValue("@LyDoKhenThuong", If(String.IsNullOrWhiteSpace(lyDoKhenThuong), DBNull.Value, lyDoKhenThuong))

                            insertCmd.Parameters.AddWithValue("@CapKyQuyetDinh", If(String.IsNullOrWhiteSpace(capKyQuyetDinh), DBNull.Value, capKyQuyetDinh))

                            insertCmd.Parameters.AddWithValue("@NguoiKyQuyetDinh", If(String.IsNullOrWhiteSpace(nguoiKyQuyetDinh), DBNull.Value, nguoiKyQuyetDinh))

                            insertCmd.Parameters.AddWithValue("@SoQuyetDinh", If(String.IsNullOrWhiteSpace(soQuyetDinh), DBNull.Value, soQuyetDinh))

                            insertCmd.Parameters.AddWithValue("@NgayVaoDang", If(String.IsNullOrWhiteSpace(ngayVaoDang), DBNull.Value, ngayVaoDang))

                            insertCmd.Parameters.AddWithValue("@CCCD", If(String.IsNullOrWhiteSpace(cccd), DBNull.Value, cccd))

                            insertCmd.Parameters.AddWithValue("@QueQuan", If(String.IsNullOrWhiteSpace(queQuan), DBNull.Value, queQuan))

                            insertCmd.Parameters.AddWithValue("@GhiChu", If(String.IsNullOrWhiteSpace(ghiChu), DBNull.Value, ghiChu))

                            insertCmd.ExecuteNonQuery()

                            insertCmd.Dispose()

                            importedCount += 1

                        End If

                    Next

                    conn.Close()

                End Using

            End Using

            Dim message As String = "Đã nhập thành công " & importedCount & " bản ghi."

            If duplicates.Count > 0 Then

                message &= vbCrLf & "Bỏ qua " & duplicates.Count & " bản ghi trùng lặp."

            End If

            MessageBox.Show(message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

            LoadAllRecords()

            ClearForm()

        Catch ex As Exception

            MessageBox.Show("Lỗi nhập Excel: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub
    Private Sub btnXoaBanGhi_Click(sender As Object, e As EventArgs) Handles btnXoaBanGhi.Click

        Try

            If currentRecordId = -1 Then

                MessageBox.Show("Vui lòng chọn bản ghi để xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                Return

            End If

            Dim result As DialogResult = MessageBox.Show("Bạn có chắc chắn muốn xóa bản ghi này? Hành động này không thể hoàn tác!", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.No Then
                Return
            End If

            Using conn As New SQLiteConnection("Data Source=" & dbPath & ";Version=3;")

                conn.Open()

                ' ===== XÓA BẢN GHI =====

                Dim deleteCmd As New SQLiteCommand("DELETE FROM KhenThuong WHERE ID = @ID", conn)

                deleteCmd.Parameters.AddWithValue("@ID", currentRecordId)

                deleteCmd.ExecuteNonQuery()

                deleteCmd.Dispose()

                ' ===== SẮP XẾP LẠI ID =====

                Dim ids As New List(Of Integer)

                Dim selectCmd As New SQLiteCommand("SELECT ID FROM KhenThuong ORDER BY ID", conn)

                Dim reader As SQLiteDataReader = selectCmd.ExecuteReader()

                While reader.Read()

                    ids.Add(Convert.ToInt32(reader("ID")))

                End While

                reader.Close()
                selectCmd.Dispose()

                Dim newId As Integer = 1

                For Each oldId As Integer In ids

                    If oldId <> newId Then

                        Dim updateCmd As New SQLiteCommand("UPDATE KhenThuong " & "SET ID = @NewID " & "WHERE ID = @OldID", conn)

                        updateCmd.Parameters.AddWithValue("@NewID", newId)

                        updateCmd.Parameters.AddWithValue("@OldID", oldId)

                        updateCmd.ExecuteNonQuery()

                        updateCmd.Dispose()

                    End If

                    newId += 1

                Next

                ' ===== RESET SQLITE AUTOINCREMENT =====

                Dim resetCmd As New SQLiteCommand("UPDATE sqlite_sequence " & "SET seq = (SELECT IFNULL(MAX(ID),0) FROM KhenThuong) " & "WHERE name='KhenThuong'", conn)

                resetCmd.ExecuteNonQuery()
                resetCmd.Dispose()

                conn.Close()

            End Using

            MessageBox.Show("Xóa bản ghi thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information)

            LoadAllRecords()

            ClearForm()

            currentRecordId = -1

        Catch ex As Exception

            MessageBox.Show("Lỗi khi xóa: " & ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

End Class
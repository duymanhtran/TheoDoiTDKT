<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        btnThemMoi = New Button()
        btnTimKiem = New Button()
        btnCapNhat = New Button()
        btnLamMoi = New Button()
        btnXuatFileExcel = New Button()
        btnXuatDB = New Button()
        btnNhapDB = New Button()
        btnXoaBanGhi = New Button()
        Label1 = New Label()
        Label2 = New Label()
        Label3 = New Label()
        Label4 = New Label()
        Label5 = New Label()
        Label6 = New Label()
        Label7 = New Label()
        Label8 = New Label()
        Label9 = New Label()
        Label10 = New Label()
        Label11 = New Label()
        Label12 = New Label()
        Label13 = New Label()
        Label14 = New Label()
        Label15 = New Label()
        Label16 = New Label()
        Label17 = New Label()
        Label18 = New Label()
        txtSTT = New TextBox()
        txtHoVaTen = New TextBox()
        txtNgaySinh = New TextBox()
        txtNgayVaoDang = New TextBox()
        txtCCCD = New TextBox()
        txtQueQuan = New TextBox()
        txtSoQuyetDinh = New TextBox()
        Label19 = New Label()
        txtGhiChu = New TextBox()
        cbCapKyQuyetDinh = New ComboBox()
        cbNguoiKyQuyetDinh = New ComboBox()
        cbHinhThucKhenThuong = New ComboBox()
        cbDonVi = New ComboBox()
        cbChucVu = New ComboBox()
        cbCapBac = New ComboBox()
        cbLyDoKhenThuong = New ComboBox()
        cbLoaiKhenThuong = New ComboBox()
        lvKQTK = New ListView()
        lbKQTK = New Label()
        dtpNgaySinh = New DateTimePicker()
        dtpNgayVaoDang = New DateTimePicker()
        cbNam = New ComboBox()
        Label20 = New Label()
        lbThongTin = New Label()
        Timer1 = New Timer(components)
        Timer2 = New Timer(components)
        btnNhapExcel = New Button()
        SuspendLayout()
        ' 
        ' btnThemMoi
        ' 
        btnThemMoi.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnThemMoi.Location = New Point(23, 39)
        btnThemMoi.Name = "btnThemMoi"
        btnThemMoi.Size = New Size(104, 39)
        btnThemMoi.TabIndex = 17
        btnThemMoi.Text = "Thêm mới"
        btnThemMoi.UseVisualStyleBackColor = True
        ' 
        ' btnTimKiem
        ' 
        btnTimKiem.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnTimKiem.Location = New Point(142, 39)
        btnTimKiem.Name = "btnTimKiem"
        btnTimKiem.Size = New Size(103, 39)
        btnTimKiem.TabIndex = 18
        btnTimKiem.Text = "Tìm kiếm"
        btnTimKiem.UseVisualStyleBackColor = True
        ' 
        ' btnCapNhat
        ' 
        btnCapNhat.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnCapNhat.Location = New Point(535, 39)
        btnCapNhat.Name = "btnCapNhat"
        btnCapNhat.Size = New Size(94, 39)
        btnCapNhat.TabIndex = 19
        btnCapNhat.Text = "Cập nhật"
        btnCapNhat.UseVisualStyleBackColor = True
        ' 
        ' btnLamMoi
        ' 
        btnLamMoi.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnLamMoi.Location = New Point(640, 39)
        btnLamMoi.Name = "btnLamMoi"
        btnLamMoi.Size = New Size(95, 39)
        btnLamMoi.TabIndex = 20
        btnLamMoi.Text = "Làm mới"
        btnLamMoi.UseVisualStyleBackColor = True
        ' 
        ' btnXuatFileExcel
        ' 
        btnXuatFileExcel.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnXuatFileExcel.Location = New Point(749, 39)
        btnXuatFileExcel.Name = "btnXuatFileExcel"
        btnXuatFileExcel.Size = New Size(113, 39)
        btnXuatFileExcel.TabIndex = 21
        btnXuatFileExcel.Text = "Xuất Excel"
        btnXuatFileExcel.UseVisualStyleBackColor = True
        ' 
        ' btnXuatDB
        ' 
        btnXuatDB.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnXuatDB.Location = New Point(1002, 39)
        btnXuatDB.Name = "btnXuatDB"
        btnXuatDB.Size = New Size(101, 39)
        btnXuatDB.TabIndex = 22
        btnXuatDB.Text = "Xuất DB"
        btnXuatDB.UseVisualStyleBackColor = True
        ' 
        ' btnNhapDB
        ' 
        btnNhapDB.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnNhapDB.Location = New Point(1117, 39)
        btnNhapDB.Name = "btnNhapDB"
        btnNhapDB.Size = New Size(101, 39)
        btnNhapDB.TabIndex = 23
        btnNhapDB.Text = "Nhập DB"
        btnNhapDB.UseVisualStyleBackColor = True
        ' 
        ' btnXoaBanGhi
        ' 
        btnXoaBanGhi.BackColor = Color.Yellow
        btnXoaBanGhi.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnXoaBanGhi.ForeColor = Color.Red
        btnXoaBanGhi.Location = New Point(1279, 27)
        btnXoaBanGhi.Name = "btnXoaBanGhi"
        btnXoaBanGhi.Size = New Size(85, 63)
        btnXoaBanGhi.TabIndex = 101
        btnXoaBanGhi.Text = "Xóa " & vbCrLf & "bản ghi"
        btnXoaBanGhi.UseVisualStyleBackColor = False
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label1.Location = New Point(251, 47)
        Label1.Name = "Label1"
        Label1.Size = New Size(88, 22)
        Label1.TabIndex = 100
        Label1.Text = "Tìm thấy:"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label2.Location = New Point(380, 47)
        Label2.Name = "Label2"
        Label2.Size = New Size(142, 22)
        Label2.TabIndex = 100
        Label2.Text = "kết quả phù hợp"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label3.Location = New Point(23, 110)
        Label3.Name = "Label3"
        Label3.Size = New Size(47, 22)
        Label3.TabIndex = 100
        Label3.Text = "STT"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label4.Location = New Point(94, 110)
        Label4.Name = "Label4"
        Label4.Size = New Size(49, 22)
        Label4.TabIndex = 100
        Label4.Text = "Năm"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label5.Location = New Point(181, 110)
        Label5.Name = "Label5"
        Label5.Size = New Size(155, 22)
        Label5.TabIndex = 100
        Label5.Text = "Loại khen thưởng"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label6.Location = New Point(420, 110)
        Label6.Name = "Label6"
        Label6.Size = New Size(88, 22)
        Label6.TabIndex = 100
        Label6.Text = "Họ và tên"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label7.Location = New Point(590, 110)
        Label7.Name = "Label7"
        Label7.Size = New Size(91, 22)
        Label7.TabIndex = 100
        Label7.Text = "Ngày sinh"
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label8.Location = New Point(710, 88)
        Label8.Name = "Label8"
        Label8.Size = New Size(88, 44)
        Label8.TabIndex = 100
        Label8.Text = "Ngày " & vbCrLf & "vào Đảng"
        Label8.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label9.Location = New Point(863, 110)
        Label9.Name = "Label9"
        Label9.Size = New Size(137, 22)
        Label9.TabIndex = 100
        Label9.Text = "Số TĐV/CCCD"
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label10.Location = New Point(1153, 110)
        Label10.Name = "Label10"
        Label10.Size = New Size(89, 22)
        Label10.TabIndex = 100
        Label10.Text = "Quê quán"
        ' 
        ' Label11
        ' 
        Label11.AutoSize = True
        Label11.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label11.Location = New Point(79, 200)
        Label11.Name = "Label11"
        Label11.Size = New Size(78, 22)
        Label11.TabIndex = 100
        Label11.Text = "Cấp bậc"
        ' 
        ' Label12
        ' 
        Label12.AutoSize = True
        Label12.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label12.Location = New Point(304, 200)
        Label12.Name = "Label12"
        Label12.Size = New Size(78, 22)
        Label12.TabIndex = 100
        Label12.Text = "Chức vụ"
        ' 
        ' Label13
        ' 
        Label13.AutoSize = True
        Label13.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label13.Location = New Point(570, 200)
        Label13.Name = "Label13"
        Label13.Size = New Size(64, 22)
        Label13.TabIndex = 100
        Label13.Text = "Đơn vị"
        ' 
        ' Label14
        ' 
        Label14.AutoSize = True
        Label14.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label14.Location = New Point(780, 200)
        Label14.Name = "Label14"
        Label14.Size = New Size(198, 22)
        Label14.TabIndex = 100
        Label14.Text = "Hình thức khen thưởng"
        ' 
        ' Label15
        ' 
        Label15.AutoSize = True
        Label15.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label15.Location = New Point(622, 291)
        Label15.Name = "Label15"
        Label15.Size = New Size(177, 22)
        Label15.TabIndex = 100
        Label15.Text = "Cấp ký khen thưởng"
        ' 
        ' Label16
        ' 
        Label16.AutoSize = True
        Label16.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label16.Location = New Point(123, 291)
        Label16.Name = "Label16"
        Label16.Size = New Size(164, 22)
        Label16.TabIndex = 100
        Label16.Text = "Lý do khen thưởng"
        ' 
        ' Label17
        ' 
        Label17.AutoSize = True
        Label17.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label17.Location = New Point(439, 291)
        Label17.Name = "Label17"
        Label17.Size = New Size(120, 22)
        Label17.TabIndex = 100
        Label17.Text = "Số quyết định"
        ' 
        ' Label18
        ' 
        Label18.AutoSize = True
        Label18.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label18.Location = New Point(831, 291)
        Label18.Name = "Label18"
        Label18.Size = New Size(175, 22)
        Label18.TabIndex = 100
        Label18.Text = "Người ký quyết định"
        ' 
        ' txtSTT
        ' 
        txtSTT.Location = New Point(23, 135)
        txtSTT.Name = "txtSTT"
        txtSTT.ReadOnly = True
        txtSTT.Size = New Size(47, 23)
        txtSTT.TabIndex = 0
        ' 
        ' txtHoVaTen
        ' 
        txtHoVaTen.Location = New Point(357, 135)
        txtHoVaTen.Name = "txtHoVaTen"
        txtHoVaTen.Size = New Size(214, 23)
        txtHoVaTen.TabIndex = 3
        ' 
        ' txtNgaySinh
        ' 
        txtNgaySinh.Location = New Point(582, 135)
        txtNgaySinh.Name = "txtNgaySinh"
        txtNgaySinh.Size = New Size(91, 23)
        txtNgaySinh.TabIndex = 4
        ' 
        ' txtNgayVaoDang
        ' 
        txtNgayVaoDang.Location = New Point(700, 135)
        txtNgayVaoDang.Name = "txtNgayVaoDang"
        txtNgayVaoDang.Size = New Size(91, 23)
        txtNgayVaoDang.TabIndex = 5
        ' 
        ' txtCCCD
        ' 
        txtCCCD.Location = New Point(837, 135)
        txtCCCD.Name = "txtCCCD"
        txtCCCD.Size = New Size(189, 23)
        txtCCCD.TabIndex = 6
        ' 
        ' txtQueQuan
        ' 
        txtQueQuan.Location = New Point(1052, 135)
        txtQueQuan.Name = "txtQueQuan"
        txtQueQuan.Size = New Size(290, 23)
        txtQueQuan.TabIndex = 7
        ' 
        ' txtSoQuyetDinh
        ' 
        txtSoQuyetDinh.Location = New Point(394, 316)
        txtSoQuyetDinh.Name = "txtSoQuyetDinh"
        txtSoQuyetDinh.Size = New Size(210, 23)
        txtSoQuyetDinh.TabIndex = 13
        ' 
        ' Label19
        ' 
        Label19.AutoSize = True
        Label19.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label19.Location = New Point(1160, 200)
        Label19.Name = "Label19"
        Label19.Size = New Size(74, 22)
        Label19.TabIndex = 100
        Label19.Text = "Ghi chú"
        ' 
        ' txtGhiChu
        ' 
        txtGhiChu.Location = New Point(1052, 225)
        txtGhiChu.Multiline = True
        txtGhiChu.Name = "txtGhiChu"
        txtGhiChu.Size = New Size(290, 473)
        txtGhiChu.TabIndex = 16
        ' 
        ' cbCapKyQuyetDinh
        ' 
        cbCapKyQuyetDinh.FormattingEnabled = True
        cbCapKyQuyetDinh.Location = New Point(612, 317)
        cbCapKyQuyetDinh.Name = "cbCapKyQuyetDinh"
        cbCapKyQuyetDinh.Size = New Size(195, 23)
        cbCapKyQuyetDinh.TabIndex = 14
        ' 
        ' cbNguoiKyQuyetDinh
        ' 
        cbNguoiKyQuyetDinh.FormattingEnabled = True
        cbNguoiKyQuyetDinh.Location = New Point(812, 316)
        cbNguoiKyQuyetDinh.Name = "cbNguoiKyQuyetDinh"
        cbNguoiKyQuyetDinh.Size = New Size(213, 23)
        cbNguoiKyQuyetDinh.TabIndex = 15
        ' 
        ' cbHinhThucKhenThuong
        ' 
        cbHinhThucKhenThuong.FormattingEnabled = True
        cbHinhThucKhenThuong.Location = New Point(733, 225)
        cbHinhThucKhenThuong.Name = "cbHinhThucKhenThuong"
        cbHinhThucKhenThuong.Size = New Size(293, 23)
        cbHinhThucKhenThuong.TabIndex = 11
        ' 
        ' cbDonVi
        ' 
        cbDonVi.FormattingEnabled = True
        cbDonVi.Location = New Point(478, 225)
        cbDonVi.Name = "cbDonVi"
        cbDonVi.Size = New Size(249, 23)
        cbDonVi.TabIndex = 10
        ' 
        ' cbChucVu
        ' 
        cbChucVu.FormattingEnabled = True
        cbChucVu.Location = New Point(215, 225)
        cbChucVu.Name = "cbChucVu"
        cbChucVu.Size = New Size(257, 23)
        cbChucVu.TabIndex = 9
        ' 
        ' cbCapBac
        ' 
        cbCapBac.FormattingEnabled = True
        cbCapBac.Location = New Point(33, 225)
        cbCapBac.Name = "cbCapBac"
        cbCapBac.Size = New Size(171, 23)
        cbCapBac.TabIndex = 8
        ' 
        ' cbLyDoKhenThuong
        ' 
        cbLyDoKhenThuong.FormattingEnabled = True
        cbLyDoKhenThuong.Location = New Point(24, 316)
        cbLyDoKhenThuong.Name = "cbLyDoKhenThuong"
        cbLyDoKhenThuong.Size = New Size(362, 23)
        cbLyDoKhenThuong.TabIndex = 12
        ' 
        ' cbLoaiKhenThuong
        ' 
        cbLoaiKhenThuong.FormattingEnabled = True
        cbLoaiKhenThuong.Location = New Point(165, 135)
        cbLoaiKhenThuong.Name = "cbLoaiKhenThuong"
        cbLoaiKhenThuong.Size = New Size(186, 23)
        cbLoaiKhenThuong.TabIndex = 2
        ' 
        ' lvKQTK
        ' 
        lvKQTK.Location = New Point(24, 406)
        lvKQTK.Name = "lvKQTK"
        lvKQTK.Size = New Size(1003, 292)
        lvKQTK.TabIndex = 25
        lvKQTK.UseCompatibleStateImageBehavior = False
        ' 
        ' lbKQTK
        ' 
        lbKQTK.AutoSize = True
        lbKQTK.BackColor = Color.Yellow
        lbKQTK.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lbKQTK.ForeColor = Color.Red
        lbKQTK.Location = New Point(358, 47)
        lbKQTK.Name = "lbKQTK"
        lbKQTK.Size = New Size(20, 22)
        lbKQTK.TabIndex = 100
        lbKQTK.Text = "0"
        ' 
        ' dtpNgaySinh
        ' 
        dtpNgaySinh.Location = New Point(673, 135)
        dtpNgaySinh.Name = "dtpNgaySinh"
        dtpNgaySinh.Size = New Size(18, 23)
        dtpNgaySinh.TabIndex = 98
        ' 
        ' dtpNgayVaoDang
        ' 
        dtpNgayVaoDang.Location = New Point(790, 135)
        dtpNgayVaoDang.Name = "dtpNgayVaoDang"
        dtpNgayVaoDang.Size = New Size(18, 23)
        dtpNgayVaoDang.TabIndex = 99
        ' 
        ' cbNam
        ' 
        cbNam.FormattingEnabled = True
        cbNam.Location = New Point(85, 134)
        cbNam.Name = "cbNam"
        cbNam.Size = New Size(67, 23)
        cbNam.TabIndex = 1
        ' 
        ' Label20
        ' 
        Label20.AutoSize = True
        Label20.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        Label20.ForeColor = Color.Blue
        Label20.Location = New Point(427, 369)
        Label20.Name = "Label20"
        Label20.Size = New Size(196, 22)
        Label20.TabIndex = 100
        Label20.Text = "KẾT QUẢ TÌM KIẾM"
        ' 
        ' lbThongTin
        ' 
        lbThongTin.AutoSize = True
        lbThongTin.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lbThongTin.Location = New Point(-3, 709)
        lbThongTin.Name = "lbThongTin"
        lbThongTin.Size = New Size(1478, 22)
        lbThongTin.TabIndex = 1
        lbThongTin.Text = "Phần mềm TĐKT - giải pháp ứng dụng chuyển đổi số trong theo dõi, thống kê hoạt động thi đua khen thưởng tại đơn vị - Tác giả sản phẩm: Trần Duy Mạnh, SĐT: 098.110.7698    --"
        ' 
        ' Timer1
        ' 
        ' 
        ' Timer2
        ' 
        ' 
        ' btnNhapExcel
        ' 
        btnNhapExcel.Font = New Font("Times New Roman", 14.25F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnNhapExcel.Location = New Point(876, 39)
        btnNhapExcel.Name = "btnNhapExcel"
        btnNhapExcel.Size = New Size(113, 39)
        btnNhapExcel.TabIndex = 23
        btnNhapExcel.Text = "Nhập Excel"
        btnNhapExcel.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1367, 737)
        Controls.Add(dtpNgayVaoDang)
        Controls.Add(dtpNgaySinh)
        Controls.Add(lvKQTK)
        Controls.Add(cbNam)
        Controls.Add(cbLoaiKhenThuong)
        Controls.Add(cbCapBac)
        Controls.Add(cbLyDoKhenThuong)
        Controls.Add(cbChucVu)
        Controls.Add(cbDonVi)
        Controls.Add(cbHinhThucKhenThuong)
        Controls.Add(cbNguoiKyQuyetDinh)
        Controls.Add(cbCapKyQuyetDinh)
        Controls.Add(txtGhiChu)
        Controls.Add(txtQueQuan)
        Controls.Add(txtSoQuyetDinh)
        Controls.Add(txtCCCD)
        Controls.Add(txtNgayVaoDang)
        Controls.Add(txtNgaySinh)
        Controls.Add(txtHoVaTen)
        Controls.Add(txtSTT)
        Controls.Add(Label8)
        Controls.Add(Label19)
        Controls.Add(Label10)
        Controls.Add(Label9)
        Controls.Add(Label7)
        Controls.Add(Label6)
        Controls.Add(Label5)
        Controls.Add(Label4)
        Controls.Add(Label15)
        Controls.Add(Label14)
        Controls.Add(Label13)
        Controls.Add(Label12)
        Controls.Add(Label18)
        Controls.Add(Label20)
        Controls.Add(Label17)
        Controls.Add(Label16)
        Controls.Add(Label11)
        Controls.Add(Label3)
        Controls.Add(lbThongTin)
        Controls.Add(Label2)
        Controls.Add(lbKQTK)
        Controls.Add(Label1)
        Controls.Add(btnXoaBanGhi)
        Controls.Add(btnNhapExcel)
        Controls.Add(btnNhapDB)
        Controls.Add(btnXuatDB)
        Controls.Add(btnXuatFileExcel)
        Controls.Add(btnLamMoi)
        Controls.Add(btnCapNhat)
        Controls.Add(btnTimKiem)
        Controls.Add(btnThemMoi)
        FormBorderStyle = FormBorderStyle.Fixed3D
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "Form1"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Phần mềm theo dõi thi đua khen thưởng"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnThemMoi As Button
    Friend WithEvents btnTimKiem As Button
    Friend WithEvents btnCapNhat As Button
    Friend WithEvents btnLamMoi As Button
    Friend WithEvents btnXuatFileExcel As Button
    Friend WithEvents btnXuatDB As Button
    Friend WithEvents btnNhapDB As Button
    Friend WithEvents btnXoaBanGhi As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents txtSTT As TextBox
    Friend WithEvents txtHoVaTen As TextBox
    Friend WithEvents txtNgaySinh As TextBox
    Friend WithEvents txtNgayVaoDang As TextBox
    Friend WithEvents txtCCCD As TextBox
    Friend WithEvents txtQueQuan As TextBox
    Friend WithEvents txtSoQuyetDinh As TextBox
    Friend WithEvents Label19 As Label
    Friend WithEvents txtGhiChu As TextBox
    Friend WithEvents cbCapKyQuyetDinh As ComboBox
    Friend WithEvents cbNguoiKyQuyetDinh As ComboBox
    Friend WithEvents cbHinhThucKhenThuong As ComboBox
    Friend WithEvents cbDonVi As ComboBox
    Friend WithEvents cbChucVu As ComboBox
    Friend WithEvents cbCapBac As ComboBox
    Friend WithEvents cbLyDoKhenThuong As ComboBox
    Friend WithEvents cbLoaiKhenThuong As ComboBox
    Friend WithEvents lvKQTK As ListView
    Friend WithEvents lbKQTK As Label
    Friend WithEvents dtpNgaySinh As DateTimePicker
    Friend WithEvents dtpNgayVaoDang As DateTimePicker
    Friend WithEvents cbNam As ComboBox
    Friend WithEvents Label20 As Label
    Friend WithEvents lbThongTin As Label
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Timer2 As Timer
    Friend WithEvents btnNhapExcel As Button

End Class

Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim waitDlg As WaitDialog   '�i�s�󋵃t�H�[���N���X 

    Public Declare Function GetSystemMenu Lib "user32.dll" Alias "GetSystemMenu" (ByVal hwnd As IntPtr, ByVal bRevert As Long) As IntPtr
    Public Declare Function RemoveMenu Lib "user32.dll" Alias "RemoveMenu" (ByVal hMenu As IntPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE As Long = &HF060
    Public Const MF_BYCOMMAND As Long = &H0

    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsList1, DsImp, DsCMB1, WK_DsList1 As New DataSet
    ' 2015/08/13 �d���H��ۏؒǉ��Ή�
    Dim SqlCmd2 As SqlClient.SqlCommand
    Dim DaList2 = New SqlClient.SqlDataAdapter
    Dim DsList2, DsCMB2, WK_DsList2 As New DataSet

    Dim DtView1, WK_DtView1 As DataView

    Dim dttable0, dttable1, dttable2 As DataTable
    Dim dtRow0, dtRow1, dtRow2 As DataRow

    Dim strSQL, strSQL2, Err_F, CX_F, ans, WK_str, WK_str2 As String
    Dim wk_comp, wk_plan As String
    Dim i, j, k, r As Integer
    Dim file_name, file_name2, kbn, sokatsu_kbn As String
    Dim wrn_fee2, wrn_fee3 As Integer
    Dim wrn_fee_wtax, commission_fee_wtax, admin_fee_wtax As Integer

#Region " Windows �t�H�[�� �f�U�C�i�Ő������ꂽ�R�[�h "

    Public Sub New()
        MyBase.New()

        ' ���̌Ăяo���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        InitializeComponent()

        ' InitializeComponent() �Ăяo���̌�ɏ�������ǉ����܂��B

    End Sub

    ' Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    Private components As System.ComponentModel.IContainer

    ' ���� : �ȉ��̃v���V�[�W���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    'Windows �t�H�[�� �f�U�C�i���g���ĕύX���Ă��������B  
    ' �R�[�h �G�f�B�^���g���ĕύX���Ȃ��ł��������B
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn4 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn5 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn6 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn7 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn8 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn9 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn10 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn11 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn12 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn13 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn14 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn15 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn16 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn17 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn18 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn19 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn20 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn21 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents DataGridTextBoxColumn22 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn23 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn4 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn5 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn6 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn7 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn8 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn9 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn10 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn11 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn12 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn22 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn23 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn13 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn14 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn15 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn16 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn17 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn18 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn19 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn20 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn21 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Button4 = New System.Windows.Forms.Button
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Location = New System.Drawing.Point(12, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 28)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "���ٓǍ�"
        '
        'DataGrid1
        '
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(12, 48)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.RowHeaderWidth = 10
        Me.DataGrid1.Size = New System.Drawing.Size(960, 488)
        Me.DataGrid1.TabIndex = 1
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn3, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn5, Me.DataGridTextBoxColumn6, Me.DataGridTextBoxColumn7, Me.DataGridTextBoxColumn8, Me.DataGridTextBoxColumn9, Me.DataGridTextBoxColumn10, Me.DataGridTextBoxColumn11, Me.DataGridTextBoxColumn12, Me.DataGridTextBoxColumn22, Me.DataGridTextBoxColumn23, Me.DataGridTextBoxColumn13, Me.DataGridTextBoxColumn14, Me.DataGridTextBoxColumn15, Me.DataGridTextBoxColumn16, Me.DataGridTextBoxColumn17, Me.DataGridTextBoxColumn18, Me.DataGridTextBoxColumn19, Me.DataGridTextBoxColumn20, Me.DataGridTextBoxColumn21})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "imp"
        Me.DataGridTableStyle1.RowHeaderWidth = 10
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "������"
        Me.DataGridTextBoxColumn1.MappingName = "ordr_date"
        Me.DataGridTextBoxColumn1.Width = 70
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "�ۏؔԍ�"
        Me.DataGridTextBoxColumn2.MappingName = "ordr_no"
        Me.DataGridTextBoxColumn2.Width = 80
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "���i�^��"
        Me.DataGridTextBoxColumn3.MappingName = "model_name"
        Me.DataGridTextBoxColumn3.Width = 120
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.HeaderText = "����CD"
        Me.DataGridTextBoxColumn4.MappingName = "item_cat_code"
        Me.DataGridTextBoxColumn4.Width = 70
        '
        'DataGridTextBoxColumn5
        '
        Me.DataGridTextBoxColumn5.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn5.Format = ""
        Me.DataGridTextBoxColumn5.FormatInfo = Nothing
        Me.DataGridTextBoxColumn5.HeaderText = "���[�JCD"
        Me.DataGridTextBoxColumn5.MappingName = "bend_code"
        Me.DataGridTextBoxColumn5.Width = 70
        '
        'DataGridTextBoxColumn6
        '
        Me.DataGridTextBoxColumn6.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn6.Format = ""
        Me.DataGridTextBoxColumn6.FormatInfo = Nothing
        Me.DataGridTextBoxColumn6.HeaderText = "�w�����z"
        Me.DataGridTextBoxColumn6.MappingName = "prch_price"
        Me.DataGridTextBoxColumn6.Width = 70
        '
        'DataGridTextBoxColumn7
        '
        Me.DataGridTextBoxColumn7.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn7.Format = ""
        Me.DataGridTextBoxColumn7.FormatInfo = Nothing
        Me.DataGridTextBoxColumn7.HeaderText = "�����"
        Me.DataGridTextBoxColumn7.MappingName = "prch_tax"
        Me.DataGridTextBoxColumn7.Width = 50
        '
        'DataGridTextBoxColumn8
        '
        Me.DataGridTextBoxColumn8.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn8.Format = ""
        Me.DataGridTextBoxColumn8.FormatInfo = Nothing
        Me.DataGridTextBoxColumn8.HeaderText = "�ō�"
        Me.DataGridTextBoxColumn8.MappingName = "prch_price_tax"
        Me.DataGridTextBoxColumn8.Width = 50
        '
        'DataGridTextBoxColumn9
        '
        Me.DataGridTextBoxColumn9.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn9.Format = ""
        Me.DataGridTextBoxColumn9.FormatInfo = Nothing
        Me.DataGridTextBoxColumn9.HeaderText = "����"
        Me.DataGridTextBoxColumn9.MappingName = "unit"
        Me.DataGridTextBoxColumn9.Width = 40
        '
        'DataGridTextBoxColumn10
        '
        Me.DataGridTextBoxColumn10.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn10.Format = ""
        Me.DataGridTextBoxColumn10.FormatInfo = Nothing
        Me.DataGridTextBoxColumn10.HeaderText = "�ۏؗ�"
        Me.DataGridTextBoxColumn10.MappingName = "wrn_fee"
        Me.DataGridTextBoxColumn10.Width = 50
        '
        'DataGridTextBoxColumn11
        '
        Me.DataGridTextBoxColumn11.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn11.Format = ""
        Me.DataGridTextBoxColumn11.FormatInfo = Nothing
        Me.DataGridTextBoxColumn11.HeaderText = "�����"
        Me.DataGridTextBoxColumn11.MappingName = "wrn_fee_tax"
        Me.DataGridTextBoxColumn11.Width = 50
        '
        'DataGridTextBoxColumn12
        '
        Me.DataGridTextBoxColumn12.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn12.Format = ""
        Me.DataGridTextBoxColumn12.FormatInfo = Nothing
        Me.DataGridTextBoxColumn12.HeaderText = "�ۏؔN"
        Me.DataGridTextBoxColumn12.MappingName = "wrn_prod"
        Me.DataGridTextBoxColumn12.Width = 50
        '
        'DataGridTextBoxColumn22
        '
        Me.DataGridTextBoxColumn22.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn22.Format = ""
        Me.DataGridTextBoxColumn22.FormatInfo = Nothing
        Me.DataGridTextBoxColumn22.HeaderText = "�Z�b�g"
        Me.DataGridTextBoxColumn22.MappingName = "set_flg"
        Me.DataGridTextBoxColumn22.Width = 40
        '
        'DataGridTextBoxColumn23
        '
        Me.DataGridTextBoxColumn23.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn23.Format = ""
        Me.DataGridTextBoxColumn23.FormatInfo = Nothing
        Me.DataGridTextBoxColumn23.HeaderText = "���v"
        Me.DataGridTextBoxColumn23.MappingName = "ttl"
        Me.DataGridTextBoxColumn23.Width = 60
        '
        'DataGridTextBoxColumn13
        '
        Me.DataGridTextBoxColumn13.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn13.Format = ""
        Me.DataGridTextBoxColumn13.FormatInfo = Nothing
        Me.DataGridTextBoxColumn13.HeaderText = "��"
        Me.DataGridTextBoxColumn13.MappingName = "cont_flg"
        Me.DataGridTextBoxColumn13.Width = 40
        '
        'DataGridTextBoxColumn14
        '
        Me.DataGridTextBoxColumn14.Format = ""
        Me.DataGridTextBoxColumn14.FormatInfo = Nothing
        Me.DataGridTextBoxColumn14.HeaderText = "�ڋq��"
        Me.DataGridTextBoxColumn14.MappingName = "cust_name"
        Me.DataGridTextBoxColumn14.Width = 110
        '
        'DataGridTextBoxColumn15
        '
        Me.DataGridTextBoxColumn15.Format = ""
        Me.DataGridTextBoxColumn15.FormatInfo = Nothing
        Me.DataGridTextBoxColumn15.HeaderText = "�X�֔ԍ�"
        Me.DataGridTextBoxColumn15.MappingName = "zip_code"
        Me.DataGridTextBoxColumn15.Width = 70
        '
        'DataGridTextBoxColumn16
        '
        Me.DataGridTextBoxColumn16.Format = ""
        Me.DataGridTextBoxColumn16.FormatInfo = Nothing
        Me.DataGridTextBoxColumn16.HeaderText = "�Z��"
        Me.DataGridTextBoxColumn16.MappingName = "adrs"
        Me.DataGridTextBoxColumn16.Width = 120
        '
        'DataGridTextBoxColumn17
        '
        Me.DataGridTextBoxColumn17.Format = ""
        Me.DataGridTextBoxColumn17.FormatInfo = Nothing
        Me.DataGridTextBoxColumn17.HeaderText = "�Œ�d�b"
        Me.DataGridTextBoxColumn17.MappingName = "tel1"
        Me.DataGridTextBoxColumn17.Width = 80
        '
        'DataGridTextBoxColumn18
        '
        Me.DataGridTextBoxColumn18.Format = ""
        Me.DataGridTextBoxColumn18.FormatInfo = Nothing
        Me.DataGridTextBoxColumn18.HeaderText = "�g�ѓd�b"
        Me.DataGridTextBoxColumn18.MappingName = "tel2"
        Me.DataGridTextBoxColumn18.Width = 80
        '
        'DataGridTextBoxColumn19
        '
        Me.DataGridTextBoxColumn19.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn19.Format = ""
        Me.DataGridTextBoxColumn19.FormatInfo = Nothing
        Me.DataGridTextBoxColumn19.HeaderText = "�񍐓�"
        Me.DataGridTextBoxColumn19.MappingName = "entry_date"
        Me.DataGridTextBoxColumn19.Width = 80
        '
        'DataGridTextBoxColumn20
        '
        Me.DataGridTextBoxColumn20.Format = ""
        Me.DataGridTextBoxColumn20.FormatInfo = Nothing
        Me.DataGridTextBoxColumn20.HeaderText = "JAN"
        Me.DataGridTextBoxColumn20.MappingName = "jan"
        Me.DataGridTextBoxColumn20.Width = 85
        '
        'DataGridTextBoxColumn21
        '
        Me.DataGridTextBoxColumn21.Format = ""
        Me.DataGridTextBoxColumn21.FormatInfo = Nothing
        Me.DataGridTextBoxColumn21.HeaderText = "��JAN"
        Me.DataGridTextBoxColumn21.MappingName = "moto_jan"
        Me.DataGridTextBoxColumn21.Width = 85
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(24, 548)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(228, 24)
        Me.ComboBox1.TabIndex = 2
        Me.ComboBox1.Text = "ComboBox1"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(780, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(188, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(400, 548)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "�m�@��"
        '
        'Button4
        '
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.Location = New System.Drawing.Point(872, 544)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 28)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "�߁@��"
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(264, 548)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(121, 24)
        Me.ComboBox2.TabIndex = 7
        Me.ComboBox2.Text = "ComboBox2"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(140, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(504, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Label2"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.ClientSize = New System.Drawing.Size(982, 583)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.DataGrid1)
        Me.Controls.Add(Me.Button1)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "�X�g���[���f�[�^�捞��"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '******************************************************************
    '** �N����
    '******************************************************************
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'DB_INIT()
        CmbSet()
        Button2.Enabled = False

        strSQL = "SELECT '' AS ordr_date, '' AS ordr_no, '' AS model_name, '' AS item_cat_code"
        strSQL += ", '' AS bend_code, '' AS prch_price, '' AS prch_tax, '' AS prch_price_tax"
        strSQL += ", '' AS unit, '' AS wrn_fee, '' AS wrn_fee_tax, '' AS wrn_prod, '' AS set_flg"
        strSQL += ", '' AS ttl, '' AS cont_flg, '' AS cust_name, '' AS zip_code, '' AS adrs"
        strSQL += ", '' AS tel1, '' AS tel2, '' AS entry_date, '' AS jan, '' AS moto_jan"

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(DsImp, "imp")
        DB_CLOSE()

        DsImp.Clear()

        Dim tbl1 As New DataTable
        tbl1 = DsImp.Tables("imp")
        DataGrid1.DataSource = tbl1

        '�ǂݍ��݃t�@�C����
        Label2.Text = Nothing

    End Sub

    '********************************************************************
    '** �R���{�{�b�N�X�Z�b�g
    '********************************************************************
    Sub CmbSet()
        DsCMB1.Clear()
        DB_OPEN()

        strSQL = "SELECT ltrim(rtrim(CLS_CODE)) as comp , CLS_CODE_NAME FROM CLS_CODE WHERE (CLS_NO = '001')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DaList1.Fill(DsCMB1, "cls_001")
        ComboBox1.DataSource = DsCMB1.Tables("cls_001")
        ComboBox1.DisplayMember = "CLS_CODE_NAME"
        ComboBox1.ValueMember = "comp"

        ComboBox1.Text = Nothing

        ' 2015/08/13 �d���H��ۏؒǉ��Ή� Start
        strSQL = "SELECT ltrim(rtrim(CLS_CODE)) as wplan, CLS_CODE_NAME FROM CLS_CODE WHERE (CLS_NO = '003')"
        SqlCmd2 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList2.SelectCommand = SqlCmd2
        DaList2.Fill(DsCMB1, "cls_003")
        ComboBox2.DataSource = DsCMB1.Tables("cls_003")
        ComboBox2.DisplayMember = "CLS_CODE_NAME"
        ComboBox2.ValueMember = "wplan"

        ComboBox2.Text = Nothing
        ' 2015/08/13 �d���H��ۏؒǉ��Ή� End

        DB_CLOSE()
    End Sub

    '********************************************************************
    '** ���ٓǍ�
    '********************************************************************
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        DsImp.Clear()
        Button2.Enabled = False
        Label1.Text = Nothing

        With OpenFileDialog1
            .CheckFileExists = True     '�t�@�C�������݂��邩�m�F
            .RestoreDirectory = True    '�f�B���N�g���̕���
            .ReadOnlyChecked = True
            .ShowReadOnly = True
            .Filter = "Excel ̧��(*.xls)|*.xls|���ׂẴt�@�C��(*.*)|*.*"
            .FilterIndex = 1            'Filter�v���p�e�B��2�ڂ�\��
            '�_�C�A���O�{�b�N�X��\�����A�m�J��]���N���b�N�����ꍇ
            If .ShowDialog = DialogResult.OK Then

                file_name = .FileName
                k = file_name.LastIndexOf("\")
                kbn = Mid(file_name, k + 2, 2)
                file_name2 = Mid(file_name, k + 2, 100)

                '�ǂݍ��݃t�@�C�����\��
                Label2.Text = file_name2

                Select Case kbn
                    Case Is = "E-"
                        ComboBox1.SelectedValue = "1"
                    Case Is = "EC"
                        ComboBox1.SelectedValue = "2"
                    Case Is = "����"
                        ComboBox1.SelectedValue = "3"
                    Case Is = "LA"
                        ComboBox1.SelectedValue = "4"
                    Case Else

                End Select

                Dim oExcel As Object
                Dim oBook As Object
                Dim oSheet As Object
                oExcel = CreateObject("Excel.Application")
                oBook = oExcel.Workbooks.Open(filename:=file_name)
                oSheet = oBook.Worksheets(1)

                For j = 2 To 65536

                    If oSheet.Range("A" & j).Value = Nothing Then Exit For

                    dttable0 = DsImp.Tables("imp")
                    dtRow0 = dttable0.NewRow
                    dtRow0("ordr_date") = Trim(oSheet.Range("A" & j).Value)         '������
                    dtRow0("ordr_no") = Trim(oSheet.Range("B" & j).Value)           '�ۏؔԍ�
                    dtRow0("model_name") = Trim(oSheet.Range("C" & j).Value)        '���i�^��
                    dtRow0("item_cat_code") = Trim(oSheet.Range("D" & j).Value)     '����CD
                    dtRow0("bend_code") = Trim(oSheet.Range("E" & j).Value)         '���[�JCD
                    dtRow0("prch_price") = Trim(oSheet.Range("F" & j).Value)        '�w�����z
                    dtRow0("prch_tax") = Trim(oSheet.Range("G" & j).Value)          '�w�����z(�����)
                    dtRow0("prch_price_tax") = Trim(oSheet.Range("H" & j).Value)    '�w�����z(�ō�)
                    dtRow0("unit") = Trim(oSheet.Range("I" & j).Value)              '����
                    dtRow0("wrn_fee") = Trim(oSheet.Range("J" & j).Value)           '�ۏؗ�
                    dtRow0("wrn_fee_tax") = Trim(oSheet.Range("K" & j).Value)       '�ۏؗ�(�����)
                    dtRow0("wrn_prod") = Trim(oSheet.Range("L" & j).Value)          '�����ۏؔN
                    dtRow0("set_flg") = Trim(oSheet.Range("M" & j).Value)           '�Z�b�g
                    dtRow0("ttl") = Trim(oSheet.Range("N" & j).Value)               '���v
                    dtRow0("cont_flg") = Trim(oSheet.Range("O" & j).Value)          '�X�e�[�^�X
                    dtRow0("cust_name") = Trim(oSheet.Range("P" & j).Value)         '�ڋq��
                    dtRow0("zip_code") = Trim(oSheet.Range("Q" & j).Value)          '�ڋq�X�֔ԍ�
                    dtRow0("adrs") = Trim(oSheet.Range("R" & j).Value)              '�ڋq�Z��
                    WK_str = Trim(oSheet.Range("S" & j).Value)                      '�Œ�d�b
                    If WK_str = Nothing Then
                        dtRow0("tel1") = ""
                    Else
                        If Mid(WK_str, 1, 1) = "0" Then
                            dtRow0("tel1") = WK_str
                        Else
                            dtRow0("tel1") = "0" & WK_str
                        End If
                    End If
                    WK_str = Trim(oSheet.Range("T" & j).Value)                      '�g�ѓd�b
                    If WK_str = Nothing Then
                        dtRow0("tel2") = ""
                    Else
                        If Mid(WK_str, 1, 1) = "0" Then
                            dtRow0("tel2") = WK_str
                        Else
                            dtRow0("tel2") = "0" & WK_str
                        End If
                    End If
                    dtRow0("entry_date") = Trim(oSheet.Range("U" & j).Value)        '�񍐓�
                    dtRow0("jan") = Trim(oSheet.Range("V" & j).Value)               'JAN
                    dtRow0("moto_jan") = Trim(oSheet.Range("W" & j).Value)          '��JAN
                    dttable0.Rows.Add(dtRow0)

                Next

                '==================  �I������  =====================  
                oSheet = Nothing
                oBook = Nothing
                oExcel.Quit()
                oExcel = Nothing
                GC.Collect()

                Label1.Text = Format(j - 2, "##,##0") & "��"
                If j > 2 Then

                    WK_DsList1.Clear()
                    strSQL = "SELECT file_name FROM txt_data WHERE (file_name = '" & file_name2 & "')"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    DaList1.SelectCommand = SqlCmd1
                    DB_OPEN()
                    r = DaList1.Fill(WK_DsList1, "file_name2")
                    DB_CLOSE()
                    If r <> 0 Then
                        MsgBox("���Ɏ�荞�񂾃t�@�C���ł��B")
                    Else
                        Button2.Enabled = True
                    End If
                End If
            End If
        End With

        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** DB���f�i�m��{�^���j
    '********************************************************************
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        If ComboBox1.SelectedValue = Nothing Then
            ComboBox1.Focus()
            MsgBox("�f�[�^��ނ�I�����Ă��������B", MsgBoxStyle.Critical)
            Exit Sub
        Else
            wk_comp = ComboBox1.SelectedValue
        End If

        ' 2015/08/13 �d���H��ۏؒǉ��Ή� Start
        If ComboBox2.SelectedValue = Nothing Then
            ComboBox2.Focus()
            MsgBox("�ۏ؎�ނ�I�����Ă��������B", MsgBoxStyle.Critical)
            Exit Sub
        Else
            wk_plan = ComboBox2.SelectedValue
        End If
        ' 2015/08/13 �d���H��ۏؒǉ��Ή� End

        ans = MessageBox.Show("�G�N�Z���f�[�^��ۑ����܂��B", "�m�F", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
        If ans = "2" Then Exit Sub '������

        Cursor = System.Windows.Forms.Cursors.WaitCursor

        ' �i�s�󋵃_�C�A���O�̏���������
        waitDlg = New WaitDialog        ' �i�s�󋵃_�C�A���O
        waitDlg.Owner = Me              ' �_�C�A���O�̃I�[�i�[��ݒ肷��
        waitDlg.MainMsg = "�ް����o��"  ' �����̊T�v��\��
        waitDlg.ProgressMax = 0         ' �S�̂̏���������ݒ�
        waitDlg.ProgressMin = 0         ' ���������̍ŏ��l��ݒ�i0������J�n�j
        waitDlg.ProgressStep = 1        ' �������ƂɃ��[�^��i�߂邩��ݒ�
        waitDlg.ProgressValue = 0       ' �ŏ��̌�����ݒ�
        Me.Enabled = False              ' �I�[�i�[�̃t�H�[���𖳌��ɂ���
        waitDlg.Show()                  ' �i�s�󋵃_�C�A���O��\������

        DtView1 = New DataView(DsImp.Tables("imp"), "", "", DataViewRowState.CurrentRows)

        waitDlg.MainMsg = "�ް��o�^��"      ' �i�s�󋵃_�C�A���O�̃��[�^�[��ݒ�
        waitDlg.ProgressMax = DtView1.Count ' �S�̂̏���������ݒ�
        waitDlg.ProgressValue = 0           ' �ŏ��̌�����ݒ�
        Application.DoEvents()              ' ���b�Z�[�W�����𑣂��ĕ\�����X�V����

        For j = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((j + 1) * 100 / DtView1.Count) & "%�@�i" & Format(j + 1, "##,##0") & " / " & Format(DtView1.Count, "##,##0") & " ���j"
            waitDlg.Text = "���s���E�E�E" & Fix((j + 1) * 100 / DtView1.Count) & "%"
            Application.DoEvents()  ' ���b�Z�[�W�����𑣂��ĕ\�����X�V����
            waitDlg.PerformStep()   ' �����J�E���g��1�X�e�b�v�i�߂�

            '2015/08/13 �d���H��ۏؒǉ��Ή� Start
            Select Case wk_plan
                Case Is = "10" '�Ɠd�ۏ�
                    Select Case DtView1(j)("wrn_prod")
                        '�Ɠd3�N�ۏ�
                    Case Is = "3"
                            Select Case DtView1(j)("item_cat_code")
                                'PC
                            Case Is = "101020", "101030", "101035", "101040", "101050", "102020", "102030", "102040", "102050", "102070", "103010", "103020", "108010"

                                    If CInt(DtView1(j)("prch_price_tax")) > 110000 Then
                                        sokatsu_kbn = "11" 'PC10���~��
                                        Select Case wk_comp
                                            Case Is = "1", "2", "3"
                                                'eBest�AEC�J�����g�A����COM
                                                wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)          '�̔��ۏؗ��i�ō��j
                                                commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)  '�̔��萔���i�ō��j
                                                admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)       '�����ϑ����i�ō��j

                                            Case Is = "4"
                                                'Laox
                                                wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)         '�̔��ۏؗ��i�ō��j
                                                commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)  '�̔��萔���i�ō��j
                                                admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)       '�����ϑ����i�ō��j
                                        End Select
                                    Else
                                        sokatsu_kbn = "12" 'PC10���~�ȉ�
                                        Select Case wk_comp
                                            Case Is = "1", "2", "3"
                                                'eBest�AEC�J�����g�A����COM
                                                wrn_fee_wtax = 5500         '�̔��ۏؗ��i�ō��j
                                                commission_fee_wtax = 1650  '�̔��萔���i�ō��j
                                                admin_fee_wtax = 1650       '�����ϑ����i�ō��j

                                            Case Is = "4"
                                                'Laox
                                                wrn_fee_wtax = 5000         '�̔��ۏؗ��i�ō��j
                                                commission_fee_wtax = 1000  '�̔��萔���i�ō��j
                                                admin_fee_wtax = 1000       '�����ϑ����i�ō��j
                                        End Select
                                    End If

                                    '�v�����^
                                Case Is = "301015", "301020", "301030", "301035", "301040", "301045"
                                    sokatsu_kbn = "13" '�v�����^
                                    Select Case wk_comp
                                        Case Is = "1", "2", "3"
                                            'eBest�AEC�J�����g�A����COM
                                            wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)          '�̔��ۏؗ��i�ō��j
                                            commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)  '�̔��萔���i�ō��j
                                            admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)       '�����ϑ����i�ō��j

                                        Case Is = "4"
                                            'Laox
                                            wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)         '�̔��ۏؗ��i�ō��j
                                            commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)  '�̔��萔���i�ō��j
                                            admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)       '�����ϑ����i�ō��j
                                    End Select


                                    '���̑�3�N
                                Case Else
                                    sokatsu_kbn = "14" '3�N���̑�
                                    Select Case wk_comp
                                        Case Is = "1", "2", "3"
                                            'eBest�AEC�J�����g�A����COM
                                            wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)          '�̔��ۏؗ��i�ō��j
                                            commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)  '�̔��萔���i�ō��j
                                            admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)       '�����ϑ����i�ō��j

                                        Case Is = "4"
                                            'Laox
                                            wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)         '�̔��ۏؗ��i�ō��j
                                            commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)  '�̔��萔���i�ō��j
                                            admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)       '�����ϑ����i�ō��j
                                    End Select
                            End Select

                            '�T�N�ۏ�
                        Case Is = "5"
                            sokatsu_kbn = "21" '5�N�ۏ�
                            Select Case wk_comp
                                Case Is = "1", "2", "3"
                                    If CInt(DtView1(j)("prch_price_tax")) > 11000 Then
                                        'eBest�AEC�J�����g�A����COM
                                        wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)          '�̔��ۏؗ��i�ō��j
                                        commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)  '�̔��萔���i�ō��j
                                        admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.015, 0)       '�����ϑ����i�ō��j
                                    Else
                                        wrn_fee_wtax = 550 'Sales commission (tax included)
                                        commission_fee_wtax = 165 'Sales commission (tax included)
                                        admin_fee_wtax = 165 'Administrative consignment fee (tax included)
                                    End If
                                Case Is = "4"
                                    'Laox
                                    wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.05, 0)         '�̔��ۏؗ��i�ō��j
                                    commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)  '�̔��萔���i�ō��j
                                    admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.01, 0)       '�����ϑ����i�ō��j
                            End Select

                    End Select

                Case Is = "20" '�H��ۏ�
                    If CInt(DtView1(j)("prch_price_tax")) > 16500 Then
                        sokatsu_kbn = "31" '�H��15000�~��
                        wrn_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.1, 0)           '�̔��ۏؗ��i�ō��j
                        commission_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.02, 0)   '�̔��萔���i�ō��j
                        admin_fee_wtax = RoundDOWN(CInt(DtView1(j)("prch_price_tax")) * 0.02, 0)        '�����ϑ����i�ō��j

                    Else
                        sokatsu_kbn = "32" '�H��15000�~�ȉ�
                        wrn_fee_wtax = 1650         '�̔��ۏؗ��i�ō��j
                        commission_fee_wtax = 330  '�̔��萔���i�ō��j
                        admin_fee_wtax = 330       '�����ϑ����i�ō��j
                    End If
            End Select



            strSQL = "INSERT INTO txt_data"
            strSQL += " (ordr_date, ordr_no, model_name, item_cat_code, bend_code, prch_price, prch_tax"
            strSQL += ", prch_price_tax, unit, wrn_fee, wrn_fee_tax, wrn_prod, set_flg, ttl, cont_flg, cust_name"
            strSQL += ", zip_code, adrs, tel1, tel2, entry_date, jan, moto_jan, comp"
            strSQL += ", file_name, add_F"
            strSQL += ", wrn_plan, wrn_fee_wtax, commission_fee_wtax, admin_fee_wtax, sokatsu_kbn)" '2015/08/13 �d���H��ۏؒǉ��Ή�
            strSQL += " VALUES ('" & MidB(DtView1(j)("ordr_date"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("ordr_no"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("model_name"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("item_cat_code"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("bend_code"), 1, 50) & "'"
            strSQL += ", " & DtView1(j)("prch_price")
            strSQL += ", " & DtView1(j)("prch_tax")
            strSQL += ", " & DtView1(j)("prch_price_tax")
            strSQL += ", " & DtView1(j)("unit")

            strSQL += ", " & RoundUP(wrn_fee_wtax / 1.05, 0)
            strSQL += ", " & wrn_fee_wtax - RoundUP(wrn_fee_wtax / 1.05, 0)
            strSQL += ", " & DtView1(j)("wrn_prod")
            strSQL += ", '" & DtView1(j)("set_flg") & "'"
            strSQL += ", " & DtView1(j)("ttl")
            strSQL += ", '" & MidB(DtView1(j)("cont_flg"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("cust_name"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("zip_code"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("adrs"), 1, 100) & "'"
            strSQL += ", '" & MidB(DtView1(j)("tel1"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("tel2"), 1, 50) & "'"
            strSQL += ", '" & DtView1(j)("entry_date") & "'"
            strSQL += ", '" & MidB(DtView1(j)("jan"), 1, 50) & "'"
            strSQL += ", '" & MidB(DtView1(j)("moto_jan"), 1, 50) & "'"
            strSQL += ", '" & wk_comp & "'"
            strSQL += ", '" & MidB(file_name2, 1, 50) & "'"
            strSQL += ", 0"
            strSQL += ", '" & wk_plan & "'"
            strSQL += ", " & wrn_fee_wtax
            strSQL += ", " & commission_fee_wtax
            strSQL += ", " & admin_fee_wtax
            strSQL += ", '" & sokatsu_kbn & "'"
            strSQL += ") "
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DB_OPEN()
            SqlCmd1.ExecuteNonQuery()
            DB_CLOSE()

            '2015/08/13 �d���H��ۏؒǉ��Ή� End
        Next

        MsgBox("Text�捞�ݏI��")
        DsImp.Clear()
        Button2.Enabled = False
        Label1.Text = Nothing

        waitDlg.Close()         '�i�s�󋵃_�C�A���O�����
        Me.Enabled = True       '�I�[�i�[�̃t�H�[���𖳌��ɂ���

        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** �߂�
    '********************************************************************
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DsList1.Clear()
        DsImp.Clear()
        DsCMB1.Clear()
        WK_DsList1.Clear()
        WK_DsList2.Clear()
        Me.Close()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub
End Class

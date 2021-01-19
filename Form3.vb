Imports System.ComponentModel
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form3


    Inherits System.Windows.Forms.Form
    Dim waitDlg As WaitDialog   '進行状況フォームクラス 

    Public Declare Function GetSystemMenu Lib "user32.dll" Alias "GetSystemMenu" (ByVal hwnd As IntPtr, ByVal bRevert As Long) As IntPtr
    Public Declare Function RemoveMenu Lib "user32.dll" Alias "RemoveMenu" (ByVal hMenu As IntPtr, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE As Long = &HF060
    Public Const MF_BYCOMMAND As Long = &H0

    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsList1, WK_DsList1, WK_DsList2 As New DataSet
    Dim DtView1, WK_DtView1 As DataView

    Dim dttable0, dttable1, dttable2 As DataTable
    Dim dtRow0, dtRow1, dtRow2 As DataRow

    Dim strSQL, strSQL2, Err_F, CX_F, ans, WK_str, WK_str2 As String
    Dim i, j, k, r As Integer
    Dim file_name, file_name2, kbn, dir, msg As String
    Dim wrn_fee1, wrn_fee2, wrn_fee3 As Integer

    Dim WK_comp As String

    Dim ret_HO, ret_TE, ret_RD, ret_JM As Decimal
    Dim gak_HO, gak_TE, gak_RD, gak_JM As Integer

    Dim L_ret_HO, L_ret_TE, L_ret_RD, L_ret_JM As Decimal
    Dim L_gak_HO, L_gak_TE, L_gak_RD, L_gak_JM As Integer

    '2015/08/13 電動工具保証
    Dim TLS_ret_HO, TLS_ret_TE, TLS_ret_RD, TLS_ret_JM As Decimal
    Dim TLS_gak_HO, TLS_gak_TE, TLS_gak_RD, TLS_gak_JM As Integer

#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使って変更してください。  
    ' コード エディタを使って変更しないでください。
    Friend WithEvents Button01 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button02 As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button23 As System.Windows.Forms.Button
    Friend WithEvents Button24 As System.Windows.Forms.Button
    Friend WithEvents Button22 As System.Windows.Forms.Button
    Friend WithEvents Button21 As System.Windows.Forms.Button
    Friend WithEvents Button99 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button01 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button02 = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button21 = New System.Windows.Forms.Button
        Me.Button22 = New System.Windows.Forms.Button
        Me.Button23 = New System.Windows.Forms.Button
        Me.Button24 = New System.Windows.Forms.Button
        Me.Button99 = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button01
        '
        Me.Button01.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button01.Location = New System.Drawing.Point(28, 40)
        Me.Button01.Name = "Button01"
        Me.Button01.Size = New System.Drawing.Size(100, 28)
        Me.Button01.TabIndex = 0
        Me.Button01.Text = "ＳＴＲＭ"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button01)
        Me.GroupBox1.Controls.Add(Me.Button02)
        Me.GroupBox1.Location = New System.Drawing.Point(24, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(524, 92)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "送付用 総括表 ＆ 明細"
        '
        'Button02
        '
        Me.Button02.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button02.Location = New System.Drawing.Point(152, 40)
        Me.Button02.Name = "Button02"
        Me.Button02.Size = New System.Drawing.Size(100, 28)
        Me.Button02.TabIndex = 1
        Me.Button02.Text = "Ｌａｏｘ"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button21)
        Me.GroupBox3.Controls.Add(Me.Button22)
        Me.GroupBox3.Controls.Add(Me.Button23)
        Me.GroupBox3.Controls.Add(Me.Button24)
        Me.GroupBox3.Location = New System.Drawing.Point(24, 124)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(524, 92)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "確認用"
        '
        'Button21
        '
        Me.Button21.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button21.Location = New System.Drawing.Point(28, 40)
        Me.Button21.Name = "Button21"
        Me.Button21.Size = New System.Drawing.Size(100, 28)
        Me.Button21.TabIndex = 0
        Me.Button21.Text = "ＥＣカレント"
        '
        'Button22
        '
        Me.Button22.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button22.Location = New System.Drawing.Point(152, 40)
        Me.Button22.Name = "Button22"
        Me.Button22.Size = New System.Drawing.Size(100, 28)
        Me.Button22.TabIndex = 1
        Me.Button22.Text = "Ｌａｏｘ"
        '
        'Button23
        '
        Me.Button23.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button23.Location = New System.Drawing.Point(276, 40)
        Me.Button23.Name = "Button23"
        Me.Button23.Size = New System.Drawing.Size(100, 28)
        Me.Button23.TabIndex = 2
        Me.Button23.Text = "特価COM"
        '
        'Button24
        '
        Me.Button24.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button24.Location = New System.Drawing.Point(400, 40)
        Me.Button24.Name = "Button24"
        Me.Button24.Size = New System.Drawing.Size(100, 28)
        Me.Button24.TabIndex = 3
        Me.Button24.Text = "ｅＢｅｓｔ"
        '
        'Button99
        '
        Me.Button99.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button99.Location = New System.Drawing.Point(424, 236)
        Me.Button99.Name = "Button99"
        Me.Button99.Size = New System.Drawing.Size(100, 28)
        Me.Button99.TabIndex = 3
        Me.Button99.Text = "戻　る"
        '
        'Form3
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.ClientSize = New System.Drawing.Size(574, 271)
        Me.Controls.Add(Me.Button99)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ストリーム総括表"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '******************************************************************
    '** 起動時
    '******************************************************************
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '定率設定分（PC10万円超、5年商品など）
        '　販売保証料	5.0%
        '　販売手数料	1.0%　⇒1.5%
        '　RD保証料	4.0%　⇒3.5%
        '　事務委託料	1.0%　⇒1.5%
        'ret_HO = 5 : ret_TE = 1 : ret_RD = 4 : ret_JM = 1              '2013/03まで
        ret_HO = 5.0 : ret_TE = 1.5 : ret_RD = 3.5 : ret_JM = 1.5 '2013/04から
        L_ret_HO = 5 : L_ret_TE = 1 : L_ret_RD = 4 : L_ret_JM = 1 'Laox

        '定額設定分（10万円以下PCなど）
        '　販売保証料	5,000
        '　販売手数料	1,000　⇒1,500
        '　RD保証料	4,000　⇒3,500
        '　事務委託料	1,000　⇒1,500
        'gak_HO = 5000 : gak_TE = 1000 : gak_RD = 4000 : gak_JM = 1000  '2013/03まで
        'gak_HO = 5000 : gak_TE = 1500 : gak_RD = 3500 : gak_JM = 1500 '2013/04から
        gak_HO = 5500 : gak_TE = 1650 : gak_RD = 3850 : gak_JM = 1650 '2020/12	
        L_gak_HO = 5000 : L_gak_TE = 1000 : L_gak_RD = 4000 : L_gak_JM = 1000 'Laox


        '工具定率設定分（15000円超）' 2015/08/13 電動工具保証追加対応
        '　販売保証料	10.0%
        '　販売手数料	2.0%
        '　RD保証料	    8.0%
        '　事務委託料	2.0%
        TLS_ret_HO = 10.0 : TLS_ret_TE = 2.0 : TLS_ret_RD = 8.0 : TLS_ret_JM = 2.0

        '工具定額設定分（15000円以下）' 2015/08/13 電動工具保証追加対応
        '　販売保証料	1500
        '　販売手数料	300
        '　RD保証料	    1200
        '　事務委託料	300
        'TLS_gak_HO = 1500 : TLS_gak_TE = 300 : TLS_gak_RD = 1200 : TLS_gak_JM = 300
        TLS_gak_HO = 1650 : TLS_gak_TE = 330 : TLS_gak_RD = 1320 : TLS_gak_JM = 330 '2020/12




    End Sub

    '********************************************************************
    '** 総括表(STRM)
    '********************************************************************
    Private Sub Button01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button01.Click

        '2015/8の取込実施分までは、総括表出力時に販売手数料等を計算
        '2015/9の取込実施分以降は、Excel取込時にに販売手数料等を計算し、DBへ格納
        'に変更しました。

        STRM_soukatsu()     '** 総括表(STRM)

        If CX_F = "0" Then
            WK_comp = "2"
            STRM_meisai()       '** 明細(ＥＣカレント)

            If CX_F = "0" Then
                WK_comp = "3"
                STRM_meisai()       '** 明細(特価COM)

                If CX_F = "0" Then
                    WK_comp = "1"
                    eBEST_meisai()      '** 明細(ｅＢｅｓｔ)
                End If
            End If

            MessageBox.Show(msg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Sub STRM_soukatsu()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlFilePath As String = p_dir & "\STRM総括表(YYMM).xls"
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date, entry_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max, MAX(entry_date) AS entry_max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
            WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_max"))), "yyyyMM")
        Else
            WK_str = Nothing
            DtView1 = New DataView(WK_DsList2.Tables("add_date"), "", "", DataViewRowState.CurrentRows)
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If
        End If

        '*****************************
        '** 総括表　全社計
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange1 As Excel.Range
        Dim strDat1(1, 1) As Object
        xlRange1 = xlSheet.Range("G1:G1")    'データの入力セル範囲

        ' strDat1(0, 0) = StrConv(Mid(WK_str2, 1, 4), VbStrConv.Wide) & "年" & StrConv(Mid(WK_str2, 5, 2), VbStrConv.Wide) & "月受注分"
        strDat1(0, 0) = Mid(WK_str2, 1, 4) & "年" & Mid(WK_str2, 5, 2) & "月受注分"

        xlRange1.Value = strDat1          'セルへデータの入力
        ' xlRange1.Value = strDat1(0, 0)
        MRComObject(xlRange1)            'xlRange の解放

        '*****************************
        '** 総括表　ＥＣカレント
        '*****************************
        WK_comp = "2"
        '==================  データの入力処理  ==================  
        '
        '家電保証分
        '
        xlSheet = xlSheets.Item(2)  'Sheet2
        Dim xlRange2 As Excel.Range
        Dim strDat2(5, 6) As Object     '2015/02/09 サイズ変更
        xlRange2 = xlSheet.Range("C4:H8")    'データの入力セル範囲  2015/02/09 範囲変更

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '11'"
        If WK_str = Nothing Then
            strSQL += " AND txt_data.add_date IS NULL"
        Else
            strSQL += " AND txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102)"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2(0, 0) = DtView1(0)("商品価格")
            strDat2(0, 1) = DtView1(0)("販売保証料")
            strDat2(0, 2) = DtView1(0)("販売手数料")
            strDat2(0, 3) = DtView1(0)("RD保証料")
            strDat2(0, 4) = DtView1(0)("事務委託料")
            strDat2(0, 5) = DtView1(0)("cnt")
        Else
            strDat2(0, 0) = "0"
            strDat2(0, 1) = "0"
            strDat2(0, 2) = "0"
            strDat2(0, 3) = "0"
            strDat2(0, 4) = "0"
            strDat2(0, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '12'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then

            strDat2(1, 0) = DtView1(0)("商品価格")
            strDat2(1, 1) = DtView1(0)("販売保証料")
            strDat2(1, 2) = DtView1(0)("販売手数料")
            strDat2(1, 3) = DtView1(0)("RD保証料")
            strDat2(1, 4) = DtView1(0)("事務委託料")
            strDat2(1, 5) = DtView1(0)("cnt")

        Else
            strDat2(1, 0) = "0"
            strDat2(1, 1) = "0"
            strDat2(1, 2) = "0"
            strDat2(1, 3) = "0"
            strDat2(1, 4) = "0"
            strDat2(1, 5) = "0"
        End If

        'プリンタ

        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '13'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2(2, 0) = DtView1(0)("商品価格")
            strDat2(2, 1) = DtView1(0)("販売保証料")
            strDat2(2, 2) = DtView1(0)("販売手数料")
            strDat2(2, 3) = DtView1(0)("RD保証料")
            strDat2(2, 4) = DtView1(0)("事務委託料")
            strDat2(2, 5) = DtView1(0)("cnt")
        Else
            strDat2(2, 0) = "0"
            strDat2(2, 1) = "0"
            strDat2(2, 2) = "0"
            strDat2(2, 3) = "0"
            strDat2(2, 4) = "0"
            strDat2(2, 5) = "0"
        End If


        '3年その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '14'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2(3, 0) = DtView1(0)("商品価格")
            strDat2(3, 1) = DtView1(0)("販売保証料")
            strDat2(3, 2) = DtView1(0)("販売手数料")
            strDat2(3, 3) = DtView1(0)("RD保証料")
            strDat2(3, 4) = DtView1(0)("事務委託料")
            strDat2(3, 5) = DtView1(0)("cnt")
        Else
            strDat2(3, 0) = "0"
            strDat2(3, 1) = "0"
            strDat2(3, 2) = "0"
            strDat2(3, 3) = "0"
            strDat2(3, 4) = "0"
            strDat2(3, 5) = "0"
        End If

        '5年全商品
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '21'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2(4, 0) = DtView1(0)("商品価格")
            strDat2(4, 1) = DtView1(0)("販売保証料")
            strDat2(4, 2) = DtView1(0)("販売手数料")
            strDat2(4, 3) = DtView1(0)("RD保証料")
            strDat2(4, 4) = DtView1(0)("事務委託料")
            strDat2(4, 5) = DtView1(0)("cnt")
        Else
            strDat2(4, 0) = "0"
            strDat2(4, 1) = "0"
            strDat2(4, 2) = "0"
            strDat2(4, 3) = "0"
            strDat2(4, 4) = "0"
            strDat2(4, 5) = "0"
        End If

        xlRange2.Value = strDat2          'セルへデータの入力
        MRComObject(xlRange2)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応Start
        '
        '工具保証分
        '
        Dim xlRange2t As Excel.Range
        Dim strDat2t(2, 6) As Object
        xlRange2t = xlSheet.Range("C10:H11")

        WK_DsList1.Clear()

        '工具15000円超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '31'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2t(0, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat2t(0, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat2t(0, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat2t(0, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat2t(0, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat2t(0, 5) = DtView1(0)("cnt")                                    '件数
        Else
            strDat2t(0, 0) = "0"
            strDat2t(0, 1) = "0"
            strDat2t(0, 2) = "0"
            strDat2t(0, 3) = "0"
            strDat2t(0, 4) = "0"
            strDat2t(0, 5) = "0"
        End If

        '電動工具　15000円以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '32'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat2t(1, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat2t(1, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat2t(1, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat2t(1, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat2t(1, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat2t(1, 5) = DtView1(0)("cnt")                                    '件数

        Else
            strDat2t(1, 0) = "0"
            strDat2t(1, 1) = "0"
            strDat2t(1, 2) = "0"
            strDat2t(1, 3) = "0"
            strDat2t(1, 4) = "0"
            strDat2t(1, 5) = "0"
        End If

        xlRange2t.Value = strDat2t          'セルへデータの入力
        MRComObject(xlRange2t)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応End

        '*****************************
        '** 総括表　eBest
        '*****************************
        WK_comp = "1"
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(3)  'Sheet3
        Dim xlRange3 As Excel.Range
        Dim strDat3(5, 6) As Object     '2015/02/09 サイズ変更
        xlRange3 = xlSheet.Range("C4:H8")    'データの入力セル範囲  2015/02/09 範囲変更

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '11'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3(0, 0) = DtView1(0)("商品価格")
            strDat3(0, 1) = DtView1(0)("販売保証料")
            strDat3(0, 2) = DtView1(0)("販売手数料")
            strDat3(0, 3) = DtView1(0)("RD保証料")
            strDat3(0, 4) = DtView1(0)("事務委託料")
            strDat3(0, 5) = DtView1(0)("cnt")
        Else
            strDat3(0, 0) = "0"
            strDat3(0, 1) = "0"
            strDat3(0, 2) = "0"
            strDat3(0, 3) = "0"
            strDat3(0, 4) = "0"
            strDat3(0, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '12'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3(1, 0) = DtView1(0)("商品価格")
            strDat3(1, 1) = DtView1(0)("販売保証料")
            strDat3(1, 2) = DtView1(0)("販売手数料")
            strDat3(1, 3) = DtView1(0)("RD保証料")
            strDat3(1, 4) = DtView1(0)("事務委託料")
            strDat3(1, 5) = DtView1(0)("cnt")
        Else
            strDat3(1, 0) = "0"
            strDat3(1, 1) = "0"
            strDat3(1, 2) = "0"
            strDat3(1, 3) = "0"
            strDat3(1, 4) = "0"
            strDat3(1, 5) = "0"
        End If

        'プリンタ
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '13'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3(2, 0) = DtView1(0)("商品価格")
            strDat3(2, 1) = DtView1(0)("販売保証料")
            strDat3(2, 2) = DtView1(0)("販売手数料")
            strDat3(2, 3) = DtView1(0)("RD保証料")
            strDat3(2, 4) = DtView1(0)("事務委託料")
            strDat3(2, 5) = DtView1(0)("cnt")
        Else
            strDat3(2, 0) = "0"
            strDat3(2, 1) = "0"
            strDat3(2, 2) = "0"
            strDat3(2, 3) = "0"
            strDat3(2, 4) = "0"
            strDat3(2, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD START 
        '3年保証その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '14'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3(3, 0) = DtView1(0)("商品価格")
            strDat3(3, 1) = DtView1(0)("販売保証料")
            strDat3(3, 2) = DtView1(0)("販売手数料")
            strDat3(3, 3) = DtView1(0)("RD保証料")
            strDat3(3, 4) = DtView1(0)("事務委託料")
            strDat3(3, 5) = DtView1(0)("cnt")
        Else
            strDat3(3, 0) = "0"
            strDat3(3, 1) = "0"
            strDat3(3, 2) = "0"
            strDat3(3, 3) = "0"
            strDat3(3, 4) = "0"
            strDat3(3, 5) = "0"
        End If
        '2015/02/09 3年保証商品　20120620追加対応 ADD END 

        '5年全商品
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '21'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3(4, 0) = DtView1(0)("商品価格")
            strDat3(4, 1) = DtView1(0)("販売保証料")
            strDat3(4, 2) = DtView1(0)("販売手数料")
            strDat3(4, 3) = DtView1(0)("RD保証料")
            strDat3(4, 4) = DtView1(0)("事務委託料")
            strDat3(4, 5) = DtView1(0)("cnt")
        Else
            strDat3(4, 0) = "0"
            strDat3(4, 1) = "0"
            strDat3(4, 2) = "0"
            strDat3(4, 3) = "0"
            strDat3(4, 4) = "0"
            strDat3(4, 5) = "0"
        End If

        xlRange3.Value = strDat3          'セルへデータの入力
        MRComObject(xlRange3)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応Start
        '
        '工具保証分
        '
        Dim xlRange3t As Excel.Range
        Dim strDat3t(2, 6) As Object
        xlRange3t = xlSheet.Range("C10:H11")

        WK_DsList1.Clear()

        '工具15000円超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '31'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat3t(0, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat3t(0, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat3t(0, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat3t(0, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat3t(0, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat3t(0, 5) = DtView1(0)("cnt")                                    '件数
        Else
            strDat3t(0, 0) = "0"
            strDat3t(0, 1) = "0"
            strDat3t(0, 2) = "0"
            strDat3t(0, 3) = "0"
            strDat3t(0, 4) = "0"
            strDat3t(0, 5) = "0"
        End If

        '電動工具　15000円以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '32'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then

            strDat3t(1, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat3t(1, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat3t(1, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat3t(1, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat3t(1, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat3t(1, 5) = DtView1(0)("cnt")                                    '件数

        Else
            strDat3t(1, 0) = "0"
            strDat3t(1, 1) = "0"
            strDat3t(1, 2) = "0"
            strDat3t(1, 3) = "0"
            strDat3t(1, 4) = "0"
            strDat3t(1, 5) = "0"
        End If

        xlRange3t.Value = strDat3t          'セルへデータの入力
        MRComObject(xlRange3t)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応End

        '*****************************
        '** 総括表　特価COM
        '*****************************
        WK_comp = "3"
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(4)  'Sheet4
        Dim xlRange4 As Excel.Range
        Dim strDat4(5, 6) As Object     '2015/02/09 サイズ変更
        xlRange4 = xlSheet.Range("C4:H8")    'データの入力セル範囲  2015/02/09 範囲変更

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '11'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4(0, 0) = DtView1(0)("商品価格")
            strDat4(0, 1) = DtView1(0)("販売保証料")
            strDat4(0, 2) = DtView1(0)("販売手数料")
            strDat4(0, 3) = DtView1(0)("RD保証料")
            strDat4(0, 4) = DtView1(0)("事務委託料")
            strDat4(0, 5) = DtView1(0)("cnt")
        Else
            strDat4(0, 0) = "0"
            strDat4(0, 1) = "0"
            strDat4(0, 2) = "0"
            strDat4(0, 3) = "0"
            strDat4(0, 4) = "0"
            strDat4(0, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '12'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4(1, 0) = DtView1(0)("商品価格")
            strDat4(1, 1) = DtView1(0)("販売保証料")
            strDat4(1, 2) = DtView1(0)("販売手数料")
            strDat4(1, 3) = DtView1(0)("RD保証料")
            strDat4(1, 4) = DtView1(0)("事務委託料")
            strDat4(1, 5) = DtView1(0)("cnt")
        Else
            strDat4(1, 0) = "0"
            strDat4(1, 1) = "0"
            strDat4(1, 2) = "0"
            strDat4(1, 3) = "0"
            strDat4(1, 4) = "0"
            strDat4(1, 5) = "0"
        End If

        'プリンタ
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '13'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4(2, 0) = DtView1(0)("商品価格")
            strDat4(2, 1) = DtView1(0)("販売保証料")
            strDat4(2, 2) = DtView1(0)("販売手数料")
            strDat4(2, 3) = DtView1(0)("RD保証料")
            strDat4(2, 4) = DtView1(0)("事務委託料")
            strDat4(2, 5) = DtView1(0)("cnt")
        Else
            strDat4(2, 0) = "0"
            strDat4(2, 1) = "0"
            strDat4(2, 2) = "0"
            strDat4(2, 3) = "0"
            strDat4(2, 4) = "0"
            strDat4(2, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD START 
        '3年保証その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '14'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4(3, 0) = DtView1(0)("商品価格")
            strDat4(3, 1) = DtView1(0)("販売保証料")
            strDat4(3, 2) = DtView1(0)("販売手数料")
            strDat4(3, 3) = DtView1(0)("RD保証料")
            strDat4(3, 4) = DtView1(0)("事務委託料")
            strDat4(3, 5) = DtView1(0)("cnt")
        Else
            strDat4(3, 0) = "0"
            strDat4(3, 1) = "0"
            strDat4(3, 2) = "0"
            strDat4(3, 3) = "0"
            strDat4(3, 4) = "0"
            strDat4(3, 5) = "0"
        End If
        '2015/02/09 3年保証商品　20120620追加対応 ADD END 

        '5年全商品
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '21'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4(4, 0) = DtView1(0)("商品価格")
            strDat4(4, 1) = DtView1(0)("販売保証料")
            strDat4(4, 2) = DtView1(0)("販売手数料")
            strDat4(4, 3) = DtView1(0)("RD保証料")
            strDat4(4, 4) = DtView1(0)("事務委託料")
            strDat4(4, 5) = DtView1(0)("cnt")
        Else
            strDat4(4, 0) = "0"
            strDat4(4, 1) = "0"
            strDat4(4, 2) = "0"
            strDat4(4, 3) = "0"
            strDat4(4, 4) = "0"
            strDat4(4, 5) = "0"
        End If

        xlRange4.Value = strDat4          'セルへデータの入力
        MRComObject(xlRange4)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応Start
        '
        '工具保証分
        '
        Dim xlRange4t As Excel.Range
        Dim strDat4t(2, 6) As Object
        xlRange4t = xlSheet.Range("C10:H11")

        WK_DsList1.Clear()

        '工具15000円超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '31'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4t(0, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat4t(0, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat4t(0, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat4t(0, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat4t(0, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat4t(0, 5) = DtView1(0)("cnt")                                    '件数
        Else
            strDat4t(0, 0) = "0"
            strDat4t(0, 1) = "0"
            strDat4t(0, 2) = "0"
            strDat4t(0, 3) = "0"
            strDat4t(0, 4) = "0"
            strDat4t(0, 5) = "0"
        End If

        '電動工具　15000円以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '32'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "TOOLS2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("TOOLS2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat4t(1, 0) = DtView1(0)("商品価格")                               '商品価格（税込）
            strDat4t(1, 1) = DtView1(0)("販売保証料")                             '販売保証料（税込）
            strDat4t(1, 2) = DtView1(0)("販売手数料")                             '販売手数料（税込）
            strDat4t(1, 3) = DtView1(0)("RD保証料")                               'RD保証料（税込）
            strDat4t(1, 4) = DtView1(0)("事務委託料")                             '事務委託料（税込）
            strDat4t(1, 5) = DtView1(0)("cnt")                                    '件数
        Else
            strDat4t(1, 0) = "0"
            strDat4t(1, 1) = "0"
            strDat4t(1, 2) = "0"
            strDat4t(1, 3) = "0"
            strDat4t(1, 4) = "0"
            strDat4t(1, 5) = "0"
        End If

        xlRange4t.Value = strDat4t          'セルへデータの入力
        MRComObject(xlRange4t)            'xlRange の解放

        '2015/08/13 電動工具保証追加対応End



        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        SaveFileDialog1.FileName = "STRM総括表(" & Mid(WK_str2, 3, 4) & ").xls"
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' dir = SaveFileDialog1.FileName
            'dir = Mid(SaveFileDialog1.FileName, 1, dir.LastIndexOf("\") + 1)

            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            msg = SaveFileDialog1.FileName & " に出力しました。"
            'MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        If Err.Number <> 0 Then
            MessageBox.Show(Err.Description)
            Err.Clear()
        End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** 総括表(Laox)
    '********************************************************************
    Private Sub Button02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button02.Click

        Laox_soukatsu()     '** 総括表(Laox)

        If CX_F = "0" Then
            WK_comp = "4"
            STRM_meisai()       '** 明細(Laox)

            MessageBox.Show(msg, "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Sub Laox_soukatsu()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlFilePath As String = p_dir & "\Laox総括表(YYMM).xls"
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date, entry_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max, MAX(entry_date) AS entry_max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
            WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_max"))), "yyyyMM")
        Else
            WK_str = Nothing
            DtView1 = New DataView(WK_DsList2.Tables("add_date"), "", "", DataViewRowState.CurrentRows)
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If
        End If

        '*****************************
        '** 総括表　受注月
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange0 As Excel.Range
        Dim strDat0(1, 1) As Object
        xlRange0 = xlSheet.Range("G1:G1")    'データの入力セル範囲

        '   strDat0(0, 0) = StrConv(Mid(WK_str2, 1, 4), VbStrConv.Wide) & "年" & StrConv(Mid(WK_str2, 5, 2), VbStrConv.Wide) & "月受注分"
        strDat0(0, 0) = Mid(WK_str2, 1, 4) & "年" & Mid(WK_str2, 5, 2) & "月受注分"

        xlRange0.Value = strDat0          'セルへデータの入力
        MRComObject(xlRange0)            'xlRange の解放

        '*****************************
        '** 総括表
        '*****************************
        WK_comp = "4"
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange1 As Excel.Range
        Dim strDat1(5, 6) As Object     '2015/02/09 サイズ変更
        xlRange1 = xlSheet.Range("C4:H8")    'データの入力セル範囲  2015/02/09 範囲変更

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '11'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat1(0, 0) = DtView1(0)("商品価格")
            strDat1(0, 1) = DtView1(0)("販売保証料")
            strDat1(0, 2) = DtView1(0)("販売手数料")
            strDat1(0, 3) = DtView1(0)("RD保証料")
            strDat1(0, 4) = DtView1(0)("事務委託料")
            strDat1(0, 5) = DtView1(0)("cnt")
        Else
            strDat1(0, 0) = "0"
            strDat1(0, 1) = "0"
            strDat1(0, 2) = "0"
            strDat1(0, 3) = "0"
            strDat1(0, 4) = "0"
            strDat1(0, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '12'"

        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat1(1, 0) = DtView1(0)("商品価格")
            strDat1(1, 1) = DtView1(0)("販売保証料")
            strDat1(1, 2) = DtView1(0)("販売手数料")
            strDat1(1, 3) = DtView1(0)("RD保証料")
            strDat1(1, 4) = DtView1(0)("事務委託料")
            strDat1(1, 5) = DtView1(0)("cnt")
        Else
            strDat1(1, 0) = "0"
            strDat1(1, 1) = "0"
            strDat1(1, 2) = "0"
            strDat1(1, 3) = "0"
            strDat1(1, 4) = "0"
            strDat1(1, 5) = "0"
        End If

        'プリンタ
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '13'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat1(2, 0) = DtView1(0)("商品価格")
            strDat1(2, 1) = DtView1(0)("販売保証料")
            strDat1(2, 2) = DtView1(0)("販売手数料")
            strDat1(2, 3) = DtView1(0)("RD保証料")
            strDat1(2, 4) = DtView1(0)("事務委託料")
            strDat1(2, 5) = DtView1(0)("cnt")
        Else
            strDat1(2, 0) = "0"
            strDat1(2, 1) = "0"
            strDat1(2, 2) = "0"
            strDat1(2, 3) = "0"
            strDat1(2, 4) = "0"
            strDat1(2, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD START 
        '3年保証その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '14'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat1(3, 0) = DtView1(0)("商品価格")
            strDat1(3, 1) = DtView1(0)("販売保証料")
            strDat1(3, 2) = DtView1(0)("販売手数料")
            strDat1(3, 3) = DtView1(0)("RD保証料")
            strDat1(3, 4) = DtView1(0)("事務委託料")
            strDat1(3, 5) = DtView1(0)("cnt")
        Else
            strDat1(3, 0) = "0"
            strDat1(3, 1) = "0"
            strDat1(3, 2) = "0"
            strDat1(3, 3) = "0"
            strDat1(3, 4) = "0"
            strDat1(3, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD END 
        '5年全商品
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE txt_data.comp = '" & WK_comp & "'"
        strSQL += " AND txt_data.sokatsu_kbn = '21'"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If

        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat1(4, 0) = DtView1(0)("商品価格")
            strDat1(4, 1) = DtView1(0)("販売保証料")
            strDat1(4, 2) = DtView1(0)("販売手数料")
            strDat1(4, 3) = DtView1(0)("RD保証料")
            strDat1(4, 4) = DtView1(0)("事務委託料")
            strDat1(4, 5) = DtView1(0)("cnt")
        Else
            strDat1(4, 0) = "0"
            strDat1(4, 1) = "0"
            strDat1(4, 2) = "0"
            strDat1(4, 3) = "0"
            strDat1(4, 4) = "0"
            strDat1(4, 5) = "0"
        End If

        xlRange1.Value = strDat1          'セルへデータの入力
        MRComObject(xlRange1)            'xlRange の解放

        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        SaveFileDialog1.FileName = "Laox総括表(" & Mid(WK_str2, 3, 4) & ").xls"
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            dir = SaveFileDialog1.FileName
            dir = Mid(SaveFileDialog1.FileName, 1, dir.LastIndexOf("\") + 1)

            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            msg = SaveFileDialog1.FileName & " に出力しました。"
            'MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        'If Err.Number <> 0 Then
        '    MessageBox.Show(Err.Description)
        '    Err.Clear()
        'End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Sub STRM_meisai()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlsfile As String
        Select Case WK_comp
            Case Is = "2"
                xlsfile = "\STRM明細(YYMM).xls"
            Case Is = "3"
                xlsfile = "\特価COM明細(YYMM).xls"
            Case Is = "4"
                xlsfile = "\Laox明細(YYMM).xls"
        End Select
        Dim xlFilePath As String = p_dir & xlsfile
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
        Else
            WK_str = Nothing
        End If

        '*****************************
        '** 明細
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange2 As Excel.Range
        Dim strDat2(9999, 16) As Object
        xlRange2 = xlSheet.Range("A2:P10000")    'データの入力セル範囲
        Dim xlCells2 As Excel.Range
        Dim xlRange2_2 As Excel.Range

        '明細
        WK_DsList1.Clear()
        strSQL = "SELECT txt_data.ordr_date, txt_data.ordr_no, txt_data.model_name, txt_data.item_cat_code"
        strSQL += ", txt_data.bend_code, txt_data.prch_price_tax, txt_data.wrn_fee + txt_data.wrn_fee_tax AS wrn_fee"
        strSQL += ", txt_data.wrn_prod, txt_data.cont_flg, txt_data.cust_name, txt_data.zip_code, txt_data.adrs"
        strSQL += ", txt_data.tel1, txt_data.tel2"

        strSQL += ", txt_data.wrn_fee_wtax AS 販売保証料"
        strSQL += ", txt_data.commission_fee_wtax AS 販売手数料"
        strSQL += ", (txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", txt_data.admin_fee_wtax AS 事務委託料"
        strSQL += ", txt_data.sokatsu_kbn AS 集計区分"
        strSQL += ", txt_data.entry_date, txt_data.set_flg, txt_data.ttl"
        strSQL += " FROM txt_data"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "dtl")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("dtl"), "", "ordr_no", DataViewRowState.CurrentRows)
        If DtView1.Count <> 0 Then
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If

            If WK_comp = "4" Then   'Laox
                For i = 0 To DtView1.Count - 1
                    j = i
                    strDat2(j, 0) = j + 1
                    strDat2(j, 1) = DtView1(i)("ordr_date")
                    strDat2(j, 2) = DtView1(i)("ordr_no")
                    strDat2(j, 3) = DtView1(i)("model_name")
                    strDat2(j, 4) = DtView1(i)("item_cat_code")
                    strDat2(j, 5) = DtView1(i)("prch_price_tax")
                    strDat2(j, 6) = DtView1(i)("wrn_prod")
                    strDat2(j, 7) = DtView1(i)("cust_name")
                    strDat2(j, 8) = DtView1(i)("販売保証料")
                    strDat2(j, 9) = DtView1(i)("販売手数料")
                    strDat2(j, 10) = DtView1(i)("RD保証料")
                    strDat2(j, 11) = DtView1(i)("事務委託料")
                    strDat2(j, 12) = DtView1(i)("集計区分")

                Next
                j = j + 1
                strDat2(j, 5) = "=SUM(F2:F" & j + 1 & ")"
                strDat2(j, 8) = "=SUM(I2:I" & j + 1 & ")"
                strDat2(j, 9) = "=SUM(J2:J" & j + 1 & ")"
                strDat2(j, 10) = "=SUM(K2:K" & j + 1 & ")"
            Else
                For i = 0 To DtView1.Count - 1
                    j = i
                    strDat2(j, 0) = j + 1
                    strDat2(j, 1) = DtView1(i)("ordr_date")
                    strDat2(j, 2) = DtView1(i)("ordr_no")
                    strDat2(j, 3) = DtView1(i)("model_name")
                    strDat2(j, 4) = DtView1(i)("item_cat_code")
                    strDat2(j, 5) = DtView1(i)("wrn_prod")
                    strDat2(j, 6) = DtView1(i)("cust_name")
                    strDat2(j, 7) = DtView1(i)("prch_price_tax")
                    strDat2(j, 8) = DtView1(i)("販売保証料")
                    strDat2(j, 9) = DtView1(i)("set_flg")
                    strDat2(j, 10) = DtView1(i)("ttl")
                    strDat2(j, 11) = DtView1(i)("販売手数料")
                    strDat2(j, 12) = DtView1(i)("RD保証料")
                    strDat2(j, 13) = DtView1(i)("事務委託料")
                    strDat2(j, 14) = DtView1(i)("集計区分")

                Next
                j = j + 1
                'strDat2(j, 5) = "=SUM(F2:F" & j + 1 & ")"
                strDat2(j, 8) = "=SUM(I2:I" & j + 1 & ")"
                'strDat2(j, 9) = "=SUM(J2:J" & j + 1 & ")"
                strDat2(j, 10) = "=SUM(K2:K" & j + 1 & ")"
                strDat2(j, 11) = "=SUM(L2:L" & j + 1 & ")"
                strDat2(j, 12) = "=SUM(M2:M" & j + 1 & ")"
            End If

            xlRange2.Value = strDat2            'セルへデータの入力
            MRComObject(xlRange2)               'xlRange の解放

        End If

        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        Select Case WK_comp
            Case Is = "2"
                SaveFileDialog1.FileName = dir & "STRM明細(" & Mid(WK_str2, 3, 4) & ").xls"
            Case Is = "3"
                SaveFileDialog1.FileName = dir & "特価COM明細(" & Mid(WK_str2, 3, 4) & ").xls"
            Case Is = "4"
                SaveFileDialog1.FileName = dir & "Laox明細(" & Mid(WK_str2, 3, 4) & ").xls"
        End Select
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            msg = msg & vbCrLf & SaveFileDialog1.FileName & " に出力しました。"
            'MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        If Err.Number <> 0 Then
            MessageBox.Show(Err.Description)
            Err.Clear()
        End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Sub eBEST_meisai()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlFilePath As String = p_dir & "\eBest明細(YYMM).xls"
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
        Else
            WK_str = Nothing
        End If

        '*****************************
        '** 明細
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange2 As Excel.Range
        Dim strDat2(9999, 15) As Object
        xlRange2 = xlSheet.Range("A2:P10000")    'データの入力セル範囲
        Dim xlCells2 As Excel.Range
        Dim xlRange2_2 As Excel.Range

        '明細
        WK_DsList1.Clear()
        strSQL = "SELECT txt_data.ordr_date, txt_data.ordr_no, txt_data.model_name, txt_data.item_cat_code, txt_data.bend_code"
        strSQL += ", txt_data.prch_price, txt_data.prch_tax, txt_data.wrn_fee, txt_data.wrn_fee_tax, txt_data.wrn_prod"
        strSQL += ", txt_data.cont_flg, txt_data.cust_name, txt_data.zip_code, txt_data.adrs, txt_data.tel1"
        strSQL += ", txt_data.tel2, txt_data.prch_price_tax"

        strSQL += ", txt_data.wrn_fee_wtax AS 販売保証料"
        strSQL += ", txt_data.commission_fee_wtax AS 販売手数料"
        strSQL += ", (txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", txt_data.admin_fee_wtax AS 事務委託料"
        strSQL += ", txt_data.sokatsu_kbn AS 集計区分"

        strSQL += ", txt_data.entry_date, txt_data.set_flg, txt_data.ttl"
        strSQL += " FROM txt_data"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "dtl")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("dtl"), "", "ordr_no", DataViewRowState.CurrentRows)
        If DtView1.Count <> 0 Then
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If

            For i = 0 To DtView1.Count - 1
                j = i
                strDat2(j, 0) = j + 1
                strDat2(j, 1) = Mid(DtView1(i)("ordr_date"), 1, 4) & "/" & Mid(DtView1(i)("ordr_date"), 5, 2) & "/" & Mid(DtView1(i)("ordr_date"), 7, 2)
                strDat2(j, 2) = DtView1(i)("ordr_no")
                strDat2(j, 3) = DtView1(i)("model_name")
                strDat2(j, 4) = DtView1(i)("item_cat_code")
                strDat2(j, 5) = DtView1(i)("wrn_prod")
                strDat2(j, 6) = DtView1(i)("cust_name")
                strDat2(j, 7) = DtView1(i)("prch_price_tax")

                strDat2(j, 8) = DtView1(i)("販売保証料")
                strDat2(j, 9) = DtView1(i)("set_flg")
                strDat2(j, 10) = DtView1(i)("ttl")
                strDat2(j, 11) = DtView1(i)("販売手数料")
                strDat2(j, 12) = DtView1(i)("RD保証料")
                strDat2(j, 13) = DtView1(i)("事務委託料")
                strDat2(j, 14) = DtView1(i)("集計区分")


            Next
            j = j + 1
            strDat2(j, 7) = "=SUM(H2:H" & j + 1 & ")"
            strDat2(j, 8) = "=SUM(I2:I" & j + 1 & ")"
            'strDat2(j, 9) = "=SUM(J2:J" & j + 1 & ")"
            strDat2(j, 10) = "=SUM(K2:K" & j + 1 & ")"
            strDat2(j, 11) = "=SUM(L2:L" & j + 1 & ")"
            strDat2(j, 12) = "=SUM(M2:M" & j + 1 & ")"

            xlRange2.Value = strDat2            'セルへデータの入力
            MRComObject(xlRange2)               'xlRange の解放

        End If

        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        SaveFileDialog1.FileName = dir & "eBest明細(" & Mid(WK_str2, 3, 4) & ").xls"
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            msg = msg & vbCrLf & SaveFileDialog1.FileName & " に出力しました。"
            'MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        If Err.Number <> 0 Then
            MessageBox.Show(Err.Description)
            Err.Clear()
        End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** 確認用(ＥＣカレント)
    '********************************************************************
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        WK_comp = "2"
        STRM_kakunin()
    End Sub

    '********************************************************************
    '** 確認用(Laox)
    '********************************************************************
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        WK_comp = "4"
        STRM_kakunin()
    End Sub

    '********************************************************************
    '** 確認用(特価COM)
    '********************************************************************
    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        WK_comp = "3"
        STRM_kakunin()
    End Sub

    '********************************************************************
    '** 確認用(ｅＢｅｓｔ)
    '********************************************************************
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        WK_comp = "1"
        eBEST_kakunin()
    End Sub

    Sub STRM_kakunin()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlFilePath As String = p_dir & "\STRM総括表.xls"
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
        Else
            WK_str = Nothing
        End If

        '*****************************
        '** 総括表
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange As Excel.Range
        Dim strDat(5, 6) As Object     '2015/02/09 サイズ変更
        xlRange = xlSheet.Range("C3:H7")    'データの入力セル範囲  2015/02/09 範囲変更

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        If WK_comp = "4" Then
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        Else
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        End If
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'PC')"
        strSQL += " AND (txt_data.prch_price_tax > '110000')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(0, 0) = DtView1(0)("商品価格")
            strDat(0, 1) = DtView1(0)("販売保証料")
            strDat(0, 2) = DtView1(0)("販売手数料")
            strDat(0, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(0, 4) = DtView1(0)("事務委託料")
            strDat(0, 5) = DtView1(0)("cnt")
        Else
            strDat(0, 0) = "0"
            strDat(0, 1) = "0"
            strDat(0, 2) = "0"
            strDat(0, 3) = "0"
            strDat(0, 4) = "0"
            strDat(0, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'PC')"
        strSQL += " AND (txt_data.prch_price_tax <= '110000')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(1, 0) = DtView1(0)("商品価格")
            If WK_comp = "4" Then
                strDat(1, 1) = DtView1(0)("cnt") * L_gak_HO
                strDat(1, 2) = DtView1(0)("cnt") * L_gak_TE
                strDat(1, 3) = DtView1(0)("cnt") * L_gak_RD
                strDat(1, 4) = DtView1(0)("cnt") * L_gak_JM
            Else
                'strDat(1, 1) = DtView1(0)("cnt") * gak_HO
                'strDat(1, 2) = DtView1(0)("cnt") * gak_TE
                'strDat(1, 3) = DtView1(0)("cnt") * gak_RD
                'strDat(1, 4) = DtView1(0)("cnt") * gak_JM
                strDat(1, 1) = DtView1(0)("販売保証料")
                strDat(1, 2) = DtView1(0)("販売手数料")
                strDat(1, 3) = DtView1(0)("RD保証料")
                strDat(1, 4) = DtView1(0)("事務委託料")

            End If
            strDat(1, 5) = DtView1(0)("cnt")
        Else
            strDat(1, 0) = "0"
            strDat(1, 1) = "0"
            strDat(1, 2) = "0"
            strDat(1, 3) = "0"
            strDat(1, 4) = "0"
            strDat(1, 5) = "0"
        End If

        'プリンタ
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        If WK_comp = "4" Then
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        Else
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        End If
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'プリンタ')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(2, 0) = DtView1(0)("商品価格")
            strDat(2, 1) = DtView1(0)("販売保証料")
            strDat(2, 2) = DtView1(0)("販売手数料")
            strDat(2, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(2, 4) = DtView1(0)("事務委託料")
            strDat(2, 5) = DtView1(0)("cnt")
        Else
            strDat(2, 0) = "0"
            strDat(2, 1) = "0"
            strDat(2, 2) = "0"
            strDat(2, 3) = "0"
            strDat(2, 4) = "0"
            strDat(2, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD START 
        '3年保証その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        If WK_comp = "4" Then
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & L_ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        Else
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        End If
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (RTRIM(V_cls_002.CLS_CODE) IN ('7068', '7515', '7518', '7545')) "
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(3, 0) = DtView1(0)("商品価格")
            strDat(3, 1) = DtView1(0)("販売保証料")
            strDat(3, 2) = DtView1(0)("販売手数料")
            strDat(3, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(3, 4) = DtView1(0)("事務委託料")
            strDat(3, 5) = DtView1(0)("cnt")
        Else
            strDat(3, 0) = "0"
            strDat(3, 1) = "0"
            strDat(3, 2) = "0"
            strDat(3, 3) = "0"
            strDat(3, 4) = "0"
            strDat(3, 5) = "0"
        End If
        '2015/02/09 3年保証商品　20120620追加対応 ADD END 

        '5年全商品
        strSQL = "SELECT SUM(prch_price_tax) AS 商品価格"
        If WK_comp = "4" Then
            strSQL += ", SUM(ROUND(prch_price_tax * " & L_ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(prch_price_tax * " & L_ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(prch_price_tax * " & L_ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(prch_price_tax * " & L_ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        Else
            strSQL += ", SUM(ROUND(prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
            strSQL += ", SUM(ROUND(prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
            'strSQL +=  ", SUM(ROUND(prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
            strSQL += ", SUM(ROUND(prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        End If
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE (comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (wrn_prod = 5)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(4, 0) = DtView1(0)("商品価格")
            strDat(4, 1) = DtView1(0)("販売保証料")
            strDat(4, 2) = DtView1(0)("販売手数料")
            strDat(4, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(4, 4) = DtView1(0)("事務委託料")
            strDat(4, 5) = DtView1(0)("cnt")
        Else
            strDat(4, 0) = "0"
            strDat(4, 1) = "0"
            strDat(4, 2) = "0"
            strDat(4, 3) = "0"
            strDat(4, 4) = "0"
            strDat(4, 5) = "0"
        End If

        xlRange.Value = strDat          'セルへデータの入力
        MRComObject(xlRange)            'xlRange の解放

        '*****************************
        '** 明細
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(2)  'Sheet2
        Dim xlRange2 As Excel.Range
        Dim strDat2(9999, 20) As Object
        xlRange2 = xlSheet.Range("A2:T10000")    'データの入力セル範囲
        Dim xlCells2 As Excel.Range
        Dim xlRange2_2 As Excel.Range

        '明細
        strSQL = "SELECT txt_data.ordr_date, txt_data.ordr_no, txt_data.model_name, txt_data.item_cat_code"
        strSQL += ", txt_data.bend_code, txt_data.prch_price_tax, txt_data.wrn_fee + txt_data.wrn_fee_tax AS wrn_fee"
        strSQL += ", txt_data.wrn_prod, txt_data.cont_flg, txt_data.cust_name, txt_data.zip_code, txt_data.adrs"
        strSQL += ", txt_data.tel1, txt_data.tel2, V_cls_002.CLS_CODE_NAME"
        If WK_comp = "4" Then
            strSQL += ", ROUND(txt_data.prch_price_tax * " & L_ret_HO & " / 100, 0, - 1) AS 販売保証料"
            strSQL += ", ROUND(txt_data.prch_price_tax * " & L_ret_TE & " / 100, 0, - 1) AS 販売手数料"
            'strSQL +=  ", ROUND(txt_data.prch_price_tax * " & L_ret_RD & " / 100, 0, - 1) AS ＲＤ保証料"
            strSQL += ", ROUND(txt_data.prch_price_tax * " & L_ret_JM & " / 100, 0, - 1) AS 事務委託料"
        Else
            strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1) AS 販売保証料"
            strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1) AS 販売手数料"
            'strSQL +=  ", ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1) AS ＲＤ保証料"
            strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1) AS 事務委託料"
        End If
        strSQL += ", txt_data.entry_date"
        strSQL += " FROM txt_data LEFT OUTER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code LEFT OUTER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "dtl")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("dtl"), "", "ordr_no", DataViewRowState.CurrentRows)
        If DtView1.Count <> 0 Then
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If

            For i = 0 To DtView1.Count - 1
                j = i
                strDat2(j, 0) = DtView1(i)("ordr_date")
                strDat2(j, 1) = DtView1(i)("ordr_no")
                strDat2(j, 2) = DtView1(i)("model_name")
                strDat2(j, 3) = DtView1(i)("item_cat_code")
                strDat2(j, 4) = DtView1(i)("bend_code")
                strDat2(j, 5) = DtView1(i)("prch_price_tax")
                strDat2(j, 6) = DtView1(i)("wrn_fee")
                strDat2(j, 7) = DtView1(i)("wrn_prod")
                strDat2(j, 8) = DtView1(i)("cont_flg")
                strDat2(j, 9) = DtView1(i)("cust_name")
                strDat2(j, 10) = DtView1(i)("zip_code")
                strDat2(j, 11) = DtView1(i)("adrs")
                strDat2(j, 12) = DtView1(i)("tel1")
                strDat2(j, 13) = DtView1(i)("tel2")
                If IsDBNull(DtView1(i)("CLS_CODE_NAME")) Then DtView1(i)("CLS_CODE_NAME") = "その他"
                strDat2(j, 14) = DtView1(i)("CLS_CODE_NAME")
                ''If DtView1(i)("CLS_CODE_NAME") = "PC" _
                ''    And DtView1(i)("prch_price_tax") <= 110000 Then
                ''    If WK_comp = "4" Then
                ''        strDat2(j, 15) = L_gak_HO
                ''        strDat2(j, 16) = L_gak_TE
                ''        strDat2(j, 17) = L_gak_RD
                ''        strDat2(j, 18) = L_gak_JM
                ''    Else
                ''        strDat2(j, 15) = gak_HO
                ''        strDat2(j, 16) = gak_TE
                ''        strDat2(j, 17) = gak_RD
                ''        strDat2(j, 18) = gak_JM
                ''    End If
                'Else
                strDat2(j, 15) = DtView1(i)("販売保証料")
                    strDat2(j, 16) = DtView1(i)("販売手数料")
                    strDat2(j, 17) = DtView1(i)("販売保証料") - DtView1(i)("販売手数料")
                    strDat2(j, 18) = DtView1(i)("事務委託料")
                'End If
                strDat2(j, 19) = "=G" & j + 2 & "-P" & j + 2
            Next
            xlRange2.Value = strDat2            'セルへデータの入力
            MRComObject(xlRange2)               'xlRange の解放

        End If

        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        Select Case WK_comp
            Case Is = "2"
                SaveFileDialog1.FileName = "ECカレント総括表(" & Mid(WK_str2, 3, 4) & ")_ソニア確認用.xls"
            Case Is = "3"
                SaveFileDialog1.FileName = "特価COM総括表(" & Mid(WK_str2, 3, 4) & ")_ソニア確認用.xls"
            Case Is = "4"
                SaveFileDialog1.FileName = "Laox総括表(" & Mid(WK_str2, 3, 4) & ")_ソニア確認用.xls"
        End Select
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        'If Err.Number <> 0 Then
        '    MessageBox.Show(Err.Description)
        '    Err.Clear()
        'End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Sub eBEST_kakunin()
        On Error GoTo ErrorHandler
        Cursor = System.Windows.Forms.Cursors.WaitCursor
        '==================  起動時の処理  ===================  
        Dim xlApp As New Excel.Application
        Dim xlBooks As Excel.Workbooks = xlApp.Workbooks
        '既存のファイルを開く場合
        Dim xlFilePath As String = p_dir & "\eBest総括表.xls"
        Dim xlBook As Excel.Workbook = xlBooks.Open(xlFilePath)
        Dim xlSheets As Excel.Sheets = xlBook.Worksheets
        Dim xlSheet As Excel.Worksheet = xlSheets.Item(1)
        xlApp.Visible = False

        WK_DsList2.Clear()
        '最後のデータがDBに追加されたか
        strSQL = "SELECT add_date FROM txt_data WHERE (add_date IS NULL)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(WK_DsList2, "add_date")
        DB_CLOSE()
        If r = 0 Then
            strSQL = "SELECT MAX(add_date) AS max FROM txt_data"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            DB_OPEN()
            DaList1.Fill(WK_DsList2, "max")
            DB_CLOSE()
            DtView1 = New DataView(WK_DsList2.Tables("max"), "", "", DataViewRowState.CurrentRows)
            WK_str = DtView1(0)("max")
        Else
            WK_str = Nothing
        End If

        '*****************************
        '** 明細
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(2)  'Sheet2
        Dim xlRange2 As Excel.Range
        Dim strDat2(9999, 25) As Object
        xlRange2 = xlSheet.Range("A2:X10000")    'データの入力セル範囲
        Dim xlCells2 As Excel.Range
        Dim xlRange2_2 As Excel.Range

        '明細
        WK_DsList1.Clear()
        strSQL = "SELECT txt_data.ordr_date, txt_data.ordr_no, txt_data.model_name, txt_data.item_cat_code, txt_data.bend_code"
        strSQL += ", txt_data.prch_price, txt_data.prch_tax, txt_data.wrn_fee, txt_data.wrn_fee_tax, txt_data.wrn_prod"
        strSQL += ", txt_data.cont_flg, txt_data.cust_name, txt_data.zip_code, txt_data.adrs, txt_data.tel1"
        strSQL += ", txt_data.tel2, V_cls_002.CLS_CODE_NAME, txt_data.prch_price_tax"
        strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1) AS 販売保証料"
        strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1) AS 販売手数料"
        'strSQL +=  ", ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1) AS ＲＤ保証料"
        strSQL += ", ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1) AS 事務委託料"
        strSQL += ", txt_data.entry_date"
        strSQL += " FROM V_cls_002 RIGHT OUTER JOIN"
        strSQL += " V_cat_mtr ON V_cls_002.CLS_CODE = V_cat_mtr.cat_code2 RIGHT OUTER JOIN"
        strSQL += " txt_data ON V_cat_mtr.cat_code = txt_data.item_cat_code"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "dtl")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("dtl"), "", "ordr_no", DataViewRowState.CurrentRows)
        If DtView1.Count <> 0 Then
            If IsDate(DtView1(0)("entry_date")) Then
                WK_str2 = Format(DateAdd(DateInterval.Month, -1, CDate(DtView1(0)("entry_date"))), "yyyyMM")
            Else
                WK_str2 = Nothing
            End If

            For i = 0 To DtView1.Count - 1
                j = i
                strDat2(j, 0) = j + 1
                strDat2(j, 1) = Mid(DtView1(i)("ordr_date"), 1, 4) & "/" & Mid(DtView1(i)("ordr_date"), 5, 2) & "/" & Mid(DtView1(i)("ordr_date"), 7, 2)
                strDat2(j, 2) = DtView1(i)("ordr_no")
                strDat2(j, 3) = DtView1(i)("model_name")
                strDat2(j, 4) = DtView1(i)("item_cat_code")
                strDat2(j, 5) = DtView1(i)("bend_code")
                strDat2(j, 6) = DtView1(i)("prch_price")
                strDat2(j, 7) = DtView1(i)("prch_tax")
                strDat2(j, 8) = DtView1(i)("wrn_fee")
                strDat2(j, 9) = DtView1(i)("wrn_fee_tax")
                strDat2(j, 10) = DtView1(i)("wrn_prod")
                strDat2(j, 11) = DtView1(i)("cont_flg")
                strDat2(j, 12) = DtView1(i)("cust_name")
                strDat2(j, 13) = DtView1(i)("zip_code")
                strDat2(j, 14) = DtView1(i)("adrs")
                strDat2(j, 15) = DtView1(i)("tel1")
                strDat2(j, 16) = DtView1(i)("tel2")
                strDat2(j, 17) = DtView1(i)("CLS_CODE_NAME")
                If Not IsDBNull(DtView1(i)("CLS_CODE_NAME")) Then
                    Select Case DtView1(i)("CLS_CODE_NAME")
                        Case Is = "PC"
                            If DtView1(i)("prch_price_tax") <= 110000 Then
                                strDat2(j, 18) = "1"
                            Else
                                strDat2(j, 18) = "2"
                            End If
                        Case Is = "プリンタ"
                            strDat2(j, 18) = "3"
                        Case Else
                            strDat2(j, 18) = "4"
                    End Select
                Else
                    strDat2(j, 18) = "4"
                End If
                strDat2(j, 19) = DtView1(i)("prch_price_tax")

                If IsDBNull(DtView1(i)("CLS_CODE_NAME")) Then DtView1(i)("CLS_CODE_NAME") = "その他"
                'If DtView1(i)("CLS_CODE_NAME") = "PC" _
                '    And DtView1(i)("prch_price_tax") <= 110000 Then
                '    strDat2(j, 20) = gak_HO
                '    strDat2(j, 21) = gak_TE
                '    strDat2(j, 22) = gak_RD
                '    strDat2(j, 23) = gak_JM
                'Else
                strDat2(j, 20) = DtView1(i)("販売保証料")
                    strDat2(j, 21) = DtView1(i)("販売手数料")
                    strDat2(j, 22) = DtView1(i)("販売保証料") - DtView1(i)("販売手数料")
                    strDat2(j, 23) = DtView1(i)("事務委託料")
                ' End If
            Next
            '  xlRange2.Value = strDat2            'セルへデータの入力
            MRComObject(xlRange2)               'xlRange の解放

        End If

        '*****************************
        '** 総括表
        '*****************************
        '==================  データの入力処理  ==================  
        xlSheet = xlSheets.Item(1)  'Sheet1
        Dim xlRange As Excel.Range
        Dim strDat(8, 6) As Object     '2015/02/09 サイズ変更
        xlRange = xlSheet.Range("C1:H8")    'データの入力セル範囲  2015/02/09 範囲変更

        strDat(0, 4) = Mid(WK_str2, 1, 4) & "年" & CInt(Mid(WK_str2, 5, 2)) & "月受注分"

        WK_DsList1.Clear()
        'PC 10万超
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
        'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'PC')"
        strSQL += " AND (txt_data.prch_price_tax > '110000')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC1")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC1"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(3, 0) = DtView1(0)("商品価格")
            strDat(3, 1) = DtView1(0)("販売保証料")
            strDat(3, 2) = DtView1(0)("販売手数料")
            strDat(3, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(3, 4) = DtView1(0)("事務委託料")
            strDat(3, 5) = DtView1(0)("cnt")
        Else
            strDat(3, 0) = "0"
            strDat(3, 1) = "0"
            strDat(3, 2) = "0"
            strDat(3, 3) = "0"
            strDat(3, 4) = "0"
            strDat(3, 5) = "0"
        End If

        'PC 10万以下
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += ", SUM(txt_data.wrn_fee_wtax) AS 販売保証料"
        strSQL += ", SUM(txt_data.commission_fee_wtax) AS 販売手数料"
        strSQL += ", SUM(txt_data.wrn_fee_wtax - txt_data.commission_fee_wtax) AS RD保証料"
        strSQL += ", SUM(txt_data.admin_fee_wtax) AS 事務委託料"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'PC')"
        strSQL += " AND (txt_data.prch_price_tax <= '110000')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PC2")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PC2"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(4, 0) = DtView1(0)("商品価格")
            'strDat(4, 1) = DtView1(0)("cnt") * gak_HO
            'strDat(4, 2) = DtView1(0)("cnt") * gak_TE
            'strDat(4, 3) = DtView1(0)("cnt") * gak_RD
            'strDat(4, 4) = DtView1(0)("cnt") * gak_JM
            strDat(4, 1) = DtView1(0)("販売保証料")
            strDat(4, 2) = DtView1(0)("販売保証料")
            strDat(4, 3) = DtView1(0)("RD保証料")
            strDat(4, 4) = DtView1(0)("事務委託料")
            strDat(4, 5) = DtView1(0)("cnt")
        Else
            strDat(4, 0) = "0"
            strDat(4, 1) = "0"
            strDat(4, 2) = "0"
            strDat(4, 3) = "0"
            strDat(4, 4) = "0"
            strDat(4, 5) = "0"
        End If

        'プリンタ
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
        'strSQL +=  ", SUM(ROUND(txt_data.prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (V_cls_002.CLS_CODE_NAME = 'プリンタ')"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "PRT")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("PRT"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(5, 0) = DtView1(0)("商品価格")
            strDat(5, 1) = DtView1(0)("販売保証料")
            strDat(5, 2) = DtView1(0)("販売手数料")
            strDat(5, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(5, 4) = DtView1(0)("事務委託料")
            strDat(5, 5) = DtView1(0)("cnt")
        Else
            strDat(5, 0) = "0"
            strDat(5, 1) = "0"
            strDat(5, 2) = "0"
            strDat(5, 3) = "0"
            strDat(5, 4) = "0"
            strDat(5, 5) = "0"
        End If

        '2015/02/09 3年保証商品　20120620追加対応 ADD START 
        '3年保証その他
        strSQL = "SELECT SUM(txt_data.prch_price_tax) AS 商品価格"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
        strSQL += ", SUM(ROUND(txt_data.prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cat_mtr ON txt_data.item_cat_code = V_cat_mtr.cat_code INNER JOIN"
        strSQL += " V_cls_002 ON V_cat_mtr.cat_code2 = V_cls_002.CLS_CODE"
        strSQL += " WHERE (txt_data.comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (txt_data.wrn_prod = 3)"
        strSQL += " AND (RTRIM(V_cls_002.CLS_CODE) IN ('7068', '7515', '7518', '7545')) "
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "3oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("3oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(6, 0) = DtView1(0)("商品価格")
            strDat(6, 1) = DtView1(0)("販売保証料")
            strDat(6, 2) = DtView1(0)("販売手数料")
            strDat(6, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(6, 4) = DtView1(0)("事務委託料")
            strDat(6, 5) = DtView1(0)("cnt")
        Else
            strDat(6, 0) = "0"
            strDat(6, 1) = "0"
            strDat(6, 2) = "0"
            strDat(6, 3) = "0"
            strDat(6, 4) = "0"
            strDat(6, 5) = "0"
        End If
        '2015/02/09 3年保証商品　20120620追加対応 ADD END

        '5年全商品
        strSQL = "SELECT SUM(prch_price_tax) AS 商品価格"
        strSQL += ", SUM(ROUND(prch_price_tax * " & ret_HO & " / 100, 0, - 1)) AS 販売保証料"
        strSQL += ", SUM(ROUND(prch_price_tax * " & ret_TE & " / 100, 0, - 1)) AS 販売手数料"
        'strSQL +=  ", SUM(ROUND(prch_price_tax * " & ret_RD & " / 100, 0, - 1)) AS ＲＤ保証料"
        strSQL += ", SUM(ROUND(prch_price_tax * " & ret_JM & " / 100, 0, - 1)) AS 事務委託料"
        strSQL += ", COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " WHERE (comp = '" & WK_comp & "')"
        If WK_str = Nothing Then
            strSQL += " AND (txt_data.add_date IS NULL)"
        Else
            strSQL += " AND (txt_data.add_date = CONVERT(DATETIME, '" & WK_str & "', 102))"
        End If
        strSQL += " AND (wrn_prod = 5)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(WK_DsList1, "oth")
        DB_CLOSE()

        DtView1 = New DataView(WK_DsList1.Tables("oth"), "", "", DataViewRowState.CurrentRows)
        If DtView1(0)("cnt") <> 0 Then
            strDat(7, 0) = DtView1(0)("商品価格")
            strDat(7, 1) = DtView1(0)("販売保証料")
            strDat(7, 2) = DtView1(0)("販売手数料")
            strDat(7, 3) = DtView1(0)("販売保証料") - DtView1(0)("販売手数料")
            strDat(7, 4) = DtView1(0)("事務委託料")
            strDat(7, 5) = DtView1(0)("cnt")
        Else
            strDat(7, 0) = "0"
            strDat(7, 1) = "0"
            strDat(7, 2) = "0"
            strDat(7, 3) = "0"
            strDat(7, 4) = "0"
            strDat(7, 5) = "0"
        End If

        xlRange.Value = strDat          'セルへデータの入力
        MRComObject(xlRange)            'xlRange の解放

        '［名前を付けて保存］ダイアログボックスを表示
        'SaveFileDialog1.InitialDirectory = Application.StartupPath & "\.."
        SaveFileDialog1.FileName = "eBest総括表(" & Mid(WK_str2, 3, 4) & ")_ソニア確認用.xls"
        SaveFileDialog1.Filter = "Excelファイル|*.xls"
        SaveFileDialog1.OverwritePrompt = False
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            xlBook.SaveAs(SaveFileDialog1.FileName)
            CX_F = "0"
        Else
            CX_F = "1"
        End If

        '==================  終了処理  =====================  
        MRComObject(xlSheet)            'xlSheet の解放
        MRComObject(xlSheets)           'xlSheets の解放
        xlBook.Close(False)             'xlBook を閉じる
        MRComObject(xlBook)             'xlBook の解放
        MRComObject(xlBooks)            'xlBooks の解放
        xlApp.Quit()                    'Excelを閉じる 
        MRComObject(xlApp)              'xlApp を解放

        If CX_F = "0" Then
            MessageBox.Show(SaveFileDialog1.FileName & " に出力しました。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Cursor = System.Windows.Forms.Cursors.Default
        Exit Sub

ErrorHandler:
        CX_F = "1"
        'If Err.Number <> 0 Then
        '    MessageBox.Show(Err.Description)
        '    Err.Clear()
        'End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub MRComObject(ByVal objXl As Object)
        'Excel 終了処理時のプロシージャ
        Try
            '提供されたランタイム呼び出し可能ラッパーの参照カウントをデクリメントします
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objXl)
        Catch
        Finally
            objXl = Nothing
        End Try
    End Sub

    '********************************************************************
    '** 戻る
    '********************************************************************
    Private Sub Button99_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button99.Click
        DsList1.Clear()
        WK_DsList1.Clear()
        WK_DsList2.Clear()
        Me.Close()
    End Sub

    Private Sub Button99_Validating(sender As Object, e As CancelEventArgs) Handles Button99.Validating

    End Sub
End Class

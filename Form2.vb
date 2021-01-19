Public Class Form2
    Inherits System.Windows.Forms.Form

    Dim waitDlg As WaitDialog   '進行状況フォームクラス 

    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsList1, WK_DsList1 As New DataSet
    Dim DtView1, WK_DtView1 As DataView

    Dim strSQL, WK_str, ans, WK_skp As String
    Dim i, j, r, WK_line As Integer
    Dim WK_now As Date

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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
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
    Friend WithEvents DataGridTextBoxColumn22 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Date1 As GrapeCity.Win.Input.Interop.Date
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DataGridTextBoxColumn23 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn24 As System.Windows.Forms.DataGridTextBoxColumn
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
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
        Me.DataGridTextBoxColumn13 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn14 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn15 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn16 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn17 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn18 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn19 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn20 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn21 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn22 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Date1 = New GrapeCity.Win.Input.Interop.Date
        Me.Label2 = New System.Windows.Forms.Label
        Me.DataGridTextBoxColumn23 = New System.Windows.Forms.DataGridTextBoxColumn
        Me.DataGridTextBoxColumn24 = New System.Windows.Forms.DataGridTextBoxColumn
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Date1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(12, 28)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ReadOnly = True
        Me.DataGrid1.RowHeaderWidth = 10
        Me.DataGrid1.Size = New System.Drawing.Size(960, 508)
        Me.DataGrid1.TabIndex = 0
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2, Me.DataGridTextBoxColumn3, Me.DataGridTextBoxColumn4, Me.DataGridTextBoxColumn5, Me.DataGridTextBoxColumn6, Me.DataGridTextBoxColumn7, Me.DataGridTextBoxColumn8, Me.DataGridTextBoxColumn9, Me.DataGridTextBoxColumn10, Me.DataGridTextBoxColumn11, Me.DataGridTextBoxColumn12, Me.DataGridTextBoxColumn13, Me.DataGridTextBoxColumn23, Me.DataGridTextBoxColumn24, Me.DataGridTextBoxColumn14, Me.DataGridTextBoxColumn15, Me.DataGridTextBoxColumn16, Me.DataGridTextBoxColumn17, Me.DataGridTextBoxColumn18, Me.DataGridTextBoxColumn19, Me.DataGridTextBoxColumn20, Me.DataGridTextBoxColumn21, Me.DataGridTextBoxColumn22})
        Me.DataGridTableStyle1.HeaderFont = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "txt_data"
        Me.DataGridTableStyle1.ReadOnly = True
        Me.DataGridTableStyle1.RowHeaderWidth = 10
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "区分"
        Me.DataGridTextBoxColumn1.MappingName = "comp_name"
        Me.DataGridTextBoxColumn1.Width = 70
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn2.Format = ""
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "発注日"
        Me.DataGridTextBoxColumn2.MappingName = "ordr_date"
        Me.DataGridTextBoxColumn2.Width = 60
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "保証番号"
        Me.DataGridTextBoxColumn3.MappingName = "ordr_no"
        Me.DataGridTextBoxColumn3.Width = 75
        '
        'DataGridTextBoxColumn4
        '
        Me.DataGridTextBoxColumn4.Format = ""
        Me.DataGridTextBoxColumn4.FormatInfo = Nothing
        Me.DataGridTextBoxColumn4.HeaderText = "製品型式"
        Me.DataGridTextBoxColumn4.MappingName = "model_name"
        Me.DataGridTextBoxColumn4.Width = 90
        '
        'DataGridTextBoxColumn5
        '
        Me.DataGridTextBoxColumn5.Format = ""
        Me.DataGridTextBoxColumn5.FormatInfo = Nothing
        Me.DataGridTextBoxColumn5.HeaderText = "分類"
        Me.DataGridTextBoxColumn5.MappingName = "item_cat_code"
        Me.DataGridTextBoxColumn5.Width = 60
        '
        'DataGridTextBoxColumn6
        '
        Me.DataGridTextBoxColumn6.Format = ""
        Me.DataGridTextBoxColumn6.FormatInfo = Nothing
        Me.DataGridTextBoxColumn6.HeaderText = "メーカ"
        Me.DataGridTextBoxColumn6.MappingName = "bend_code"
        Me.DataGridTextBoxColumn6.Width = 50
        '
        'DataGridTextBoxColumn7
        '
        Me.DataGridTextBoxColumn7.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn7.Format = ""
        Me.DataGridTextBoxColumn7.FormatInfo = Nothing
        Me.DataGridTextBoxColumn7.HeaderText = "購入金額"
        Me.DataGridTextBoxColumn7.MappingName = "prch_price"
        Me.DataGridTextBoxColumn7.Width = 75
        '
        'DataGridTextBoxColumn8
        '
        Me.DataGridTextBoxColumn8.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn8.Format = ""
        Me.DataGridTextBoxColumn8.FormatInfo = Nothing
        Me.DataGridTextBoxColumn8.HeaderText = "消費税"
        Me.DataGridTextBoxColumn8.MappingName = "prch_tax"
        Me.DataGridTextBoxColumn8.Width = 60
        '
        'DataGridTextBoxColumn9
        '
        Me.DataGridTextBoxColumn9.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn9.Format = ""
        Me.DataGridTextBoxColumn9.FormatInfo = Nothing
        Me.DataGridTextBoxColumn9.HeaderText = "購入金額(税込)"
        Me.DataGridTextBoxColumn9.MappingName = "prch_price_tax"
        Me.DataGridTextBoxColumn9.Width = 75
        '
        'DataGridTextBoxColumn10
        '
        Me.DataGridTextBoxColumn10.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn10.Format = ""
        Me.DataGridTextBoxColumn10.FormatInfo = Nothing
        Me.DataGridTextBoxColumn10.HeaderText = "数量"
        Me.DataGridTextBoxColumn10.MappingName = "unit"
        Me.DataGridTextBoxColumn10.Width = 50
        '
        'DataGridTextBoxColumn11
        '
        Me.DataGridTextBoxColumn11.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn11.Format = ""
        Me.DataGridTextBoxColumn11.FormatInfo = Nothing
        Me.DataGridTextBoxColumn11.HeaderText = "保証料"
        Me.DataGridTextBoxColumn11.MappingName = "wrn_fee"
        Me.DataGridTextBoxColumn11.Width = 60
        '
        'DataGridTextBoxColumn12
        '
        Me.DataGridTextBoxColumn12.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn12.Format = ""
        Me.DataGridTextBoxColumn12.FormatInfo = Nothing
        Me.DataGridTextBoxColumn12.HeaderText = "消費税"
        Me.DataGridTextBoxColumn12.MappingName = "wrn_fee_tax"
        Me.DataGridTextBoxColumn12.Width = 60
        '
        'DataGridTextBoxColumn13
        '
        Me.DataGridTextBoxColumn13.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn13.Format = ""
        Me.DataGridTextBoxColumn13.FormatInfo = Nothing
        Me.DataGridTextBoxColumn13.HeaderText = "保証年"
        Me.DataGridTextBoxColumn13.MappingName = "wrn_prod"
        Me.DataGridTextBoxColumn13.Width = 60
        '
        'DataGridTextBoxColumn14
        '
        Me.DataGridTextBoxColumn14.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn14.Format = ""
        Me.DataGridTextBoxColumn14.FormatInfo = Nothing
        Me.DataGridTextBoxColumn14.HeaderText = "状況"
        Me.DataGridTextBoxColumn14.MappingName = "cont_flg"
        Me.DataGridTextBoxColumn14.Width = 40
        '
        'DataGridTextBoxColumn15
        '
        Me.DataGridTextBoxColumn15.Format = ""
        Me.DataGridTextBoxColumn15.FormatInfo = Nothing
        Me.DataGridTextBoxColumn15.HeaderText = "顧客名"
        Me.DataGridTextBoxColumn15.MappingName = "cust_name"
        Me.DataGridTextBoxColumn15.Width = 90
        '
        'DataGridTextBoxColumn16
        '
        Me.DataGridTextBoxColumn16.Format = ""
        Me.DataGridTextBoxColumn16.FormatInfo = Nothing
        Me.DataGridTextBoxColumn16.HeaderText = "郵便番号"
        Me.DataGridTextBoxColumn16.MappingName = "zip_code"
        Me.DataGridTextBoxColumn16.Width = 75
        '
        'DataGridTextBoxColumn17
        '
        Me.DataGridTextBoxColumn17.Format = ""
        Me.DataGridTextBoxColumn17.FormatInfo = Nothing
        Me.DataGridTextBoxColumn17.HeaderText = "住所"
        Me.DataGridTextBoxColumn17.MappingName = "adrs"
        Me.DataGridTextBoxColumn17.Width = 90
        '
        'DataGridTextBoxColumn18
        '
        Me.DataGridTextBoxColumn18.Format = ""
        Me.DataGridTextBoxColumn18.FormatInfo = Nothing
        Me.DataGridTextBoxColumn18.HeaderText = "固定電話"
        Me.DataGridTextBoxColumn18.MappingName = "tel1"
        Me.DataGridTextBoxColumn18.Width = 75
        '
        'DataGridTextBoxColumn19
        '
        Me.DataGridTextBoxColumn19.Format = ""
        Me.DataGridTextBoxColumn19.FormatInfo = Nothing
        Me.DataGridTextBoxColumn19.HeaderText = "携帯電話"
        Me.DataGridTextBoxColumn19.MappingName = "tel2"
        Me.DataGridTextBoxColumn19.Width = 75
        '
        'DataGridTextBoxColumn20
        '
        Me.DataGridTextBoxColumn20.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn20.Format = ""
        Me.DataGridTextBoxColumn20.FormatInfo = Nothing
        Me.DataGridTextBoxColumn20.HeaderText = "報告日"
        Me.DataGridTextBoxColumn20.MappingName = "entry_date"
        Me.DataGridTextBoxColumn20.Width = 75
        '
        'DataGridTextBoxColumn21
        '
        Me.DataGridTextBoxColumn21.Format = ""
        Me.DataGridTextBoxColumn21.FormatInfo = Nothing
        Me.DataGridTextBoxColumn21.HeaderText = "JAN"
        Me.DataGridTextBoxColumn21.MappingName = "jan"
        Me.DataGridTextBoxColumn21.Width = 75
        '
        'DataGridTextBoxColumn22
        '
        Me.DataGridTextBoxColumn22.Format = ""
        Me.DataGridTextBoxColumn22.FormatInfo = Nothing
        Me.DataGridTextBoxColumn22.HeaderText = "元JAN"
        Me.DataGridTextBoxColumn22.MappingName = "moto_jan"
        Me.DataGridTextBoxColumn22.Width = 75
        '
        'Button4
        '
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.Location = New System.Drawing.Point(864, 552)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 28)
        Me.Button4.TabIndex = 2
        Me.Button4.Text = "戻　る"
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(8, 548)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "反　映"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(784, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(188, 20)
        Me.Label1.TabIndex = 9
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Date1
        '
        Me.Date1.DisplayFormat = New GrapeCity.Win.Input.Interop.DateDisplayFormat("yyyy/MM", "", "")
        Me.Date1.DropDown = New GrapeCity.Win.Input.Interop.DropDown(GrapeCity.Win.Input.Interop.ButtonPosition.Inside, True, GrapeCity.Win.Input.Interop.Visibility.NotShown, System.Windows.Forms.FlatStyle.System)
        Me.Date1.DropDownCalendar.Size = New System.Drawing.Size(158, 165)
        Me.Date1.Format = New GrapeCity.Win.Input.Interop.DateFormat("yyyy/MM", "", "")
        Me.Date1.Location = New System.Drawing.Point(88, 4)
        Me.Date1.MaxDate = New GrapeCity.Win.Input.Interop.DateTimeEx(New Date(2020, 12, 31, 23, 59, 59, 0))
        Me.Date1.MinDate = New GrapeCity.Win.Input.Interop.DateTimeEx(New Date(2000, 1, 1, 0, 0, 0, 0))
        Me.Date1.Name = "Date1"
        Me.Date1.Shortcuts = New GrapeCity.Win.Input.Interop.ShortcutCollection(New String() {"F2", "F5"}, New GrapeCity.Win.Input.Interop.KeyActions() {GrapeCity.Win.Input.Interop.KeyActions.Clear, GrapeCity.Win.Input.Interop.KeyActions.Now})
        Me.Date1.Size = New System.Drawing.Size(88, 24)
        Me.Date1.TabIndex = 0
        Me.Date1.TextHAlign = GrapeCity.Win.Input.Interop.AlignHorizontal.Center
        Me.Date1.TextVAlign = GrapeCity.Win.Input.Interop.AlignVertical.Middle
        Me.Date1.Value = New GrapeCity.Win.Input.Interop.DateTimeEx(New Date(2012, 10, 15, 11, 8, 44, 0))
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Navy
        Me.Label2.ForeColor = System.Drawing.SystemColors.Window
        Me.Label2.Location = New System.Drawing.Point(12, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 24)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "処理日"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DataGridTextBoxColumn23
        '
        Me.DataGridTextBoxColumn23.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridTextBoxColumn23.Format = ""
        Me.DataGridTextBoxColumn23.FormatInfo = Nothing
        Me.DataGridTextBoxColumn23.HeaderText = "セット"
        Me.DataGridTextBoxColumn23.MappingName = "set_flg"
        Me.DataGridTextBoxColumn23.Width = 50
        '
        'DataGridTextBoxColumn24
        '
        Me.DataGridTextBoxColumn24.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn24.Format = ""
        Me.DataGridTextBoxColumn24.FormatInfo = Nothing
        Me.DataGridTextBoxColumn24.HeaderText = "合計"
        Me.DataGridTextBoxColumn24.MappingName = "ttl"
        Me.DataGridTextBoxColumn24.Width = 60
        '
        'Form2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.ClientSize = New System.Drawing.Size(982, 583)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Date1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DataGrid1)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "DB反映"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Date1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '******************************************************************
    '** 起動時
    '******************************************************************
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        DsList1.Clear()
        strSQL = "SELECT MAX(proc_date) AS max_date FROM Wrn_mtr"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        DaList1.Fill(DsList1, "max_date")
        DB_CLOSE()
        DtView1 = New DataView(DsList1.Tables("max_date"), "", "", DataViewRowState.CurrentRows)
        If Not IsDBNull(DtView1(0)("max_date")) Then
            Date1.Text = Format(DateAdd(DateInterval.Month, 1, CDate(Trim(DtView1(0)("max_date")) & "/01")), "yyyy/MM")
        Else
            Date1.Text = Format(DateAdd(DateInterval.Month, -1, Now.Date), "yyyy/MM")
        End If

        strSQL = "SELECT txt_data.*, V_cls_001.CLS_CODE_NAME AS comp_name"
        strSQL += " FROM txt_data INNER JOIN"
        strSQL += " V_cls_001 ON txt_data.comp = V_cls_001.CLS_CODE"
        strSQL += " WHERE (txt_data.add_F = 0)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(DsList1, "txt_data")
        DB_CLOSE()

        Dim tbl1 As New DataTable
        tbl1 = DsList1.Tables("txt_data")
        DataGrid1.DataSource = tbl1

        Label1.Text = Format(r, "##,##0") & "件"

        If r = 0 Then
            Button2.Enabled = False
            MsgBox("対象データなし")
        Else
            Button2.Enabled = True
        End If

    End Sub

    '******************************************************************
    '** DB反映
    '******************************************************************
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ans = MessageBox.Show("本番データへ登録します。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)
        If ans = "2" Then Exit Sub 'いいえ

        Cursor = System.Windows.Forms.Cursors.WaitCursor
        WK_now = Now

        ' 進行状況ダイアログの初期化処理
        waitDlg = New WaitDialog        ' 進行状況ダイアログ
        waitDlg.Owner = Me              ' ダイアログのオーナーを設定する
        waitDlg.MainMsg = "ﾃﾞｰﾀ抽出中"  ' 処理の概要を表示
        waitDlg.ProgressMax = 0         ' 全体の処理件数を設定
        waitDlg.ProgressMin = 0         ' 処理件数の最小値を設定（0件から開始）
        waitDlg.ProgressStep = 1        ' 何件ごとにメータを進めるかを設定
        waitDlg.ProgressValue = 0       ' 最初の件数を設定
        Me.Enabled = False              ' オーナーのフォームを無効にする
        waitDlg.Show()                  ' 進行状況ダイアログを表示する

        DtView1 = New DataView(DsList1.Tables("txt_data"), "", "", DataViewRowState.CurrentRows)

        waitDlg.MainMsg = "ﾃﾞｰﾀ登録中"      ' 進行状況ダイアログのメーターを設定
        waitDlg.ProgressMax = DtView1.Count ' 全体の処理件数を設定
        waitDlg.ProgressValue = 0           ' 最初の件数を設定
        Application.DoEvents()              ' メッセージ処理を促して表示を更新する

        DB_OPEN()
        For i = 0 To DtView1.Count - 1

            waitDlg.ProgressMsg = Fix((i + 1) * 100 / DtView1.Count) & "%　（" & Format(i + 1, "##,##0") & " / " & Format(DtView1.Count, "##,##0") & " 件）"
            waitDlg.Text = "実行中・・・" & Fix((i + 1) * 100 / DtView1.Count) & "%"
            Application.DoEvents()  ' メッセージ処理を促して表示を更新する
            waitDlg.PerformStep()   ' 処理カウントを1ステップ進める

            'Wrn_mtr
            WK_DsList1.Clear()
            strSQL = "SELECT ordr_no FROM Wrn_mtr WHERE (ordr_no = '" & DtView1(i)("ordr_no") & "')"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            DaList1.SelectCommand = SqlCmd1
            r = DaList1.Fill(WK_DsList1, "Wrn_mtr")
            If r = 0 Then
                strSQL = "INSERT INTO Wrn_mtr"
                strSQL += " (ordr_no, cust_name, cust_name_srch, adrs, zip_code, tel1, tel2, entry_date, kbn, proc_date)"
                strSQL += " VALUES ('" & MidB(DtView1(i)("ordr_no"), 1, 14) & "'"
                strSQL += ", '" & MidB(DtView1(i)("cust_name"), 1, 30) & "'"
                WK_str = DtView1(i)("cust_name")
                WK_str = WK_str.Replace(" ", "")
                WK_str = WK_str.Replace("　", "")
                strSQL += ", '" & MidB(WK_str, 1, 30) & "'"
                strSQL += ", '" & MidB(DtView1(i)("adrs"), 1, 100) & "'"
                WK_str = DtView1(i)("zip_code")
                WK_str = WK_str.Replace("-", "")
                strSQL += ", '" & MidB(WK_str, 1, 7) & "'"
                strSQL += ", '" & MidB(DtView1(i)("tel1"), 1, 15) & "'"
                strSQL += ", '" & MidB(DtView1(i)("tel2"), 1, 15) & "'"
                If IsDate(DtView1(i)("entry_date")) Then
                    strSQL += ", CONVERT(DATETIME, '" & DtView1(i)("entry_date") & "', 102)"
                Else
                    strSQL += ", NULL"
                End If
                strSQL += ", " & DtView1(i)("comp") & ""
                strSQL += ", N'" & Date1.Text & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.ExecuteNonQuery()
                'WK_skp = "0"
            Else
                'WK_str = "重複 : " & DtView1(i)("ordr_no") & vbLf
                'WK_str = WK_str & vbLf & "氏  名 ： " & DtView1(i)("cust_name")
                'WK_str = WK_str & vbLf & "型  式 ： " & DtView1(i)("model_name")
                'WK_str = WK_str & vbLf & "購入額 ： " & DtView1(i)("prch_price_tax")
                'WK_str = WK_str & vbLf & "購入日 ： " & DtView1(i)("ordr_date")

                'WK_str = WK_str & vbLf & vbLf & "読み飛ばします。"
                'MsgBox(WK_str, MsgBoxStyle.OKOnly, "Error")
                'WK_skp = "1"
            End If

            'Wrn_sub
            If DtView1(i)("cont_flg") = "1" Then    '加入

                strSQL = "SELECT MAX(line_no) AS max_line FROM Wrn_sub WHERE (ordr_no = '" & DtView1(i)("ordr_no") & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                DaList1.SelectCommand = SqlCmd1
                r = DaList1.Fill(WK_DsList1, "Wrn_sub")
                If r = 0 Then
                    WK_line = 1
                Else
                    WK_DtView1 = New DataView(WK_DsList1.Tables("Wrn_sub"), "", "", DataViewRowState.CurrentRows)
                    If Not IsDBNull(WK_DtView1(0)("max_line")) Then
                        WK_line = WK_DtView1(0)("max_line") + 1
                    Else
                        WK_line = 1
                    End If
                End If

                For j = 1 To DtView1(i)("unit") '数量
                    strSQL = "INSERT INTO Wrn_sub"
                    strSQL += " (ordr_no, line_no, seq, prch_price, prch_tax, prch_date, item_cat_code, wrn_prod, set_flg, ttl"
                    strSQL += ", wrn_plan"  '2015/08/13 電動工具保証追加対応
                    strSQL += ", cont_flg, bend_code, model_name, wrn_fee, wrn_fee_tax, op_date, jan, moto_jan)"
                    strSQL += " VALUES ('" & MidB(DtView1(i)("ordr_no"), 1, 14) & "'"
                    strSQL += ", " & WK_line
                    strSQL += ", " & j
                    strSQL += ", " & DtView1(i)("prch_price") & ""
                    strSQL += ", " & DtView1(i)("prch_tax") & ""
                    strSQL += ", CONVERT(DATETIME, '" & Mid(DtView1(i)("ordr_date"), 1, 4) & "/" & Mid(DtView1(i)("ordr_date"), 5, 2) & "/" & Mid(DtView1(i)("ordr_date"), 7, 2) & "', 102)"
                    '2015/08/13 電動工具保証追加対応
                    strSQL += ", '" & DtView1(i)("item_cat_code") & "'"
                    'Select Case MidB(DtView1(i)("item_cat_code"), 2, 1)
                    '    Case Is = "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                    'strSQL += ", '" & MidB(DtView1(i)("item_cat_code"), 1, 6) & "'"
                    '    Case Else
                    'strSQL += ", '" & MidB(DtView1(i)("item_cat_code"), 3, 6) & "'"
                    'End Select
                    strSQL += ", " & DtView1(i)("wrn_prod") & ""
                    strSQL += ", '" & DtView1(i)("set_flg") & "'"
                    strSQL += ", " & DtView1(i)("ttl") & ""
                    strSQL += ", " & DtView1(i)("wrn_plan") & ""   '2015/08/13 電動工具保証追加対応
                    strSQL += ", 'A'"
                    If DtView1(i)("bend_code") = Nothing Then
                        strSQL += ", 0"
                    Else
                        strSQL += ", " & DtView1(i)("bend_code") & ""
                    End If
                    strSQL += ", '" & MidB(DtView1(i)("model_name"), 1, 50) & "'"
                    strSQL += ", " & DtView1(i)("wrn_fee") & ""
                    strSQL += ", " & DtView1(i)("wrn_fee_tax") & ""
                    strSQL += ", CONVERT(DATETIME, '" & WK_now & "', 102)"
                    strSQL += ", '" & MidB(DtView1(i)("jan"), 1, 13) & "'"
                    strSQL += ", '" & MidB(DtView1(i)("moto_jan"), 1, 13) & "')"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    SqlCmd1.ExecuteNonQuery()
                Next

            Else                                    '取消

                strSQL = "SELECT line_no, seq FROM Wrn_sub"
                strSQL += " WHERE (ordr_no = '" & MidB(DtView1(i)("ordr_no"), 1, 14) & "')"
                strSQL += " AND (cont_flg = 'A')"
                strSQL += " AND (model_name = '" & MidB(DtView1(i)("model_name"), 1, 50) & "')"
                strSQL += " AND (prch_price = " & DtView1(i)("prch_price") & ")"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                DaList1.SelectCommand = SqlCmd1
                r = DaList1.Fill(WK_DsList1, "Wrn_sub")
                If r <> 0 Then
                    WK_DtView1 = New DataView(WK_DsList1.Tables("Wrn_sub"), "", "", DataViewRowState.CurrentRows)

                    strSQL = "UPDATE Wrn_sub"
                    strSQL += " SET cont_flg = 'C'"
                    strSQL += ", cxl_date = CONVERT(DATETIME, '" & DtView1(i)("entry_date") & "', 102)"
                    strSQL += " WHERE (ordr_no = '" & MidB(DtView1(i)("ordr_no"), 1, 14) & "')"
                    strSQL += " AND (line_no = " & WK_DtView1(0)("line_no") & ")"
                    strSQL += " AND (seq = " & WK_DtView1(0)("seq") & ")"
                    SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                    SqlCmd1.ExecuteNonQuery()
                End If
            End If

            strSQL = "UPDATE txt_data"
            strSQL += " SET add_F = 1"
            strSQL += ", add_date = CONVERT(DATETIME, '" & WK_now & "', 102)"
            strSQL += " WHERE (id = " & DtView1(i)("id") & ")"
            SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
            SqlCmd1.ExecuteNonQuery()

        Next
        DB_CLOSE()
        MsgBox("反映しました")
        DsList1.Clear()
        Button2.Enabled = False

        waitDlg.Close()         '進行状況ダイアログを閉じる
        Me.Enabled = True       'オーナーのフォームを無効にする

        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** 戻る
    '********************************************************************
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        DsList1.Clear()
        WK_DsList1.Clear()
        Me.Close()
    End Sub
End Class

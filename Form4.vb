Public Class Form4
    Inherits System.Windows.Forms.Form

    Dim waitDlg As WaitDialog   '進行状況フォームクラス 

    Dim SqlCmd1 As SqlClient.SqlCommand
    Dim DaList1 = New SqlClient.SqlDataAdapter
    Dim DsList1 As New DataSet
    Dim DtView1 As DataView

    Dim strSQL As String
    Dim i, j, r As Integer
    Friend WithEvents Button1 As Button

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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridBoolColumn1 As System.Windows.Forms.DataGridBoolColumn
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridTextBoxColumn2 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridBoolColumn1 = New System.Windows.Forms.DataGridBoolColumn()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridTextBoxColumn2 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.CaptionVisible = False
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(12, 28)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.RowHeaderWidth = 10
        Me.DataGrid1.Size = New System.Drawing.Size(492, 220)
        Me.DataGrid1.TabIndex = 0
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridBoolColumn1, Me.DataGridTextBoxColumn1, Me.DataGridTextBoxColumn2})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "txt_data"
        Me.DataGridTableStyle1.RowHeaderWidth = 10
        '
        'DataGridBoolColumn1
        '
        Me.DataGridBoolColumn1.Alignment = System.Windows.Forms.HorizontalAlignment.Center
        Me.DataGridBoolColumn1.HeaderText = "取消"
        Me.DataGridBoolColumn1.MappingName = "add_F"
        Me.DataGridBoolColumn1.NullValue = "False"
        Me.DataGridBoolColumn1.Width = 60
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "ファイル名"
        Me.DataGridTextBoxColumn1.MappingName = "file_name"
        Me.DataGridTextBoxColumn1.ReadOnly = True
        Me.DataGridTextBoxColumn1.Width = 300
        '
        'DataGridTextBoxColumn2
        '
        Me.DataGridTextBoxColumn2.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.DataGridTextBoxColumn2.Format = "##,##0"
        Me.DataGridTextBoxColumn2.FormatInfo = Nothing
        Me.DataGridTextBoxColumn2.HeaderText = "件数"
        Me.DataGridTextBoxColumn2.MappingName = "cnt"
        Me.DataGridTextBoxColumn2.ReadOnly = True
        Me.DataGridTextBoxColumn2.Width = 90
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(12, 256)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "削　除"
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(288, 4)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(24, 16)
        Me.CheckBox2.TabIndex = 1250
        Me.CheckBox2.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(316, 4)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(24, 16)
        Me.CheckBox1.TabIndex = 1249
        Me.CheckBox1.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(400, 257)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 28)
        Me.Button1.TabIndex = 1251
        Me.Button1.Text = "戻　る"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form4
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.ClientSize = New System.Drawing.Size(518, 291)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DataGrid1)
        Me.Font = New System.Drawing.Font("MS UI Gothic", 12.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form4"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "取込みの取消"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    '******************************************************************
    '** 起動時
    '******************************************************************
    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dsp_set() '** 画面セット
        If r = 0 Then
            MsgBox("対象データなし")
        End If
    End Sub

    '******************************************************************
    '** 画面セット
    '******************************************************************
    Sub dsp_set()
        DsList1.Clear()
        strSQL = "SELECT add_F, file_name, COUNT(*) AS cnt"
        strSQL += " FROM txt_data"
        strSQL += " GROUP BY file_name, add_F"
        strSQL += " HAVING (add_F = 0)"
        SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
        DaList1.SelectCommand = SqlCmd1
        DB_OPEN()
        r = DaList1.Fill(DsList1, "txt_data")
        DB_CLOSE()

        Dim tbl1 As New DataTable
        tbl1 = DsList1.Tables("txt_data")
        DataGrid1.DataSource = tbl1

        If r = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If

        '新しい行の追加を禁止する
        Dim cm As CurrencyManager
        cm = CType(Me.BindingContext(DataGrid1.DataSource), CurrencyManager)
        Dim dv As DataView = CType(cm.List, DataView)
        dv.AllowNew = False
    End Sub
    '********************************************************************
    '**  データグリッド色
    '********************************************************************
    'Private Sub DataGrid1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles DataGrid1.Paint
    '    If DataGrid1.CurrentRowIndex >= 0 Then
    '        DataGrid1.Select(DataGrid1.CurrentRowIndex)
    '    End If
    'End Sub

    '********************************************************************
    '**  データグリッド　クリック
    '********************************************************************
    Private Sub DataGrid1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles DataGrid1.MouseUp
        Dim hitinfo As DataGrid.HitTestInfo
        Try
            hitinfo = (CType(sender, DataGrid).HitTest(e.X, e.Y))

            If hitinfo.Column = 0 Then  'ﾁｪｯｸﾎﾞｯｸｽ ｸﾘｯｸ
                If DataGrid1(DataGrid1.CurrentRowIndex, 0) = "False" Then
                    DataGrid1(DataGrid1.CurrentRowIndex, 0) = CheckBox1.Checked
                Else
                    DataGrid1(DataGrid1.CurrentRowIndex, 0) = CheckBox2.Checked
                End If
            End If
        Catch ex As System.Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DsList1.Clear()
        Me.Close()
    End Sub

    '********************************************************************
    '** 削除
    '********************************************************************
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Cursor = System.Windows.Forms.Cursors.WaitCursor

        DtView1 = New DataView(DsList1.Tables("txt_data"), "add_F = 1", "", DataViewRowState.CurrentRows)
        If DtView1.Count > 0 Then
            DB_OPEN()
            For i = 0 To DtView1.Count - 1
                strSQL = "DELETE FROM txt_data"
                strSQL += " WHERE (file_name = '" & DtView1(i)("file_name") & "')"
                SqlCmd1 = New SqlClient.SqlCommand(strSQL, cnsqlclient)
                SqlCmd1.ExecuteNonQuery()
            Next
            DB_CLOSE()
            dsp_set() '** 画面セット
            MsgBox("削除しました。", MsgBoxStyle.OKOnly, "info")
        Else
            MsgBox("選択がありません。", MsgBoxStyle.OKOnly, "Error")
        End If
        Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    '********************************************************************
    '** 戻る
    '********************************************************************
    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    '    DsList1.Clear()
    '    Me.Close()
    'End Sub
End Class

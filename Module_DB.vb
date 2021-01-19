Module Module_DB
    Public cnsqlclient As New System.Data.SqlClient.SqlConnection
    Public source As String
    Public catalog As String
    Public uid As String
    Public pwd As String

    Public Function DB_INIT()
        Call DB_INIT_sub()
    End Function

    Sub DB_INIT_sub()
        Dim sr As System.io.StreamReader
        sr = New System.IO.StreamReader("STRM.ini")

        Dim line As String
        Dim line_len As Integer
        Dim eq_pos As Integer
        Dim line_key As String
        Do
            line = sr.ReadLine()
            line_len = Len(line)
            If line_len = 0 Then
            Else
                eq_pos = InStr(1, line, "=", 1)
                line_key = Mid(line, 1, eq_pos - 1)
                If line_key = "source" Then
                    source = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
                If line_key = "catalog" Then
                    catalog = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
                If line_key = "uid" Then
                    uid = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
                If line_key = "pwd" Then
                    pwd = Mid(line, eq_pos + 1, line_len - eq_pos)
                End If
            End If
        Loop Until line Is Nothing
        sr.Close()
    End Sub

    Public Function DB_OPEN() As Boolean
        DB_OPEN = False

        '*****  接続文字列を作成して接続を開始する  *****
        'cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source & ";" & _
        '                               "persist security info=False;initial catalog=" & catalog
        If Trim(uid) <> Nothing Then
            cnsqlclient.ConnectionString = "integrated security=false; uid=" & uid & "; pwd=" & pwd & "; data source=" & source & ";persist security info=False; initial catalog=" & catalog
        Else
            cnsqlclient.ConnectionString = "integrated security=SSPI;data source=" & source & ";persist security info=False;initial catalog=" & catalog
        End If

        Try
            '*****  Connectionが接続されているかチェック  *****
            If cnsqlclient.State = ConnectionState.Closed Then
                cnsqlclient.Open()
            End If
        Catch
            MsgBox(Err.Description, 16, "接続エラー")
            DB_OPEN = False
            Exit Function
        End Try

        DB_OPEN = True
    End Function

    Public Sub DB_CLOSE()
        '接続を終了する
        cnsqlclient.Close()
    End Sub
End Module

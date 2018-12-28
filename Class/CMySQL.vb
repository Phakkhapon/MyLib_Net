
Option Explicit On
Imports MySql.Data.MySqlClient

Public Class CMySQL

    Public Sub New()

    End Sub

    Public Function CommandMySqlDataTable(ByVal strSQL As String, ByVal MysqlConn As MySqlConnection, Optional ByVal nTimeout As Integer = 300) As DataTable

        If MysqlConn.State = ConnectionState.Closed Then MysqlConn.Open()
        'Return CommandOleDB(strSQL)
        Dim myData As New DataTable
        Dim cmd As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        'Try
        cmd.CommandTimeout = nTimeout
        cmd.CommandText = strSQL
            cmd.Connection = MysqlConn
            myAdapter.SelectCommand = cmd
            myAdapter.Fill(myData)
            'myData = Nothing
            cmd = Nothing
            myAdapter = Nothing
            ' End Try
            If MysqlConn.State = ConnectionState.Open Then MysqlConn.Close()
        MySqlConnection.ClearPool(MysqlConn)
        Return myData
    End Function

    Public Function CommandMySqlDataset(ByVal strSQL As String, ByVal MysqlConn As MySqlConnection, Optional ByVal nTimeOut As Integer = 300) As DataSet
        If MysqlConn.State = ConnectionState.Closed Then MysqlConn.Open()
        Dim myData As New DataSet
        Dim cmd As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        'Try

        cmd.CommandTimeout = nTimeOut
        cmd.CommandText = strSQL
        cmd.Connection = MysqlConn
        myAdapter.SelectCommand = cmd
        myAdapter.Fill(myData)
        cmd = Nothing
        myAdapter = Nothing
        If MysqlConn.State = ConnectionState.Open Then MysqlConn.Close()
        MySqlConnection.ClearPool(MysqlConn)
        Return myData
    End Function

    Public Sub CommandNonQuery(ByVal strSQL As String, ByVal MysqlConn As MySqlConnection, ByVal strProduct As String, ByVal strClassName As String, Optional ByVal nTimeOut As Integer = 300)
        If MysqlConn.State = ConnectionState.Closed Then MysqlConn.Open()
        Dim cmd As New MySqlCommand

        'Try
        cmd.CommandTimeout = nTimeOut
        strSQL = Replace(strSQL, """", "")
        cmd.CommandText = strSQL
        cmd.Connection = MysqlConn

        Dim nRowAffect As Integer = cmd.ExecuteNonQuery()
        'myData = Nothing
        ' myData.Clear()
        'Catch ex As Exception
        'myData = Nothing
        'myAdapter = Nothing
        ' End Try
        If nRowAffect > 0 Then
            Dim strActionTime As String = Format(Now, "yyyy-MM-dd HH:mm:ss")
            Dim strLogSQL As String
            strLogSQL = "INSERT INTO ctr_user_rttc.tabuseraction("
            strLogSQL = strLogSQL & "ActionTime,"
            strLogSQL = strLogSQL & "ActionUser,"
            strLogSQL = strLogSQL & "PageClass,"
            strLogSQL = strLogSQL & "ProductName,"
            strLogSQL = strLogSQL & "SqlScript) VALUES("
            strLogSQL = strLogSQL & "'" & strActionTime & "',"
            strLogSQL = strLogSQL & "'" & g_sCurrentUserDetail.strUserName & "',"
            strLogSQL = strLogSQL & "'" & strClassName & "',"
            strLogSQL = strLogSQL & "'" & strProduct & "',"
            strLogSQL = strLogSQL & """" & strSQL & """);"
            cmd.CommandText = strLogSQL
            cmd.ExecuteNonQuery()
        End If
        If MysqlConn.State = ConnectionState.Open Then MysqlConn.Close()
        MySqlConnection.ClearPool(MysqlConn)
    End Sub

    Public Function CommandNoQuery(ByVal strSQL As String, ByVal MysqlConn As MySqlConnection, Optional ByVal nTimeOut As Integer = 600) As Integer
        If MysqlConn.State = ConnectionState.Closed Then MysqlConn.Open()
        Dim cmd As New MySqlCommand

        'Try
        cmd.CommandTimeout = nTimeOut
        strSQL = Replace(strSQL, """", "")
        cmd.CommandText = strSQL
        cmd.Connection = MysqlConn
        CommandNoQuery = cmd.ExecuteNonQuery
        'myData = Nothing
        ' myData.Clear()
        'Catch ex As Exception
        'myData = Nothing
        'myAdapter = Nothing
        ' End Try
        If MysqlConn.State = ConnectionState.Open Then MysqlConn.Close()
        MySqlConnection.ClearPool(MysqlConn)
    End Function
End Class

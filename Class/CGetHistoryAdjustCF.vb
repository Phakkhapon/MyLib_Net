
Imports MySql.Data.MySqlClient

Public Class CGetHistoryAdjustCF
    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

    Public Function GetHistoryAdjCF(ByVal strProduct As String, ByVal dtbParam As DataTable, ByVal dtbSearch As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "ID,"
        strSQL = strSQL & "'" & strProduct & "' Product,"
        strSQL = strSQL & "Test_Time,"
        strSQL = strSQL & "Adjust_time,"
        strSQL = strSQL & "Tester,"
        strSQL = strSQL & "Spec,"
        strSQL = strSQL & "Lot,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "Send_By,"
        strSQL = strSQL & "Criteria,"
        strSQL = strSQL & "param_rttc,"
        strSQL = strSQL & "ParameterType,"
        strSQL = strSQL & "Old_Factor,"
        strSQL = strSQL & "New_Factor,"
        strSQL = strSQL & "Description "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabhistory_adjcf A "
        strSQL = strSQL & " WHERE (A.Test_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & "AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester

        strSQL = strSQL & "AND ("

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            If nParam <> dtbParam.Rows.Count - 1 Then
                strSQL = strSQL & "A.param_rttc='" & dtbParam.Rows(nParam).Item("param_rttc") & "' OR "
            Else
                strSQL = strSQL & "A.param_rttc='" & dtbParam.Rows(nParam).Item("param_rttc") & "') "
            End If
        Next nParam
        strSQL = strSQL & "ORDER BY A." & strSearchBy & ",Test_time;"
        Dim clsSql As New CMySQL
        GetHistoryAdjCF = clsSql.CommandMySqlDataTable(strSQL, m_mySqlConn)

    End Function

    Public Function GetHistoryNoAdjustX2Lot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String
        Dim clsMySql As New CMySQL

        strSQL = "SELECT * FROM db_" & strProduct & ".tabhistory_datanoadj A "
        strSQL = strSQL & "WHERE ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        GetHistoryNoAdjustX2Lot = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)

    End Function

End Class

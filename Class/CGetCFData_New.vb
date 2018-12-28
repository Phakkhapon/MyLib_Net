
Imports MySql.Data.MySqlClient

Public Class CGetCFData_New
    Private m_MySqlConn As MySqlConnection

    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetCFNow(ByVal strProduct As String, ByVal dtbTester As DataTable, ByVal dtbParam As DataTable) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT Test_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Shoe,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            Dim strDisplay As String = drParam.Item("param_display")
            If drParam.Item("param_add") Then strSQL = strSQL & "(SELECT para" & strParamID & " FROM db_" & strProduct & ".tabcfnow WHERE Tester=A.Tester AND Shoe=A.Shoe AND CFTypeID=False) '" & strDisplay & ".CFAdd',"
            If drParam.Item("param_mul") Then strSQL = strSQL & "(SELECT para" & strParamID & " FROM db_" & strProduct & ".tabcfnow WHERE Tester=A.Tester AND Shoe=A.Shoe AND CFTypeID=True) '" & strDisplay & ".CFMul',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabcfnow A "
        strSQL = strSQL & "WHERE ("
        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            Dim drTester As DataRow = dtbTester.Rows(nTester)
            Dim strTester As String = drTester.Item("Tester")
            strSQL = strSQL & "A.Tester='MT" & strTester & "' OR "
        Next
        If Right(strSQL, 4) = " OR " Then strSQL = Left(strSQL, strSQL.Length - 4) & ") "
        If InStr(strProduct, "_DCT_SDET") Then
            strSQL = strSQL & "AND A.Shoe='1' "
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Shoe "
        strSQL = strSQL & "ORDER BY A.Tester,A.Shoe;"
        Dim clsMySql As New CMySQL
        GetCFNow = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)

    End Function

    Public Function GetCFChange(ByVal strProduct As String, ByVal dtbTester As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT A.ChangeTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Shoe,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            Dim strDisplay As String = drParam.Item("param_display")
            If drParam.Item("param_add") Then strSQL = strSQL & "(SELECT para" & strParamID & " FROM db_" & strProduct & ".tabcfchange WHERE ChangeTime=A.ChangeTime AND Tester=A.Tester AND Shoe=A.Shoe AND CFTypeID=False) '" & strDisplay & ".CFAdd',"
            If drParam.Item("param_mul") Then strSQL = strSQL & "(SELECT para" & strParamID & " FROM db_" & strProduct & ".tabcfchange WHERE ChangeTime=A.ChangeTime AND Tester=A.Tester AND Shoe=A.Shoe AND CFTypeID=True) '" & strDisplay & ".CFMul',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabcfchange A "
        strSQL = strSQL & "WHERE A.ChangeTime BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "AND ("
        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            Dim drTester As DataRow = dtbTester.Rows(nTester)
            Dim strTester As String = drTester.Item("Tester")
            strSQL = strSQL & "A.Tester='MT" & strTester & "' OR "
        Next
        If Right(strSQL, 4) = " OR " Then strSQL = Left(strSQL, strSQL.Length - 4) & ") "
        If InStr(strProduct, "_DCT_SDET") Then
            strSQL = strSQL & "AND A.Shoe='1' "
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Shoe "
        strSQL = strSQL & "ORDER BY A.ChangeTime,A.Tester,A.Shoe;"
        Dim clsMySql As New CMySQL
        GetCFChange = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)

    End Function

    Private Function GetCFChangeAtTime(ByVal strProduct As String, ByVal dtbTester As DataTable, ByVal dtbParam As DataTable, ByVal dtTime As DateTime) As DataTable
        Return Nothing
    End Function

End Class

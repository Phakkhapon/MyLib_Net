
Imports MySql.Data.MySqlClient

Public Class CGetSummaryTdType
    Private m_MySqlConn As MySqlConnection
    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetSummaryTdType(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim strSearchBy As String = dtbSearch.TableName

        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT "
        If strSearchBy.ToUpper = "TESTER" Then
            strSQL = strSQL & "A.Tester,"
        End If
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot, "
        strSQL = strSQL & "COUNT(A.tag_id) Total,"
        strSQL = strSQL & "SUM(TdTyp=0) 'TdTyp 0',"
        strSQL = strSQL & "SUM(TdTyp=1) 'TdTyp 1',"
        strSQL = strSQL & "SUM(TdTyp=1.001) 'TdTyp 1.001',"
        strSQL = strSQL & "SUM(TdTyp=1.0011) 'TdTyp 1.0011',"
        strSQL = strSQL & "SUM(TdTyp=1.0101) 'TdTyp 1.0101',"
        strSQL = strSQL & "SUM(TdTyp=1.011) 'TdTyp 1.011',"
        strSQL = strSQL & "SUM(TdTyp=1.0111) 'TdTyp 1.0111',"
        strSQL = strSQL & "SUM(TdTyp=1.1) 'TdTyp 1.1',"
        strSQL = strSQL & "SUM(TdTyp=1.005) 'TdTyp 1.005',"
        strSQL = strSQL & "SUM(TdTyp=1.5) 'TdTyp 1.5',"
        strSQL = strSQL & "SUM(TdTyp=2) 'TdTyp 2',"
        strSQL = strSQL & "SUM(TdTyp=2.001) 'TdTyp 2.001',"
        strSQL = strSQL & "SUM(TdTyp=2.0011) 'TdTyp 2.0011',"
        strSQL = strSQL & "SUM(TdTyp=2.0101) 'TdTyp 2.0101',"
        strSQL = strSQL & "SUM(TdTyp=2.011) 'TdTyp 2.011',"
        strSQL = strSQL & "SUM(TdTyp=2.0111) 'TdTyp 2.0111',"
        strSQL = strSQL & "SUM(TdTyp=2.1) 'TdTyp 2.1',"
        strSQL = strSQL & "SUM(TdTyp=2.005) 'TdTyp 2.005',"
        strSQL = strSQL & "SUM(TdTyp=2.5) 'TdTyp 2.5',"
        strSQL = strSQL & "SUM(TdTyp=3) 'TdTyp 3',"
        strSQL = strSQL & "SUM(TdTyp=3.001) 'TdTyp 3.001',"
        strSQL = strSQL & "SUM(TdTyp=3.0011) 'TdTyp 3.0011',"
        strSQL = strSQL & "SUM(TdTyp=3.0101) 'TdTyp 3.0101',"
        strSQL = strSQL & "SUM(TdTyp=3.011) 'TdTyp 3.011',"
        strSQL = strSQL & "SUM(TdTyp=3.0111) 'TdTyp 3.0111',"
        strSQL = strSQL & "SUM(TdTyp=3.1) 'TdTyp 3.1',"
        strSQL = strSQL & "SUM(TdTyp=3.005) 'TdTyp 3.005',"
        strSQL = strSQL & "SUM(TdTyp=3.5) 'TdTyp 3.5',"
        strSQL = strSQL & "SUM(TdTyp=4) 'TdTyp 4',"
        strSQL = strSQL & "SUM(TdTyp=4.001) 'TdTyp 4.001',"
        strSQL = strSQL & "SUM(TdTyp=4.0011) 'TdTyp 4.0011',"
        strSQL = strSQL & "SUM(TdTyp=4.0101) 'TdTyp 4.0101',"
        strSQL = strSQL & "SUM(TdTyp=4.011) 'TdTyp 4.011',"
        strSQL = strSQL & "SUM(TdTyp=4.0111) 'TdTyp 4.0111',"
        strSQL = strSQL & "SUM(TdTyp=4.1) 'TdTyp 4.1',"
        strSQL = strSQL & "SUM(TdTyp=4.005) 'TdTyp 4.005',"
        strSQL = strSQL & "SUM(TdTyp=4.5) 'TdTyp 4.5',"
        strSQL = strSQL & "SUM(TdTyp=5) 'TdTyp 5',"
        strSQL = strSQL & "SUM(TdTyp=5.001) 'TdTyp 5.001',"
        strSQL = strSQL & "SUM(TdTyp=5.0011) 'TdTyp 5.0011',"
        strSQL = strSQL & "SUM(TdTyp=5.0101) 'TdTyp 5.0101',"
        strSQL = strSQL & "SUM(TdTyp=5.011) 'TdTyp 5.011',"
        strSQL = strSQL & "SUM(TdTyp=5.0111) 'TdTyp 5.0111',"
        strSQL = strSQL & "SUM(TdTyp=5.1) 'TdTyp 5.1',"
        strSQL = strSQL & "SUM(TdTyp=5.005) 'TdTyp 5.005',"
        strSQL = strSQL & "SUM(TdTyp=5.5) 'TdTyp 5.5' "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
        strSQL = strSQL & "WHERE A.test_time_bigint>'" & Format(dtStart, "yyyyMMddHHmmss") & "' AND A.test_time_bigint<'" & Format(dtEnd, "yyyyMMddHHmmss") & "' "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        'strSQL = strSQL & "AND spec LIKE 'R%' "
        If strSearchBy.ToUpper = "TESTER" Then
            strSQL = strSQL & "GROUP BY Tester,Spec,Lot "
        Else
            strSQL = strSQL & "GROUP BY Spec,Lot "
        End If
        strSQL = strSQL & "ORDER BY A.Tester,A.Spec,A.Lot;"

        Dim clsMySql As New CMySQL
        GetSummaryTdType = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function
End Class

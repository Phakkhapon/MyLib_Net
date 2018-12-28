
Imports MySql.Data.MySqlClient

Public Class CGetYieldIOCR

    Private m_myRawConn As MySqlConnection

    Public Sub New(ByVal myRawConn As MySqlConnection)
        m_myRawConn = myRawConn
    End Sub

    Public Function GetYieldIOCR(ByVal strProduct As String, ByVal dtbTester As DataTable, ByVal dtbHeader As DataTable, ByVal dtStart As DateTime, _
    ByVal dtEnd As DateTime) As DataTable

        Dim strSQL As String
        strSQL = "SELECT max(A.test_time) Date_Time,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & " (SELECT Count(B.tag_id) FROM db_" & strProduct & ".tabdetail_header B WHERE (B.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') AND B.Lot=A.Lot AND B.Tester=A.Tester AND B.Spec=A.Spec AND B.GradeName NOT LIKE '%ALIGN%' AND B.GradeName NOT LIKE '%NOHGA%' AND LENGTH(B.Hga_SN)=8) Total,"
        'strSQL = strSQL & " (SELECT Count(B.tag_id) FROM db_" & strProduct & ".tabdetail_header B WHERE (B.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') AND B.Lot=A.Lot AND B.Tester=A.Tester AND B.Spec=A.Spec AND B.GradeName NOT LIKE '%ALIGN%' AND B.GradeName NOT LIKE '%NOHGA%' AND B.Hga_SN NOT LIKE '%?%' AND LENGTH(B.Hga_SN)=8) AccuracyCount,"
        strSQL = strSQL & " (SELECT Count(B.tag_id) FROM db_" & strProduct & ".tabdetail_header B WHERE (B.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') AND B.Lot=A.Lot AND B.Tester=A.Tester AND B.Spec=A.Spec AND B.Hga_SN NOT LIKE '%?%' AND B.GradeName NOT LIKE '%ALIGN%' AND B.GradeName NOT LIKE '%NOHGA%' AND LENGTH(B.Hga_SN)=8) ReadabilityCount "
        strSQL = strSQL & " FROM db_" & strProduct & ".tabdetail_header A "
        strSQL = strSQL & " WHERE (A.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
        strSQL = strSQL & " AND ("
        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            If nTester <> dtbTester.Rows.Count - 1 Then
                strSQL = strSQL & " a.tester='" & dtbTester.Rows(nTester).Item("tester") & "' OR "
            Else
                strSQL = strSQL & "a.tester='" & dtbTester.Rows(nTester).Item("tester") & "') "
            End If
        Next nTester
        strSQL = strSQL & " AND A.Spec LIKE 'C%'"
        strSQL = strSQL & " AND LENGTH(A.Hga_SN)=8 "
        strSQL = strSQL & "group by a.Tester,a.Lot,a.Spec "
        strSQL = strSQL & "ORDER by a.Tester,Test_Time;"
        Dim clsGetYield As New CMySQL
        Dim dtbGetIOCR As DataTable = clsGetYield.CommandMySqlDataTable(strSQL, m_myRawConn)

        'dtbGetIOCR.Columns.Add("%Accuracy", System.Type.GetType("System.Double"))
        dtbGetIOCR.Columns.Add("%Readability", System.Type.GetType("System.Double"))
        dtbGetIOCR.Columns("%Readability").Expression = "ReadabilityCount/Total*100"
        'dtbGetIOCR.Columns.Add("%Capability", System.Type.GetType("System.Double"))
        'For nData As Integer = 0 To dtbGetIOCR.Rows.Count - 1
        '    'If dtbGetOCR.Rows(nData).Item("AccuracyCount") > dtbGetOCR.Rows(nData).Item("ReadabilityCount") Then
        '    '    dtbGetOCR.Rows(nData).Item("AccuracyCount") = dtbGetOCR.Rows(nData).Item("ReadabilityCount")
        '    'End If
        '    'dtbGetIOCR.Rows(nData).Item("%Accuracy") = dtbGetIOCR.Rows(nData).Item("AccuracyCount") / dtbGetIOCR.Rows(nData).Item("ReadabilityCount") * 100
        '    dtbGetIOCR.Rows(nData).Item("%Readability") = dtbGetIOCR.Rows(nData).Item("ReadabilityCount") / dtbGetIOCR.Rows(nData).Item("Total") * 100
        '    'dtbGetIOCR.Rows(nData).Item("%Capability") = dtbGetIOCR.Rows(nData).Item("%Accuracy") * dtbGetIOCR.Rows(nData).Item("%Readability") / 100
        'Next nData
        GetYieldIOCR = dtbGetIOCR
    End Function

End Class

Imports System.Data
Imports MySql.Data.MySqlClient


Public Class CRealProduct
    Private m_MysqlConn As MySqlConnection

    Public Sub New(ByVal MysqlConn As MySqlConnection)
        m_MysqlConn = MysqlConn
    End Sub

    Public Function GetTesterStatus() As DataTable

        Dim clsTesterStatus As New CMySQL
        Dim strSQL As String = "SELECT Tester,"
        strSQL = strSQL & "ProductName,"
        strSQL = strSQL & "MachineType,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "UpdateTime,"
        strSQL = strSQL & "IPAdr,"
        strSQL = strSQL & "CFVersion "
        strSQL = strSQL & "FROM db_parameter_mapping.tabmachinebyproduct A "
        strSQL = strSQL & "ORDER BY Tester;"
        Dim dtbTesterStatus As DataTable = clsTesterStatus.CommandMySqlDataTable(strSQL, m_MysqlConn)

        Dim dcPrime(0) As DataColumn
        dcPrime(0) = dtbTesterStatus.Columns("Tester")
        dtbTesterStatus.PrimaryKey = dcPrime

        Dim dtbProduct As DataTable = dtbTesterStatus.DefaultView.ToTable(True, "ProductName")
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item("ProductName").ToString
            If strProduct <> "" Then

                Dim dcTester() As DataRow = dtbTesterStatus.Select("ProductName='" & strProduct & "'")
                strSQL = "SELECT "
                strSQL = strSQL & "CONCAT('MT',A.Tester) Tester,"
                strSQL = strSQL & "DCT400Version,"
                strSQL = strSQL & "WTrayVersion,"
                strSQL = strSQL & "DoverHMI,"
                strSQL = strSQL & "DCTPostalVersion "
                strSQL = strSQL & "FROM "
                strSQL = strSQL & "(SELECT CONVERT(DATE_FORMAT(MAX(update_time),'%Y%m%d%H%i%s'),UNSIGNED) LastTime,Tester "
                strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg "
                strSQL = strSQL & "WHERE ("
                For nTester As Integer = 0 To dcTester.Length - 1
                    strSQL = strSQL & "Tester='" & Mid(dcTester(nTester).Item("Tester"), 3) & "' OR "
                Next nTester
                If Right(strSQL, 4) = " OR " Then strSQL = Left(strSQL, strSQL.Length - 4) & ") "
                strSQL = strSQL & "AND TestMode<>1 AND TestMode<>5 "     'ignore STD/GOS mode
                strSQL = strSQL & "GROUP BY Tester) MaxTime "
                strSQL = strSQL & "INNER JOIN db_" & strProduct & ".tabdetail_header A ON MaxTime.LastTime=A.test_time_bigint AND MaxTime.Tester=A.Tester "
                Dim dtbVersion As DataTable = clsTesterStatus.CommandMySqlDataTable(strSQL, m_MysqlConn)
                dtbTesterStatus.Merge(dtbVersion)
            End If
        Next nProduct
        GetTesterStatus = dtbTesterStatus
    End Function

    Public Function GetTesterTracking(ByVal dtbTester As DataTable) As DataTable
        Dim strSQL As String
        Dim clsTesterStatus As New CMySQL
        Dim clsProduct As New CParameterRTTCMapping(m_MysqlConn)
        Dim dtbProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductAll)
        Dim dtbTesterTracking As New DataTable

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item("Product").ToString
            strSQL = "SELECT A.Tester,"
            strSQL = strSQL & "'" & strProduct & "' ProductName,"
            strSQL = strSQL & "A.Update_time LastTestTime "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
            strSQL = strSQL & "WHERE (" 'A.Update_time>'" & strDate & "' "
            For nSearch As Integer = 0 To dtbTester.Rows.Count - 1
                If nSearch <> dtbTester.Rows.Count - 1 Then
                    strSQL = strSQL & "A.Tester='" & dtbTester.Rows(nSearch).Item("Tester") & "' OR "
                Else
                    strSQL = strSQL & "A.Tester='" & dtbTester.Rows(nSearch).Item("Tester") & "') "
                End If
            Next nSearch
            strSQL = strSQL & "AND A.Update_time=(SELECT MAX(G.Update_time) FROM db_" & strProduct & ".tabmean_avg G WHERE A.Tester=G.Tester) "
            Dim dtbTesterTmp As DataTable = clsTesterStatus.CommandMySqlDataTable(strSQL, m_MysqlConn)
            dtbTesterTracking.Merge(dtbTesterTmp)
        Next nProduct
        Dim dtbData As DataTable
        dtbTesterTracking.DefaultView.Sort = "LastTestTime ASC"
        dtbData = dtbTesterTracking.DefaultView.ToTable
        GetTesterTracking = dtbData
    End Function

End Class

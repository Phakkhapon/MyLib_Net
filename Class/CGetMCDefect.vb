Imports MySql.Data.MySqlClient

Public Class CGetMCDefect

    Private m_mySqlConn As MySqlConnection
    Public Sub New(ByVal MyWexSQL As MySqlConnection)
        m_mySqlConn = MyWexSQL
    End Sub

    Public Function GetMCDefect(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal nSearchOption As enumSearchOption) As DataTable
        'Dim clsGetYield As New CGetYieldRawData(m_mySqlConn)
        Dim dtbDefectData As DataTable = GetYieldDefect(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nSearchOption, enumGradeOption.eGradeAll)
        Dim dtbDefectMapping As DataTable = GetAllMCCode()

        'Count GradeU,V
        dtbDefectData.Columns.Add("cntRejectLow", System.Type.GetType("System.Int32"))
        dtbDefectData.Columns.Add("GradeU", System.Type.GetType("System.Int32"))
        dtbDefectData.Columns.Add("GradeV", System.Type.GetType("System.Int32"))

        Dim strGradeU As String = ""
        Dim strGradeV As String = ""

        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim strColName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If dtbDefectMapping.Rows(nMapping).Item("ValueType").ToString.ToUpper = "GRADEU" Then
                'nGradeU = nGradeU + dtbDefectData.Rows(nRow).Item(strColName)
                If nMapping = 0 Then
                    strGradeU = strColName
                Else
                    strGradeU = strGradeU & "+" & strColName
                End If
            ElseIf dtbDefectMapping.Rows(nMapping).Item("ValueType").ToString.ToUpper = "GRADEV" Then
                If nMapping = 0 Then
                    strGradeV = strColName
                Else
                    strGradeV = strGradeV & "+" & strColName
                End If
            End If
        Next nMapping
        dtbDefectData.Columns("GradeU").Expression = strGradeU
        dtbDefectData.Columns("GradeV").Expression = strGradeV
        dtbDefectData.Columns("cntRejectLow").Expression = "GradeU+GradeV"

        Dim nStart As Integer = dtbDefectData.Columns("UPH").Ordinal
        dtbDefectData.Columns.Item("cntRejectLow").SetOrdinal(nStart + 1)
        dtbDefectData.Columns.Item("GradeU").SetOrdinal(nStart + 2)
        dtbDefectData.Columns.Item("GradeV").SetOrdinal(nStart + 3)
        Return dtbDefectData
    End Function

    Public Function GetAllMCCode() As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM db_parameter_mapping.tabmcdefect "
        strSQL = strSQL & " WHERE IsEnable=1;"
        Dim clsSQL As New CMySQL
        GetAllMCCode = clsSQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

    Public Function GetMCCodeGradeNone() As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM db_parameter_mapping.tabmcdefect A"
        strSQL = strSQL & " WHERE A.IsEnable=1 "
        strSQL = strSQL & " AND (A.ValueType ='GradeNone' "
        strSQL = strSQL & " OR A.MCDefectName='BA1');"
        Dim clsSQL As New CMySQL
        GetMCCodeGradeNone = clsSQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

    Private Function GetYieldDefect(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
  ByVal dtEnd As DateTime, ByVal nSearchOption As enumSearchOption, ByVal nGradeOption As enumGradeOption) As DataTable

        If nSearchOption = enumSearchOption.eSearchByTester Then
            'GetYieldDefect = GetYieldSummaryByTester(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
            GetYieldDefect = GetMechanicalSummaryByTester(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        ElseIf nSearchOption = enumSearchOption.eSearchByLot Then
            'GetYieldDefect = GetYieldSummaryByLot(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
            GetYieldDefect = GetMechanicalSummaryByLot(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        ElseIf nSearchOption = enumSearchOption.eSearchBySpec Then
            'GetYieldDefect = GetYieldSummaryBySpec(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
            GetYieldDefect = GetMechanicalSummaryBySpec(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        Else
            GetYieldDefect = Nothing
        End If

    End Function
#Region " Old Get Mechanical Defect By Tester"
    Private Function GetYieldSummaryByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        'strSQL = strSQL & "(SELECT D.CGALot FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) CGALot,"
        strSQL = strSQL & "F.CGALot,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "F.OprID,"
        strSQL = strSQL & "F.Assy,"
        strSQL = strSQL & "F.MediaSN,"
        strSQL = strSQL & "F.Grade_rev,"
        'strSQL = strSQL & "(SELECT D.OprID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) OprID,"
        'strSQL = strSQL & "(SELECT D.Assy FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Assy,"
        'strSQL = strSQL & "(SELECT D.MediaSN FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) MediaSN,"
        'strSQL = strSQL & "(SELECT D.Grade_rev FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Grade_rev,"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot) 'Lot.Total',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot) 'Lot.Pass',"
        'strSQL = strSQL & "SUM(D.TotalHGA) 'Total.Lot',"
        'strSQL = strSQL & "SUM(D.TotalPass) 'Pass.Lot',"

        strSQL = strSQL & "SUM(A.TotalHGA) 'Tester.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Tester.Pass',"

        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If

        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
            End If
        Next nMapping
        strSQL = strSQL & "F.WorkID "
        'strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header F ON F.test_time_bigint=CONVERT(DATE_FORMAT(A.update_time,'%Y%m%d%H%i%s'),UNSIGNED) AND F.Tester=A.Tester AND F.Shoe=A.Shoe "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Lot,A.Spec,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Tester,A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByTester As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        dtbGetYieldByTester.Columns.Add("Yield.Tester", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester").Expression = "Tester.Pass/Tester.Total*100"
        dtbGetYieldByTester.Columns.Add("Yield.Lot", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Lot").Expression = "Lot.Pass/Lot.Total*100"
        Dim nStart As Integer = dtbGetYieldByTester.Columns("Tester.Pass").Ordinal
        dtbGetYieldByTester.Columns.Item("Yield.Lot").SetOrdinal(nStart + 1)
        dtbGetYieldByTester.Columns.Item("Yield.Tester").SetOrdinal(nStart + 2)
        dtbGetYieldByTester.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldByTester.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldByTester.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldByTester.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        dtbGetYieldByTester.Columns.Add("CavityType", GetType(String), "IIF(CGALot LIKE 'P%','CGA',IIF(CGALot LIKE 'M%','MEMs',NULL))")
        dtbGetYieldByTester.Columns("CavityType").SetOrdinal(dtbGetYieldByTester.Columns("CGALot").Ordinal + 1)
        GetYieldSummaryByTester = dtbGetYieldByTester
    End Function
#End Region

#Region " New Get Mechanical Defect By Tester"
    Private Function GetMechanicalSummaryByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        ''strSQL = strSQL & "(SELECT D.CGALot FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) CGALot,"
        strSQL = strSQL & "F.CGALot,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "F.OprID,"
        strSQL = strSQL & "F.Assy,"
        strSQL = strSQL & "F.MediaSN,"
        strSQL = strSQL & "F.Grade_rev,"
        ''strSQL = strSQL & "(SELECT D.OprID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) OprID,"
        ''strSQL = strSQL & "(SELECT D.Assy FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Assy,"
        ''strSQL = strSQL & "(SELECT D.MediaSN FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) MediaSN,"
        ''strSQL = strSQL & "(SELECT D.Grade_rev FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Grade_rev,"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect K WHERE K.Lot=A.Lot AND K.Tester = A.Tester ) 'Tester.Total',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabsummary_hgadefect K WHERE K.Lot=A.Lot AND K.Tester = A.Tester) 'Tester.Pass',"
        ''strSQL = strSQL & "SUM(D.TotalHGA) 'Total.Lot',"
        ''strSQL = strSQL & "SUM(D.TotalPass) 'Pass.Lot',"

        '' strSQL = strSQL & "SUM(A.TotalHGA) 'Tester.Total',"
        ''strSQL = strSQL & "SUM(A.TotalPass) 'Tester.Pass',"

        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If

        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
            End If
        Next nMapping
        strSQL = strSQL & "F.WorkID "
        ''strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header F ON F.test_time_bigint=CONVERT(DATE_FORMAT(A.update_time,'%Y%m%d%H%i%s'),UNSIGNED) AND F.Tester=A.Tester AND F.Shoe=A.Shoe "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Lot,A.Spec,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Tester,A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByTester As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        dtbGetYieldByTester.Columns.Add("Yield.Tester", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester").Expression = "Tester.Pass/Tester.Total*100"
        dtbGetYieldByTester.Columns.Add("Yield.Lot") ', System.Type.GetType("System.Double"))
        ''dtbGetYieldByTester.Columns("Yield.Lot").Expression = "Lot.Pass/Lot.Total*100"
        Dim nStart As Integer = dtbGetYieldByTester.Columns("Tester.Pass").Ordinal
        dtbGetYieldByTester.Columns.Item("Yield.Lot").SetOrdinal(nStart + 1)
        dtbGetYieldByTester.Columns.Item("Yield.Tester").SetOrdinal(nStart + 2)
        dtbGetYieldByTester.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldByTester.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldByTester.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldByTester.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        dtbGetYieldByTester.Columns.Add("CavityType", GetType(String), "IIF(CGALot LIKE 'P%','CGA',IIF(CGALot LIKE 'M%','MEMs',NULL))")
        dtbGetYieldByTester.Columns("CavityType").SetOrdinal(dtbGetYieldByTester.Columns("CGALot").Ordinal + 1)
        GetMechanicalSummaryByTester = dtbGetYieldByTester

    End Function

#End Region

#Region " Old Get Mechanical Defect by lot"
    Private Function GetYieldSummaryByLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "(SELECT D.CGALot FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) CGALot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "(SELECT D.Assy FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Assy,"
        strSQL = strSQL & "(SELECT D.Grade_rev FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Grade_rev,"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "SUM(A.TotalHGA) 'Lot.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Lot.Pass',"
        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If
        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
            End If
        Next nMapping
        strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Lot,A.Spec,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Lot,A.Tester,A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByLot As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)

        dtbGetYieldByLot.Columns.Add("Lot.Yield", System.Type.GetType("System.Double"))
        dtbGetYieldByLot.Columns("Lot.Yield").Expression = "Lot.Pass/Lot.Total*100"
        Dim nStart As Integer = dtbGetYieldByLot.Columns("Lot.Pass").Ordinal
        dtbGetYieldByLot.Columns.Item("Lot.Yield").SetOrdinal(nStart + 1)
        dtbGetYieldByLot.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldByLot.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldByLot.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldByLot.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetYieldSummaryByLot = dtbGetYieldByLot
    End Function
#End Region

#Region "New Get Mechanical Defect by Lot"
    Private Function GetMechanicalSummaryByLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        'strSQL = strSQL & "(SELECT D.CGALot FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) CGALot,"
        strSQL = strSQL & "F.CGALot,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "F.OprID,"
        strSQL = strSQL & "F.Assy,"
        strSQL = strSQL & "F.MediaSN,"
        strSQL = strSQL & "F.Grade_rev,"
        'strSQL = strSQL & "(SELECT D.OprID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) OprID,"
        'strSQL = strSQL & "(SELECT D.Assy FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Assy,"
        'strSQL = strSQL & "(SELECT D.MediaSN FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) MediaSN,"
        'strSQL = strSQL & "(SELECT D.Grade_rev FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) Grade_rev,"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect K WHERE K.Lot=A.Lot) 'Lot.Total',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabsummary_hgadefect K WHERE K.Lot=A.Lot) 'Lot.Pass',"
        'strSQL = strSQL & "SUM(D.TotalHGA) 'Total.Lot',"
        'strSQL = strSQL & "SUM(D.TotalPass) 'Pass.Lot',"

        'strSQL = strSQL & "SUM(A.TotalHGA) 'Tester.Total',"
        'strSQL = strSQL & "SUM(A.TotalPass) 'Tester.Pass',"

        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If

        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") '" & strDefectName & "',"
            End If
        Next nMapping
        strSQL = strSQL & "F.WorkID "
        'strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header F ON F.test_time_bigint=CONVERT(DATE_FORMAT(A.update_time,'%Y%m%d%H%i%s'),UNSIGNED) AND F.Tester=A.Tester AND F.Shoe=A.Shoe "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Tester,A.Lot,A.Spec,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Tester,A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByLot As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)

        dtbGetYieldByLot.Columns.Add("Lot.Yield", System.Type.GetType("System.Double"))
        dtbGetYieldByLot.Columns("Lot.Yield").Expression = "Lot.Pass/Lot.Total*100"
        Dim nStart As Integer = dtbGetYieldByLot.Columns("Lot.Pass").Ordinal
        dtbGetYieldByLot.Columns.Item("Lot.Yield").SetOrdinal(nStart + 1)
        dtbGetYieldByLot.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldByLot.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldByLot.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldByLot.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetMechanicalSummaryByLot = dtbGetYieldByLot
    End Function

#End Region

#Region " Old Get Mechanical Defect By Spec"
    Private Function GetYieldSummaryBySpec(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "SUM(A.TotalHGA) 'Spec.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Spec.Pass',"
        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If
        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
            End If
        Next nMapping
        strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Spec,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Spec;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldBySpec As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)

        dtbGetYieldBySpec.Columns.Add("Spec.Yield", System.Type.GetType("System.Double"))
        dtbGetYieldBySpec.Columns("Spec.Yield").Expression = "Spec.Pass/Spec.Total*100"
        Dim nStart As Integer = dtbGetYieldBySpec.Columns("Spec.Pass").Ordinal
        dtbGetYieldBySpec.Columns.Item("Spec.Yield").SetOrdinal(nStart + 1)
        dtbGetYieldBySpec.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldBySpec.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldBySpec.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldBySpec.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetYieldSummaryBySpec = dtbGetYieldBySpec
    End Function

#End Region

#Region "New Get Mechanical Defect by Spec"

    Private Function GetMechanicalSummaryBySpec(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_mySqlConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "B.SliderSite," ' Original  "A.SliderSite"

        If InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(SUM(K.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSQL = strSQL & "3600/(11+SUM(K.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "SUM(A.TotalHGA) 'Spec.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Spec.Pass',"
        If nGradeOption = enumGradeOption.eGradeNone Then
            dtbDefectMapping = clsMCDefect.GetMCCodeGradeNone()
        End If
        For nMapping As Integer = 0 To dtbDefectMapping.Rows.Count - 1
            Dim nCodeID As Integer = dtbDefectMapping.Rows(nMapping).Item("MCCodeID")
            Dim strDefectName As String = dtbDefectMapping.Rows(nMapping).Item("MCDefectName").ToString
            If nGradeOption = enumGradeOption.eUnloadDefect Then
                If InStr(strDefectName, "PASS", CompareMethod.Text) Then
                    strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
                End If
            Else
                strSQL = strSQL & "SUM(C.Defect" & nCodeID & ") " & strDefectName & ","
            End If
        Next nMapping
        strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Lot=A.Lot AND D.Spec=A.Spec LIMIT 0,1) WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabsummary_hgadefect A "  '  "Original tabmean_n A"
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_avg K USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,Lot,spec,shoe) "
        strSQL = strSQL & " WHERE (A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "B.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "  ' Original "A.SliderSite="
                Else
                    strSQL = strSQL & "B.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "GROUP BY A.Spec,B.SliderSite "  ' Old coding "GROUP BY A.Spec,B.SliderSite "  
        strSQL = strSQL & "ORDER BY A.Spec;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldBySpec As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)

        dtbGetYieldBySpec.Columns.Add("Spec.Yield", System.Type.GetType("System.Double"))
        dtbGetYieldBySpec.Columns("Spec.Yield").Expression = "Spec.Pass/Spec.Total*100"
        Dim nStart As Integer = dtbGetYieldBySpec.Columns("Spec.Pass").Ordinal
        dtbGetYieldBySpec.Columns.Item("Spec.Yield").SetOrdinal(nStart + 1)
        dtbGetYieldBySpec.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldBySpec.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldBySpec.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldBySpec.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetMechanicalSummaryBySpec = dtbGetYieldBySpec
    End Function

#End Region
End Class

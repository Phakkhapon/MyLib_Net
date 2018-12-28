Imports MySql.Data.MySqlClient

Public Class CGetYieldRawData
    Private m_myRawConn As MySqlConnection

    Public Sub New(ByVal myRawConn As MySqlConnection)
        m_myRawConn = myRawConn
    End Sub

    Public Function GetYield2Lot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim clsParam As New CParameterRTTCMapping(m_myRawConn)
        Dim dtbMCDefect As DataTable = clsParam.GetMCDefectMapping(CParameterRTTCMapping.enuMCGradeType.eGradeAll)

        Dim drPassBin1() As DataRow = dtbMCDefect.Select("MCDefectName='PassBin1'")
        Dim nIDPassBin1 As Integer = -1
        If drPassBin1.Length > 0 Then nIDPassBin1 = drPassBin1(0).Item("MCCodeID")

        Dim drPassBin2() As DataRow = dtbMCDefect.Select("MCDefectName='PassBin2'")
        Dim nIDPassBin2 As Integer = -1
        If drPassBin2.Length > 0 Then nIDPassBin2 = drPassBin2(0).Item("MCCodeID")

        Dim drPassBin3() As DataRow = dtbMCDefect.Select("MCDefectName='PassBin3'")
        Dim nIDPassBin3 As Integer = -1
        If drPassBin3.Length > 0 Then nIDPassBin3 = drPassBin3(0).Item("MCCodeID")

        Dim drPassBin4() As DataRow = dtbMCDefect.Select("MCDefectName='PassBin4'")
        Dim nIDPassBin4 As Integer = -1
        If drPassBin4.Length > 0 Then nIDPassBin4 = drPassBin4(0).Item("MCCodeID")

        Dim strSearchBy As String = dtbSearch.TableName

        Dim strTimeCondition As String = "A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
        Dim strSQL As String = ""
        strSQL = "SELECT A.Update_Time,"
        If strSearchBy.ToUpper = "TESTER" Then
            strSQL = strSQL & "A.Tester,"
        End If
        'strSQL = strSQL & "A.Spec,"
        'strSQL = strSQL & "A.Lot,"
        'strSQL = strSQL & "SUM(A.TotalHGA) 'TotalHGA.All',"
        'strSQL = strSQL & "SUM(A.TotalPass) 'TotalHGA.Pass',"
        'strSQL = strSQL & "B.TotalHGA 'TotalHGA.S1',"
        'strSQL = strSQL & "C.TotalHGA 'TotalHGA.S2',"
        'strSQL = strSQL & "B.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin1 & " 'TotalPass.BinA',"
        'strSQL = strSQL & "B.Defect" & nIDPassBin1 & " 'TotalPass.BinA.S1',"
        'strSQL = strSQL & "C.Defect" & nIDPassBin1 & " 'TotalPass.BinA.S2',"
        'strSQL = strSQL & "B.Defect" & nIDPassBin1 & "+B.Defect" & nIDPassBin2 & "+B.Defect" & nIDPassBin3 & "+B.Defect" & nIDPassBin4 & "+C.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin2 & "+C.Defect" & nIDPassBin3 & "+C.Defect" & nIDPassBin4 & "  'TotalPass.BinAll',"
        'strSQL = strSQL & "B.Defect" & nIDPassBin1 & "+B.Defect" & nIDPassBin2 & "+B.Defect" & nIDPassBin3 & "+B.Defect" & nIDPassBin4 & " 'TotalPass.BinAll.S1',"
        'strSQL = strSQL & "C.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin2 & "+C.Defect" & nIDPassBin3 & "+C.Defect" & nIDPassBin4 & " 'TotalPass.BinAll.S2',"
        'Else
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "SUM(A.TotalHGA) 'TotalHGA.All',"
        strSQL = strSQL & "SUM(A.TotalPass) 'TotalHGA.Pass',"
        strSQL = strSQL & "SUM(DISTINCT B.TotalHGA) 'TotalHGA.S1',"
        strSQL = strSQL & "SUM(DISTINCT C.TotalHGA) 'TotalHGA.S2',"
        strSQL = strSQL & "SUM(DISTINCT B.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin1 & ") 'TotalPass.BinA',"
        strSQL = strSQL & "SUM(DISTINCT B.Defect" & nIDPassBin1 & ") 'TotalPass.BinA.S1',"
        strSQL = strSQL & "SUM(DISTINCT C.Defect" & nIDPassBin1 & ") 'TotalPass.BinA.S2',"
        strSQL = strSQL & "SUM(DISTINCT B.Defect" & nIDPassBin1 & "+B.Defect" & nIDPassBin2 & "+B.Defect" & nIDPassBin3 & "+B.Defect" & nIDPassBin4 & "+C.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin2 & "+C.Defect" & nIDPassBin3 & "+C.Defect" & nIDPassBin4 & ")  'TotalPass.BinAll',"
        strSQL = strSQL & "SUM(DISTINCT B.Defect" & nIDPassBin1 & "+B.Defect" & nIDPassBin2 & "+B.Defect" & nIDPassBin3 & "+B.Defect" & nIDPassBin4 & ") 'TotalPass.BinAll.S1',"
        strSQL = strSQL & "SUM(DISTINCT C.Defect" & nIDPassBin1 & "+C.Defect" & nIDPassBin2 & "+C.Defect" & nIDPassBin3 & "+C.Defect" & nIDPassBin4 & ") 'TotalPass.BinAll.S2',"
        'End If
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabsummary_hgadefect A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect B ON A.Tester=B.Tester AND A.Spec=B.Spec AND A.Lot=B.Lot AND B.Shoe='1' "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C ON A.Tester=C.Tester AND A.Spec=C.Spec AND A.Lot=C.Lot AND C.Shoe='2' "
        strSQL = strSQL & "WHERE " & strTimeCondition
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & "AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        strSQL = strSQL & "AND (LEFT(A.Spec,1)='B' OR LEFT(A.Spec,1)='F' OR LEFT(A.Spec,1)='R' OR LEFT(A.Spec,1)='T') "
        strSQL = strSQL & "GROUP BY "
        If strSearchBy.ToUpper = "TESTER" Then
            strSQL = strSQL & "A.Tester,"
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "A.Lot;"
        ElseIf strSearchBy.ToUpper = "LOT" Then
            strSQL = strSQL & "Spec,"
            strSQL = strSQL & "Lot;"
        ElseIf strSearchBy.ToUpper = "SPEC" Then
            strSQL = strSQL & "Spec;"
        End If
        Dim clsMySql As New CMySQL
        Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_myRawConn)

        'dtbData.Columns.Add("Yield." & strSearchBy, Type.GetType("System.Double"), "TotalHGA.Pass/TotalHGA.All*100")
        dtbData.Columns.Add("Yield.BinA." & strSearchBy, Type.GetType("System.Double"), "TotalPass.BinA/TotalHGA.All*100")
        dtbData.Columns.Add("Yield.BinA." & strSearchBy & ".S1", Type.GetType("System.Double"), "TotalPass.BinA.S1/TotalHGA.S1*100")
        dtbData.Columns.Add("Yield.BinA." & strSearchBy & ".S2", Type.GetType("System.Double"), "TotalPass.BinA.S2/TotalHGA.S2*100")

        dtbData.Columns.Add("DeltaYield.BinA." & strSearchBy & ".S1", Type.GetType("System.Double"), "Yield.BinA." & strSearchBy & ".S1-Yield.BinA." & strSearchBy)
        dtbData.Columns.Add("DeltaYield.BinA." & strSearchBy & ".S2", Type.GetType("System.Double"), "Yield.BinA." & strSearchBy & ".S2-Yield.BinA." & strSearchBy)

        dtbData.Columns.Add("Yield.BinAll." & strSearchBy, Type.GetType("System.Double"), "TotalPass.BinAll/TotalHGA.All*100")
        dtbData.Columns.Add("Yield.BinAll." & strSearchBy & ".S1", Type.GetType("System.Double"), "TotalPass.BinAll.S1/TotalHGA.S1*100")
        dtbData.Columns.Add("Yield.BinAll." & strSearchBy & ".S2", Type.GetType("System.Double"), "TotalPass.BinAll.S2/TotalHGA.S2*100")

        dtbData.Columns.Add("DeltaYield.BinAll." & strSearchBy & "S1", Type.GetType("System.Double"), "Yield.BinAll." & strSearchBy & ".S2-Yield.BinAll." & strSearchBy)
        dtbData.Columns.Add("DeltaYield.BinAll." & strSearchBy & ".S2", Type.GetType("System.Double"), "Yield.BinAll." & strSearchBy & ".S2-Yield.BinAll." & strSearchBy)

        GetYield2Lot = dtbData
    End Function

    Public Function GetYield(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
   ByVal dtEnd As DateTime, ByVal nSearchOption As enumSearchOption, ByVal nGradeOption As enumGradeOption) As DataTable

        If nSearchOption = enumSearchOption.eSearchByTester Then
            GetYield = GetYieldSummaryByTester(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        ElseIf nSearchOption = enumSearchOption.eSearchByLot Then
            GetYield = GetYieldSummaryByLot(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        ElseIf nSearchOption = enumSearchOption.eSearchBySpec Then
            GetYield = GetYieldSummaryBySpec(strProduct, dtbSearch, dtbSliderSite, dtStart, dtEnd, nGradeOption)
        Else
            GetYield = Nothing
        End If
    End Function

    Private Function GetYieldSummaryByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_myRawConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "F.StandardLot,"
        'strSQL = strSQL & "IF(A.Spec LIKE 'C%',"
        'strSQL = strSQL & "(SELECT LotName FROM std_standard.tabstandard_hga LEFT JOIN std_standard.tabtray USING(TrayID) LEFT JOIN std_standard.tablot USING(LotID) WHERE Hga_SN=F.Hga_SN AND TrayName=F.TrayID LIMIT 0,1),"
        'strSQL = strSQL & "NULL) STD_Lot,"

        'strSQL = strSQL & "LEFT JOIN std_standard.tabtray T USING(trayID) "
        'strSQL = strSQL & "WHERE S.Hga_sn=(SELECT Hga_sn FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) AND "
        'strSQL = strSQL & "T.TrayName=(SELECT TrayID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) LIMIT 0,1),NULL) 'STD_lot',"

        'strSQL = strSQL & "(SELECT D.CGALot FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) CGALot,"
        strSQL = strSQL & "F.CGALot,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"

        'strSQL = strSQL & "(SELECT D.OprID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) OprID,"
        'strSQL = strSQL & "(SELECT D.Assy FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) Assy,"
        'strSQL = strSQL & "(SELECT D.MediaSN FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) MediaSN,"
        'strSQL = strSQL & "(SELECT D.Grade_rev FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) Grade_rev,"
        strSQL = strSQL & "F.OprID,"
        strSQL = strSQL & "F.Assy,"
        strSQL = strSQL & "F.MediaSN,"
        strSQL = strSQL & "F.Grade_rev,"
        strSQL = strSQL & "IF(LENGTH(F.TrayID)=8,'Tray500','NormalTray') 'TrayType',"
        strSQL = strSQL & "F.WTrayVersion,"
        If InStr(strProduct.ToCharArray, "_DUAL_SDET") Then
            strSQL = strSQL & "3600/(M.CycleTime)*2 UPH,"
        ElseIf InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/(M.CycleTime) UPH,"
        Else
            strSQL = strSQL & "3600/(10+M.CycleTime) UPH,"
        End If

        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot) 'Lot.Total',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot) 'Lot.Pass',"
        strSQL = strSQL & "SUM(A.TotalHGA) 'Tester.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Tester.Pass',"

        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec AND Shoe='1') 'Tester.Total.Shoe1',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec AND Shoe='1') 'Tester.Pass.Shoe1',"
        strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec AND Shoe='2') 'Tester.Total.Shoe2',"
        strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec AND Shoe='2') 'Tester.Pass.Shoe2',"

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
        'strSQL = strSQL & "(SELECT D.WorkID FROM db_" & strProduct & ".tabdetail_header D WHERE D.Tester=A.Tester AND D.Spec=A.Spec AND D.Lot=A.Lot LIMIT 0,1) WorkID "
        strSQL = strSQL & "F.WorkID "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmedian M USING(tester,Spec,lot) "
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
        strSQL = strSQL & "GROUP BY A.Tester,A.Spec,A.Lot,A.SliderSite "
        strSQL = strSQL & "ORDER BY A.Tester,A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByTester As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_myRawConn)
        dtbGetYieldByTester.Columns.Add("Yield.Lot", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Lot").Expression = "Lot.Pass/Lot.Total*100"

        dtbGetYieldByTester.Columns.Add("Yield.Tester", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester").Expression = "Tester.Pass/Tester.Total*100"
        dtbGetYieldByTester.Columns.Add("Yield.Tester.Bin1", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester.Bin1").Expression = "PassBin1/Tester.Total*100"

        dtbGetYieldByTester.Columns.Add("Yield.Tester.Shoe1", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester.Shoe1").Expression = "[Tester.Pass.Shoe1]/[Tester.Total.Shoe1]*100"
        dtbGetYieldByTester.Columns.Add("Yield.Tester.Shoe2", System.Type.GetType("System.Double"))
        dtbGetYieldByTester.Columns("Yield.Tester.Shoe2").Expression = "[Tester.Pass.Shoe2]/[Tester.Total.Shoe2]*100"

        Dim nStart As Integer = dtbGetYieldByTester.Columns("Tester.Pass.Shoe2").Ordinal
        dtbGetYieldByTester.Columns.Item("Yield.Lot").SetOrdinal(nStart + 1)
        dtbGetYieldByTester.Columns.Item("Yield.Tester").SetOrdinal(nStart + 2)
        dtbGetYieldByTester.Columns.Item("Yield.Tester.Bin1").SetOrdinal(nStart + 3)
        dtbGetYieldByTester.Columns.Item("Yield.Tester.Shoe1").SetOrdinal(nStart + 4)
        dtbGetYieldByTester.Columns.Item("Yield.Tester.Shoe2").SetOrdinal(nStart + 5)
        dtbGetYieldByTester.Columns.Item("PassBin1").SetOrdinal(nStart + 6)
        dtbGetYieldByTester.Columns.Item("PassBin2").SetOrdinal(nStart + 7)
        dtbGetYieldByTester.Columns.Item("PassBin3").SetOrdinal(nStart + 8)
        dtbGetYieldByTester.Columns.Item("PassBin4").SetOrdinal(nStart + 9)

        Dim dtbAddPettern As DataTable = CheckPattern(dtbGetYieldByTester, dtbSearch)
        GetYieldSummaryByTester = dtbAddPettern
    End Function

    Private Function CheckPattern(ByVal dtbXLotData As DataTable, ByVal dtbTester As DataTable) As DataTable
        dtbXLotData.Columns.Add("DeltaYield", Type.GetType("System.Double"), "Yield.Tester-Yield.Lot")
        dtbXLotData.Columns.Add("ChkPattern", Type.GetType("System.UInt16"))
        dtbXLotData.Columns.Add("Bad_Tst")

        Dim nStart As Integer = dtbXLotData.Columns("Yield.Tester").Ordinal
        dtbXLotData.Columns.Item("DeltaYield").SetOrdinal(nStart + 1)
        dtbXLotData.Columns.Item("ChkPattern").SetOrdinal(nStart + 2)
        dtbXLotData.Columns.Item("Bad_Tst").SetOrdinal(nStart + 3)

        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            Dim strTester As String = dtbTester.Rows(nTester).Item("Tester")
            Dim dtrSameTester() As DataRow = dtbXLotData.Select("[Tester]='" & strTester & "' AND (Spec LIKE 'R%' OR Spec LIKE 'T%')", "UpdateTime ASC")
            If dtrSameTester.Length > 0 Then
                dtrSameTester(0).Item("ChkPattern") = 1
                For nSelect As Integer = 1 To dtrSameTester.Length - 1
                    Dim dblLastDeltaYield As Double = dtrSameTester(nSelect - 1).Item("DeltaYield")
                    Dim dblDeltaYield As Double = dtrSameTester(nSelect).Item("DeltaYield")
                    If (dblLastDeltaYield > 0 And dblDeltaYield > 0) Or (dblLastDeltaYield < 0 And dblDeltaYield < 0) Then
                        dtrSameTester(nSelect).Item("ChkPattern") = dtrSameTester(nSelect - 1).Item("ChkPattern") + 1
                    Else
                        dtrSameTester(nSelect).Item("ChkPattern") = 1
                    End If
                    If dtrSameTester(nSelect).Item("ChkPattern") > 4 Then dtrSameTester(nSelect).Item("Bad_Tst") = "Bad"
                Next nSelect
            End If
        Next
        CheckPattern = dtbXLotData
    End Function

    Private Function GetYieldSummaryByLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_myRawConn)
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

        If InStr(strProduct.ToCharArray, "_DUAL_SDET") Then
            strSQL = strSQL & "3600/(AVG(M.CycleTime))*2 UPH,"
        ElseIf InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/AVG(M.CycleTime) UPH,"
        Else
            strSQL = strSQL & "3600/(10+AVG(M.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "SUM(A.TotalHGA) 'Lot.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Lot.Pass',"

        'strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot AND Shoe='1') ""Total.Shoe1"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot AND Shoe='1') ""Pass.Shoe1"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot AND Shoe='2') ""Total.Shoe2"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Lot=A.Lot AND Shoe='2') ""Pass.Shoe2"","

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
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmedian M USING(tester,Spec,lot) "
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
        strSQL = strSQL & "ORDER BY A.Update_time;"
        Dim clsMySql As New CMySQL
        Dim dtbGetYieldByLot As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_myRawConn)

        dtbGetYieldByLot.Columns.Add("Yield.Lot", System.Type.GetType("System.Double"))
        dtbGetYieldByLot.Columns("Yield.Lot").Expression = "Lot.Pass/Lot.Total*100"

        dtbGetYieldByLot.Columns.Add("Yield.Lot.Bin1", System.Type.GetType("System.Double"))
        dtbGetYieldByLot.Columns("Yield.Lot.Bin1").Expression = "PassBin1/Lot.Total*100"

        Dim nStart As Integer = dtbGetYieldByLot.Columns("Lot.Pass").Ordinal
        dtbGetYieldByLot.Columns.Item("Yield.Lot").SetOrdinal(nStart + 1)
        dtbGetYieldByLot.Columns.Item("Yield.Lot.Bin1").SetOrdinal(nStart + 2)
        dtbGetYieldByLot.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldByLot.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldByLot.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldByLot.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetYieldSummaryByLot = dtbGetYieldByLot
    End Function

    Private Function GetYieldSummaryBySpec(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime, _
ByVal dtEnd As DateTime, ByVal nGradeOption As enumGradeOption) As DataTable

        Dim clsMCDefect As New CGetMCDefect(m_myRawConn)
        Dim dtbDefectMapping As DataTable = clsMCDefect.GetAllMCCode()
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_time UpdateTime,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.SliderSite,"

        If InStr(strProduct.ToCharArray, "_DUAL_SDET") Then
            strSQL = strSQL & "3600/(AVG(M.CycleTime))*2 UPH,"
        ElseIf InStr(strProduct.ToUpper, "_SDET") Then
            strSQL = strSQL & "3600/AVG(M.CycleTime) UPH,"
        Else
            strSQL = strSQL & "3600/(10+AVG(M.CycleTime)) UPH,"
        End If

        strSQL = strSQL & "SUM(A.TotalHGA) 'Spec.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Spec.Pass',"

        'strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Spec=A.Spec AND Shoe='1') ""Total.Shoe1"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Spec=A.Spec AND Shoe='1') ""Pass.Shoe1"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalHGA) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Spec=A.Spec AND Shoe='2') ""Total.Shoe2"","
        'strSQL = strSQL & "(SELECT SUM(K.TotalPass) FROM db_" & strProduct & ".tabmean_avg K WHERE K.Spec=A.Spec AND Shoe='2') ""Pass.Shoe2"","

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
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect C USING(tester,spec,Lot,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmedian M USING(tester,Spec,lot) "
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
        Dim dtbGetYieldBySpec As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_myRawConn)

        dtbGetYieldBySpec.Columns.Add("Yield.Spec", System.Type.GetType("System.Double"))
        dtbGetYieldBySpec.Columns("Yield.Spec").Expression = "Spec.Pass/Spec.Total*100"
        dtbGetYieldBySpec.Columns.Add("Yield.Bin1", System.Type.GetType("System.Double"))
        dtbGetYieldBySpec.Columns("Yield.Bin1").Expression = "PassBin1/Spec.Total*100"

        Dim nStart As Integer = dtbGetYieldBySpec.Columns("Spec.Pass").Ordinal
        dtbGetYieldBySpec.Columns.Item("Yield.Spec").SetOrdinal(nStart + 1)
        dtbGetYieldBySpec.Columns.Item("Yield.Bin1").SetOrdinal(nStart + 1)
        dtbGetYieldBySpec.Columns.Item("PassBin1").SetOrdinal(nStart + 3)
        dtbGetYieldBySpec.Columns.Item("PassBin2").SetOrdinal(nStart + 4)
        dtbGetYieldBySpec.Columns.Item("PassBin3").SetOrdinal(nStart + 5)
        dtbGetYieldBySpec.Columns.Item("PassBin4").SetOrdinal(nStart + 6)

        GetYieldSummaryBySpec = dtbGetYieldBySpec
    End Function
End Class

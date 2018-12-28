
Imports MySql.Data.MySqlClient

Public Class CGetRealTimeX2Lot
    Private m_mySqlConn As MySqlConnection
    Private m_clsMySQL As New CMySQL

    Public Sub New(ByVal myConn As MySqlConnection)
        m_mySqlConn = myConn
    End Sub

    Public Function GetX2LotData(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal bGetCF As Boolean)
        Dim dtbDefectAdd As DataTable = GetDefectColumn()
        Select Case eSearchBy
            Case enumSearchOption.eSearchByTester
                GetX2LotData = GetX2LotByTester(strProduct, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd, enumSearchOption.eSearchByTester, dtbDefectAdd, bGetCF)
            Case enumSearchOption.eSearchByLot
                GetX2LotData = GetX2LotByLot(strProduct, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd, enumSearchOption.eSearchByLot, dtbDefectAdd)
            Case enumSearchOption.eSearchBySpec
                GetX2LotData = GetX2LotBySpecByDay(strProduct, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd, enumSearchOption.eSearchBySpec, dtbDefectAdd)
            Case Else
                Return Nothing
        End Select

    End Function

    Public Function GetXLotEndLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal bGetCF As Boolean)
        Dim strSQL As String = ""
        strSQL = "SELECT "
        strSQL = strSQL & "A.EndLotTime,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Lot,"
        'strSQL = strSQL & "IF(A.Spec LIKE 'C%',"
        'strSQL = strSQL & "(SELECT LotName FROM std_standard.tabstandard_hga LEFT JOIN std_standard.tabtray USING(TrayID) LEFT JOIN std_standard.tablot USING(LotID) WHERE Hga_SN=G.Hga_SN AND TrayName=G.TrayID LIMIT 0,1),"
        'strSQL = strSQL & "NULL) STD_Lot,"
        strSQL = strSQL & "G.StandardLot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Shoe,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "G.MediaSN,"
        'strSQL = strSQL & "(SELECT D.MediaSN FROM db_" & strProduct & ".tabdetail_header D WHERE D.Test_time_bigint=DATE_FORMAT(A.Endlottime,'%Y%m%d%H%i%s') AND D.Tester=A.Tester LIMIT 0,1) MediaSN,"
        strSQL = strSQL & "A.TotalHGA 'Tester.TotalHGA',"
        strSQL = strSQL & "A.TotalPass 'Tester.TotalPass',"
        strSQL = strSQL & "B.TotalHGA 'Lot.TotalHGA',"
        strSQL = strSQL & "B.TotalPass 'Lot.TotalPass',"
        strSQL = strSQL & "B.TotalTester,"

        Dim strSearchBy As String = dtbSearch.TableName

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("Param_rttc").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            Dim bCFAdd As Boolean = dtbParam.Rows(nParam).Item("param_add")
            Dim bCFMul As Boolean = dtbParam.Rows(nParam).Item("param_mul")
            strSQL = strSQL & "A." & strParam & " '" & strDisplayName & ".Tester',"
            strSQL = strSQL & "B." & strParam & " '" & strDisplayName & ".Lot',"
            strSQL = strSQL & "A." & strParam & "-B." & strParam & " '" & strDisplayName & ".Delta',"
            'Dim strDeltaXLot As String = strParamDisplay & ".Delta"
            If bGetCF Then
                If bCFAdd Then
                    strSQL = strSQL & "C." & strParam & "/D." & strParam & " '" & strDisplayName & ".CF',"
                ElseIf bCFMul Then
                    strSQL = strSQL & "E." & strParam & "/F." & strParam & " '" & strDisplayName & ".CF',"
                End If
            End If
        Next nParam
        strSQL = Left(strSQL, Len(strSQL) - 1)
        strSQL = strSQL & " FROM db_" & strProduct & ".tabhistory_dataendlotbytester A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabhistory_dataendlot B USING(tester,lot,spec,shoe) "
        If bGetCF Then
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd C USING(tester,lot,spec,shoe) "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd_n D USING(tester,lot,spec,shoe) "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul E USING(tester,lot,spec,shoe) "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul_n F USING(tester,lot,spec,shoe) "
        End If
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header G ON G.test_time_bigint=CONVERT(DATE_FORMAT(A.EndLotTime,'%Y%m%d%H%i%s'),UNSIGNED) AND G.Tester=A.Tester AND G.Shoe=A.Shoe "
        strSQL = strSQL & "WHERE (A.EndLotTime BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & "AND ("

        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        strSQL = Left(strSQL, Len(strSQL) - 1)
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
        strSQL = strSQL & "ORDER BY A." & strSearchBy & ",A.EndLotTime;"
        GetXLotEndLot = m_clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

    Private Function GetDefectColumn() As DataTable
        Dim strSQL As String = "SELECT MCCodeID,MCDefectName FROM db_parameter_mapping.tabmcdefect "
        strSQL = strSQL & "WHERE MCDefectName='DET_Abort' "
        strSQL = strSQL & "OR MCDefectName='FAIL_MRRCHECK' "
        strSQL = strSQL & "OR MCDefectName='FAIL_MRRCHECK_WRTFLT' "
        strSQL = strSQL & "ORDER BY MCCodeID;"
        Dim clsSql As New CMySQL
        GetDefectColumn = clsSql.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

    Private Function GetX2LotByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal dtbDefectAdd As DataTable, ByVal bGetCF As Boolean) As DataTable

        Dim strSQL As String
        strSQL = "SELECT '" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_Time,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "A.Lot,"
        'strSQL = strSQL & "IF(A.Spec LIKE 'C%',"
        'strSQL = strSQL & "(SELECT LotName FROM std_standard.tabstandard_hga LEFT JOIN std_standard.tabtray USING(TrayID) LEFT JOIN std_standard.tablot USING(LotID) WHERE Hga_SN=F.Hga_SN AND TrayName=F.TrayID LIMIT 0,1),"
        'strSQL = strSQL & "NULL) STD_Lot,"
        strSQL = strSQL & "F.StandardLot,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.IsSkipAdjustX2Lot,"
        strSQL = strSQL & "A.SliderSite,"
        'strSQL = strSQL & "(SELECT sum(C.TotalHGA) FROM db_" & strProduct & ".tabmean_avg C WHERE A.Tester=C.Tester AND A.Lot=C.Lot) ""Total.Tester"","
        'strSQL = strSQL & "(SELECT sum(C.TotalPass) FROM db_" & strProduct & ".tabmean_avg C WHERE A.Tester=C.Tester AND A.Lot=C.Lot) ""Pass.Tester"","
        'strSQL = strSQL & "(SELECT sum(C.TotalPass)/sum(C.TotalHGA)*100 FROM db_" & strProduct & ".tabmean_avg C WHERE A.Tester=C.Tester AND A.Lot=C.Lot) ""Yield.Tester"","

        strSQL = strSQL & "SUM(A.TotalHGA) 'Tester.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Tester.Pass',"
        strSQL = strSQL & " sum(A.TotalPass)/sum(A.TotalHGA)*100 'Tester.Yield',"

        strSQL = strSQL & "(SELECT sum(C.TotalHGA) FROM db_" & strProduct & ".tabmean_avg C WHERE A.Spec=C.Spec AND A.Lot=C.Lot) 'Lot.Total',"
        strSQL = strSQL & "(SELECT sum(C.TotalPass) FROM db_" & strProduct & ".tabmean_avg C WHERE A.Spec=C.Spec AND A.Lot=C.Lot) 'Lot.Pass',"
        strSQL = strSQL & "(SELECT sum(C.TotalPass)/sum(C.TotalHGA)*100 FROM db_" & strProduct & ".tabmean_avg C WHERE A.Spec=C.Spec AND A.Lot=C.Lot) 'Lot.Yield',"
        For nDefect As Integer = 0 To dtbDefectAdd.Rows.Count - 1
            strSQL = strSQL & "SUM(E.Defect" & dtbDefectAdd.Rows(nDefect).Item("MCCodeID") & ") " & dtbDefectAdd.Rows(nDefect).Item("MCDefectName") & ","
        Next nDefect
        Dim strParam As String = ""
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            strParam = dtbParam.Rows(nParam).Item("Param_rttc").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("param_mul")
            strSQL = strSQL & "SUM(S1_A." & strParam & ")/SUM(S1_N." & strParam & ") '" & strDisplayName & ".S1',"
            strSQL = strSQL & "SUM(S2_A." & strParam & ")/SUM(S2_N." & strParam & ") '" & strDisplayName & ".S2',"
            strSQL = strSQL & "(SUM(C." & strParam & ")/SUM(D." & strParam & ")) '" & strDisplayName & ".Lot',"
            If strParam = "WEW" And InStr(strProduct, "_XLOT_") > 0 Then
                strSQL = strSQL & "(SELECT SUM(G.MEW6T)/SUM(H.MEW6T) FROM db_BWT_BWT.tabmean_avg G "
                strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabmean_n H "
                strSQL = strSQL & "USING(tester,Lot,Spec,Shoe) "
                strSQL = strSQL & "WHERE A.Lot=G.Lot) '" & strDisplayName & ".BWT',"
            End If
            If strParam = "CycleTime" Then
                strSQL = strSQL & "M.CycleTime 'CycleTime.Median',"
            End If
            If bGetCF = True Then
                If bParamAdd = True Then
                    strSQL = strSQL & "SUM(CFAddS1_A." & strParam & ")/SUM(CFAddS1_N." & strParam & ") '" & strDisplayName & ".CF.S1.Add',"
                    strSQL = strSQL & "SUM(CFAddS2_A." & strParam & ")/SUM(CFAddS2_N." & strParam & ") '" & strDisplayName & ".CF.S2.Add',"
                End If
                If bParamMul = True Then
                    strSQL = strSQL & "SUM(CFMulS1_A." & strParam & ")/SUM(CFMulS1_N." & strParam & ") '" & strDisplayName & ".CF.S1.Mul',"
                    strSQL = strSQL & "SUM(CFMulS2_A." & strParam & ")/SUM(CFMulS2_N." & strParam & ") '" & strDisplayName & ".CF.S2.Mul',"
                End If
            End If
        Next nParam

        strSQL = Left(strSQL, Len(strSQL) - 1)
        strSQL = strSQL & " FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_avgbylot C USING(lot) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_nbylot D USING(lot) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect E USING(tester,lot,spec,shoe) "
        If bGetCF = True Then
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd CFAddS1_A ON A.Tester=CFAddS1_A.Tester AND A.Lot=CFAddS1_A.Lot AND A.Spec=CFAddS1_A.Spec AND CFAddS1_A.Shoe='1' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd_n CFAddS1_N ON A.Tester=CFAddS1_N.Tester AND A.Lot=CFAddS1_N.Lot AND A.Spec=CFAddS1_N.Spec AND CFAddS1_N.Shoe='1' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd CFAddS2_A ON A.Tester=CFAddS2_A.Tester AND A.Lot=CFAddS2_A.Lot AND A.Spec=CFAddS2_A.Spec AND CFAddS2_A.Shoe='2' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfadd_n CFAddS2_N ON A.Tester=CFAddS2_N.Tester AND A.Lot=CFAddS2_N.Lot AND A.Spec=CFAddS2_N.Spec AND CFAddS2_N.Shoe='2' "

            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul CFMulS1_A ON A.Tester=CFMulS1_A.Tester AND A.Lot=CFMulS1_A.Lot AND A.Spec=CFMulS1_A.Spec AND CFMulS1_A.Shoe='1' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul_n CFMulS1_N ON A.Tester=CFMulS1_N.Tester AND A.Lot=CFMulS1_N.Lot AND A.Spec=CFMulS1_N.Spec AND CFMulS1_N.Shoe='1' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul CFMulS2_A ON A.Tester=CFMulS2_A.Tester AND A.Lot=CFMulS2_A.Lot AND A.Spec=CFMulS2_A.Spec AND CFMulS2_A.Shoe='2' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_cfmul_n CFMulS2_N ON A.Tester=CFMulS2_N.Tester AND A.Lot=CFMulS2_N.Lot AND A.Spec=CFMulS2_N.Spec AND CFMulS2_N.Shoe='2' "
        End If
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_avg S1_A ON A.Tester=S1_A.Tester AND A.Lot=S1_A.Lot AND A.Spec=S1_A.Spec AND S1_A.Shoe='1' "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n S1_N ON A.Tester=S1_N.Tester AND A.Lot=S1_N.Lot AND A.Spec=S1_N.Spec AND S1_N.Shoe='1' "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_avg S2_A ON A.Tester=S2_A.Tester AND A.Lot=S2_A.Lot AND A.Spec=S2_A.Spec AND S2_A.Shoe='2' "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n S2_N ON A.Tester=S2_N.Tester AND A.Lot=S2_N.Lot AND A.Spec=S2_N.Spec AND S2_N.Shoe='2' "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header F ON F.test_time_bigint=CONVERT(DATE_FORMAT(A.update_time,'%Y%m%d%H%i%s'),UNSIGNED) AND F.Tester=A.Tester AND F.Shoe=A.Shoe "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmedian M ON M.Tester=A.Tester AND M.Spec=A.Spec AND M.Lot=A.Lot "
        strSQL = strSQL & "WHERE (A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & "AND ("

        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
        For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
            If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
            Else
                strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
            End If
        Next nSliderSite
        strSQL = strSQL & "GROUP BY A.Tester,A.Lot,A.Spec "
        strSQL = strSQL & "ORDER BY A.Tester,A.Update_Time;"

        Dim dtbXLotData As DataTable = m_clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        Dim dtbAddPettern As DataTable = CheckPattern(dtbXLotData, dtbSearch)
        GetX2LotByTester = CalculateXLotTable(strProduct, dtbAddPettern, dtbParam, eSearchBy)
    End Function

    Private Function CheckPattern(ByVal dtbXLotData As DataTable, ByVal dtbTester As DataTable) As DataTable
        dtbXLotData.Columns.Add("DeltaYield", Type.GetType("System.Double"), "Tester.Yield-Lot.Yield")
        dtbXLotData.Columns.Add("ChkPattern", Type.GetType("System.UInt16"))
        dtbXLotData.Columns.Add("Bad_Tst")

        Dim nStart As Integer = dtbXLotData.Columns("Lot.Yield").Ordinal
        dtbXLotData.Columns.Item("DeltaYield").SetOrdinal(nStart + 1)
        dtbXLotData.Columns.Item("ChkPattern").SetOrdinal(nStart + 2)
        dtbXLotData.Columns.Item("Bad_Tst").SetOrdinal(nStart + 3)

        For nTester As Integer = 0 To dtbTester.Rows.Count - 1
            Dim strTester As String = dtbTester.Rows(nTester).Item("Tester")
            Dim dtrSameTester() As DataRow = dtbXLotData.Select("[Tester]='" & strTester & "' AND (Spec LIKE 'R%' OR Spec LIKE 'T%')", "Update_Time ASC")
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

    Private Function GetX2LotByLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal dtbDefectAdd As DataTable) As DataTable

        Dim strSQL As String
        strSQL = "SELECT '" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_Time,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.IsSkipAdjustX2Lot,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "SUM(A.TotalHGA) 'Lot.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Lot.Pass',"
        strSQL = strSQL & "SUM(A.TotalPass)/SUM(A.TotalHGA)*100 'Lot.Yield',"
        For nDefect As Integer = 0 To dtbDefectAdd.Rows.Count - 1
            strSQL = strSQL & "SUM(E.Defect" & dtbDefectAdd.Rows(nDefect).Item("MCCodeID") & ") " & dtbDefectAdd.Rows(nDefect).Item("MCDefectName") & ","
        Next nDefect
        Dim strParam As String = ""
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            strParam = dtbParam.Rows(nParam).Item("Param_rttc").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n H "
            strSQL = strSQL & "USING(tester,Lot,Spec,Shoe) "
            strSQL = strSQL & "WHERE G.Shoe='1' "
            strSQL = strSQL & "AND A.Lot=G.Lot AND A.Spec=G.Spec"
            strSQL = strSQL & ") """ & strDisplayName & ".S1"","

            strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabmean_n H "
            strSQL = strSQL & "USING(tester,Lot,spec,Shoe) "
            strSQL = strSQL & "WHERE G.Shoe='2' "
            strSQL = strSQL & "AND A.Lot=G.Lot AND A.Spec=G.Spec"
            strSQL = strSQL & ") '" & strDisplayName & ".S2',"
            strSQL = strSQL & "(SUM(C." & strParam & ")/SUM(D." & strParam & ")) '" & strDisplayName & ".Lot',"
            If strParam = "WEW" And InStr(strProduct, "_XLOT_") > 0 Then
                strSQL = strSQL & "(SELECT SUM(G.MEW6T)/SUM(H.MEW6T) FROM db_BWT_BWT.tabmean_avg G "
                strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabmean_n H "
                strSQL = strSQL & "USING(tester,Lot,Spec,Shoe) "
                strSQL = strSQL & "WHERE A.Lot=G.Lot) '" & strDisplayName & ".BWT',"
            End If
            If strParam = "CycleTime" Then
                strSQL = strSQL & "AVG(M.CycleTime) 'CycleTime.Median',"
            End If
        Next nParam
        strSQL = Left(strSQL, Len(strSQL) - 1)
        strSQL = strSQL & " FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_avgbylot C USING(lot) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_nbylot D USING(lot) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect E USING(tester,lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmedian M USING(tester,lot,spec) "
        strSQL = strSQL & "WHERE (A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & "AND ("

        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        strSQL = strSQL & "GROUP BY A.Lot,A.Spec "
        strSQL = strSQL & "ORDER BY A.Lot,A.Update_Time;"

        Dim dtbXLotData As DataTable = m_clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)

        GetX2LotByLot = CalculateXLotTable(strProduct, dtbXLotData, dtbParam, eSearchBy)
    End Function

    Private Function GetX2LotBySpec(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal dtbDefectAdd As DataTable) As DataTable
        Dim strSQL As String
        strSQL = "SELECT '" & strProduct & "' ProductName,"
        strSQL = strSQL & "A.Update_Time,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
        strSQL = strSQL & "A.IsSkipAdjustX2Lot,"
        strSQL = strSQL & "A.SliderSite,"
        strSQL = strSQL & "SUM(A.TotalHGA) 'Spec.Total',"
        strSQL = strSQL & "SUM(A.TotalPass) 'Spec.Pass',"
        strSQL = strSQL & "SUM(A.TotalPass)/SUM(A.TotalHGA)*100 'Spec.Yield',"
        For nDefect As Integer = 0 To dtbDefectAdd.Rows.Count - 1
            strSQL = strSQL & "SUM(E.Defect" & dtbDefectAdd.Rows(nDefect).Item("MCCodeID") & ") " & dtbDefectAdd.Rows(nDefect).Item("MCDefectName") & ","
        Next nDefect
        Dim strParam As String = ""
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            strParam = dtbParam.Rows(nParam).Item("Param_rttc").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n H "
            strSQL = strSQL & "USING(tester,Lot,Spec,Shoe) "
            strSQL = strSQL & "WHERE G.Shoe='1' "
            strSQL = strSQL & "AND A.Spec=G.Spec "
            strSQL = strSQL & "AND (G.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
            strSQL = strSQL & ") """ & strDisplayName & ".S1"","

            strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabmean_n H "
            strSQL = strSQL & "USING(tester,lot,Spec,Shoe) "
            strSQL = strSQL & "WHERE G.Shoe='2' "
            strSQL = strSQL & "AND A.Spec=G.Spec "
            strSQL = strSQL & "AND (G.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
            strSQL = strSQL & ") """ & strDisplayName & ".S2"","

            strSQL = strSQL & "SUM(A." & strParam & ")/SUM(B." & strParam & ") """ & strDisplayName & ".Spec"","
        Next nParam

        strSQL = Left(strSQL, Len(strSQL) - 1)
        strSQL = strSQL & " FROM db_" & strProduct & ".tabmean_avg A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,lot,spec,shoe) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect E USING(tester,lot,spec,shoe) "
        strSQL = strSQL & "WHERE (A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & "AND ("

        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
        For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
            If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
            Else
                strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
            End If
        Next nSliderSite
        strSQL = strSQL & "GROUP BY A.Spec "
        strSQL = strSQL & "ORDER BY A.Spec,A.Update_Time;"

        Dim dtbXLotData As DataTable = m_clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        GetX2LotBySpec = CalculateXLotTable(strProduct, dtbXLotData, dtbParam, eSearchBy)
    End Function

    Private Function GetX2LotBySpecByDay(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchBy As enumSearchOption, ByVal dtbDefectAdd As DataTable) As DataTable
        Dim strSQL As String = ""
        dtStart = Format(dtStart, "yyyy-MM-dd 07:00:00")
        dtEnd = Format(dtEnd, "yyyy-MM-dd 07:00:00")
        Dim lngDay As Long = DateDiff(DateInterval.Day, dtStart, dtEnd)
        Dim dtbX2LotByDate As New DataTable
        Dim dtStartBydate As DateTime = dtStart
        Dim dtEndBydate As DateTime = dtStart

        For nDay As Integer = 0 To lngDay - 1
            dtStartBydate = dtEndBydate
            dtEndBydate = dtStartBydate.AddDays(1)

            strSQL = "SELECT '" & strProduct & "' ProductName,"
            strSQL = strSQL & "'" & Format(dtStartBydate, "dd-MMM-yyyy HH:mm:ss") & " To " & Format(dtEndBydate, "dd-MMM-yyyy HH:mm:ss") & "' Date_time,"
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "RIGHT(A.Spec,1) HGA_Type,"
            strSQL = strSQL & "A.IsSkipAdjustX2Lot,"
            strSQL = strSQL & "A.SliderSite,"
            strSQL = strSQL & "SUM(A.TotalHGA) 'Spec.Total',"
            strSQL = strSQL & "SUM(A.TotalPass) 'Spec.Pass',"
            strSQL = strSQL & "SUM(A.TotalPass)/SUM(A.TotalHGA)*100 'Spec.Yield',"
            For nDefect As Integer = 0 To dtbDefectAdd.Rows.Count - 1
                strSQL = strSQL & "SUM(E.Defect" & dtbDefectAdd.Rows(nDefect).Item("MCCodeID") & ") " & dtbDefectAdd.Rows(nDefect).Item("MCDefectName") & ","
            Next nDefect
            Dim strParam As String = ""
            For nParam As Integer = 0 To dtbParam.Rows.Count - 1
                strParam = dtbParam.Rows(nParam).Item("Param_rttc").ToString
                Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
                strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
                strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n H "
                strSQL = strSQL & "USING(tester,Lot,Spec,Shoe) "
                strSQL = strSQL & "WHERE G.Shoe='1' "
                strSQL = strSQL & "AND A.Spec=G.Spec "
                strSQL = strSQL & "AND (G.update_time BETWEEN '" & Format(dtStartBydate, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEndBydate, "yyyy-MM-dd HH:mm:ss") & "') "
                strSQL = strSQL & ") """ & strDisplayName & ".S1"","

                strSQL = strSQL & "(SELECT (SUM(G." & strParam & ")/SUM(H." & strParam & ")) FROM db_" & strProduct & ".tabmean_avg G "
                strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabmean_n H "
                strSQL = strSQL & "USING(tester,lot,Spec,Shoe) "
                strSQL = strSQL & "WHERE G.Shoe='2' "
                strSQL = strSQL & "AND A.Spec=G.Spec "
                strSQL = strSQL & "AND (G.update_time BETWEEN '" & Format(dtStartBydate, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEndBydate, "yyyy-MM-dd HH:mm:ss") & "') "
                strSQL = strSQL & ") """ & strDisplayName & ".S2"","

                strSQL = strSQL & "SUM(A." & strParam & ")/SUM(B." & strParam & ") """ & strDisplayName & ".Spec"","
            Next nParam

            strSQL = Left(strSQL, Len(strSQL) - 1)
            strSQL = strSQL & " FROM db_" & strProduct & ".tabmean_avg A "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,lot,spec,shoe) "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect E USING(tester,lot,spec,shoe) "
            strSQL = strSQL & "WHERE (A.update_time BETWEEN '" & Format(dtStartBydate, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEndBydate, "yyyy-MM-dd HH:mm:ss") & "') "
            strSQL = strSQL & "AND ("

            Dim strSearchBy As String = dtbSearch.TableName
            For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
                If nTester <> dtbSearch.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
                End If
            Next nTester
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
            strSQL = strSQL & "GROUP BY A.Spec "
            strSQL = strSQL & "ORDER BY A.Spec,A.Update_Time;"

            Dim dtbXLotData As DataTable = m_clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
            dtbX2LotByDate.Merge(dtbXLotData)
        Next nDay
        GetX2LotBySpecByDay = CalculateXLotTable(strProduct, dtbX2LotByDate, dtbParam, eSearchBy)
    End Function

    Private Function CalculateXLotTable(ByVal strProduct As String, ByVal dtbXLot As DataTable, ByVal dtbParam As DataTable, ByVal eSeachBy As enumSearchOption) As DataTable
        If dtbXLot.Rows.Count > 0 Then
            For nParam As Integer = 0 To dtbParam.Rows.Count - 1
                Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("param_display").ToString
                Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
                Dim strParamX As String
                If eSeachBy = enumSearchOption.eSearchBySpec Then
                    strParamX = strParamDisplay & ".Spec"
                Else
                    strParamX = strParamDisplay & ".Lot"
                End If
                Dim nParamOrdinal As Integer = dtbXLot.Columns(strParamX).Ordinal
                Dim strDeltaS1 As String = strParamDisplay & ".DeltaS1"
                Dim strDeltaS2 As String = strParamDisplay & ".DeltaS2"

                dtbXLot.Columns.Add(strDeltaS1, Type.GetType("System.Double"))
                If Not dtbXLot.Columns(strParamDisplay & ".S1") Is Nothing Then
                    dtbXLot.Columns(strDeltaS1).Expression = "[" & strParamDisplay & ".S1]-[" & strParamX & "]"
                End If
                dtbXLot.Columns.Add(strDeltaS2, Type.GetType("System.Double"))
                If Not dtbXLot.Columns(strParamDisplay & ".S2") Is Nothing Then
                    dtbXLot.Columns(strDeltaS2).Expression = "[" & strParamDisplay & ".S2]-[" & strParamX & "]"
                End If
                'dtbXLot.Columns.Add(strDeltaS2, Type.GetType("System.Double"))
                dtbXLot.Columns(strDeltaS1).SetOrdinal(nParamOrdinal + 1)
                dtbXLot.Columns(strDeltaS2).SetOrdinal(nParamOrdinal + 2)
            Next nParam
        End If
        CalculateXLotTable = dtbXLot
    End Function

End Class

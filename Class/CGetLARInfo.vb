Imports MySql.Data.MySqlClient

Public Class CGetLARInfo
    Private m_MySqlConn As MySqlConnection

    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetLARInfo(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As String, ByVal dtEnd As String)
        Dim dtbLARSetting As DataTable = GetLARSetting(strProduct)
        Dim dtbLARData As DataTable = GetLotData(strProduct, dtbSearchBy, dtStart, dtEnd, dtbLARSetting)
        GetLARInfo = dtbLARData
    End Function

    Private Function GetLotData(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal dtbLARSetting As DataTable) As DataTable
        GetLotData = Nothing
        If dtbLARSetting.Rows.Count > 0 Then
            Dim dblYieldMin As Double = dtbLARSetting.Rows(0).Item("Yield_percent")
            Dim dblDETAbortMin As Double = dtbLARSetting.Rows(0).Item("DETAbort_percent")
            Dim strSQL As String = "SELECT "
            strSQL = strSQL & "MIN(C.Test_time) StartTime,"
            strSQL = strSQL & "MAX(C.Test_time) EndTime,"
            'strSQL = strSQL & "A.Update_time EndLotTime,"
            strSQL = strSQL & "A.Tester, "
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "A.Lot, "
            strSQL = strSQL & "(A.TotalHga+B.TotalHga) 'TotalHGA',"
            strSQL = strSQL & "(A.TotalPass+B.TotalPass) 'TotalPass',"
            strSQL = strSQL & "(A.TotalPass+B.TotalPass)/(A.TotalHga+B.TotalHga)*100 'Yield',"
            strSQL = strSQL & "IF((A.TotalPass+B.TotalPass)/(A.TotalHga+B.TotalHga)*100<" & dblYieldMin & ",1,0) 'Yield.Result',"
            strSQL = strSQL & "(A.Defect3+B.Defect3) 'DETAbort',"
            strSQL = strSQL & "(A.Defect3+B.Defect3)/(A.TotalHga+B.TotalHga)*100 'DETAbort(%)',"
            strSQL = strSQL & "IF((A.Defect3+B.Defect3)/(A.TotalHga+B.TotalHga)*100>" & dblDETAbortMin & ",1,0) 'DETAbort.Result',"

            For nParam As Integer = 0 To dtbLARSetting.Rows.Count - 1
                Dim drParam As DataRow = dtbLARSetting.Rows(nParam)
                Dim bIsEnable As Boolean = drParam.Item("IsEnable")
                If bIsEnable Then
                    Dim strParam As String = drParam.Item("Param_rttc")
                    Dim dblMin As Double = drParam.Item("MinIndividual")
                    Dim dblMax As Double = drParam.Item("MaxIndividual")

                    Dim dblMeanMin As Double = drParam.Item("MeanMin")
                    Dim dblMeanMax As Double = drParam.Item("MeanMax")
                    Dim dblSigmaMax As Double = drParam.Item("Sigma")
                    Dim dblOutlierPercent As Double = drParam.Item("FailIndividual_percent")
                    strSQL = strSQL & "AVG(D." & strParam & ") '" & strParam & ".Avg',"
                    strSQL = strSQL & "STD(D." & strParam & ") '" & strParam & ".Sigma',"
                    strSQL = strSQL & "SUM(D." & strParam & "<" & dblMin & " Or D." & strParam & ">" & dblMax & ") '" & strParam & ".Outlier',"
                    strSQL = strSQL & "SUM(D." & strParam & "<" & dblMin & " Or D." & strParam & ">" & dblMax & ")/ (A.TotalHga+B.TotalHga) '" & strParam & ".Outlier(%)',"
                    strSQL = strSQL & "IF("
                    strSQL = strSQL & "AVG(D." & strParam & ")<" & dblMeanMin & " OR AVG(D." & strParam & ")>" & dblMeanMax & " OR "
                    strSQL = strSQL & "STD(D." & strParam & ")>" & dblSigmaMax & " OR "
                    strSQL = strSQL & "SUM(D." & strParam & "<" & dblMin & " Or D." & strParam & ">" & dblMax & ")/ (A.TotalHga+B.TotalHga)>" & dblOutlierPercent
                    strSQL = strSQL & ",1,0) '" & strParam & ".Result',"
                End If
            Next nParam
            If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabsummary_hgadefect A "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabsummary_hgadefect B ON A.tester=B.tester AND A.Spec=B.Spec AND A.lot=B.Lot AND B.Shoe='2' "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header C ON A.tester=C.tester AND A.Spec=C.Spec AND A.lot=C.Lot "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value D ON C.tag_id=D.tag_id "
            strSQL = strSQL & "WHERE A.Update_time BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "' AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "' "
            strSQL = strSQL & "AND A.Spec LIKE 'F%' AND A.Lot LIKE '%Q' "
            strSQL = strSQL & "AND ("
            Dim strSearchBY As String = dtbSearchBy.TableName
            For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nSearch <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "') "
                End If
            Next nSearch
            strSQL = strSQL & "AND A.Spec LIKE 'F%' AND A.Lot LIKE '%Q' "
            strSQL = strSQL & "AND A.Shoe='1' "
            strSQL = strSQL & "GROUP BY A.Tester,A.Spec,A.Lot "
            strSQL = strSQL & "ORDER BY A.Tester,A.Lot;"

            'strSQL = strSQL & "FROM " 
            'strSQL = strSQL & "(SELECT Tester,Spec,Lot FROM db_" & strProduct & ".tabmean_avg A "
            'strSQL = strSQL & "WHERE update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
            'strSQL = strSQL & "AND Spec LIKE 'F%' AND Lot LIKE '%Q' "
            'strSQL = strSQL & "AND ("
            'strSQL = strSQL & "GROUP BY Tester,Spec,Lot "
            'strSQL = strSQL & "ORDER BY A.Tester,A.Lot,A.test_time_bigint;"

            'For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
            '    If nSearch <> dtbSearchBy.Rows.Count - 1 Then
            '        strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "' OR "
            '    Else
            '        strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "') "
            '    End If
            'Next nSearch
            'strSQL = strSQL & "GROUP BY Tester,Spec,Lot) AllLot "
            'strSQL = strSQL & "INNER JOIN db_" & strProduct & ".tabdetail_header A ON A.Tester=AllLot.Tester AND A.Spec=AllLot.Spec AND A.Lot=AllLot.Lot "
            'strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
            'strSQL = strSQL & "GROUP BY A.Tester,A.Spec,A.Lot "
            'strSQL = strSQL & "ORDER BY A.Tester,A.test_time_bigint;"

            'Dim nMinYield As Double = dtbLARSetting.Rows(
            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)

            Dim strLARExpress As String = "IIF("
            For nCol As Integer = 0 To dtbData.Columns.Count - 1
                Dim strCol As String = dtbData.Columns(nCol).ColumnName
                If strCol.Contains(".Result") Then
                    strLARExpress = strLARExpress & strCol & "+"
                End If
            Next nCol
            If Right(strLARExpress, 1) = "+" Then strLARExpress = Left(strLARExpress, strLARExpress.Length - 1) & ">0,1,0)"
            Dim dcLAR As DataColumn = dtbData.Columns.Add("LAR", GetType(Int32), strLARExpress)
            dcLAR.SetOrdinal(dtbData.Columns("Lot").Ordinal + 1)
            GetLotData = dtbData
        End If
    End Function

    'Private Function GetLotData(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal dtbLARSetting As DataTable) As DataTable
    '    Dim strSQL As String = "SELECT "
    '    strSQL = strSQL & "MAX(A.Test_time) EndLotTime,"
    '    strSQL = strSQL & "A.Tester, "
    '    strSQL = strSQL & "A.Spec,"
    '    strSQL = strSQL & "A.Lot, "

    '    For nParam As Integer = 0 To dtbLARSetting.Rows.Count - 1
    '        Dim strParam As String = dtbLARSetting.Rows(nParam).Item("Param_rttc")
    '        Dim dblMin As Double = dtbLARSetting.Rows(nParam).Item("ValueMin")
    '        Dim dblMax As Double = dtbLARSetting.Rows(nParam).Item("ValueMax")
    '        strSQL = strSQL & "SUM(" & strParam & "<" & dblMin & " Or " & strParam & ">" & dblMax & ") '" & strParam & ".Outlier',"
    '        strSQL = strSQL & "AVG(" & strParam & ") '" & strParam & ".Avg',"
    '        strSQL = strSQL & "STD(" & strParam & ") '" & strParam & ".Sigma',"
    '    Next nParam
    '    strSQL = strSQL & "(SELECT SUM(TotalHga) FROM db_" & strProduct & ".tabsummary_hgadefect "
    '    strSQL = strSQL & "WHERE Tester=A.Tester AND lot=A.lot AND Spec=A.Spec) TotalHga,"
    '    strSQL = strSQL & "(SELECT SUM(TotalPass) FROM db_" & strProduct & ".tabsummary_hgadefect "
    '    strSQL = strSQL & "WHERE Tester=A.Tester AND lot=A.lot AND Spec=A.Spec) TotalPass,"
    '    strSQL = strSQL & "(SELECT SUM(Defect3) FROM db_" & strProduct & ".tabsummary_hgadefect "
    '    strSQL = strSQL & "WHERE Tester=A.Tester AND lot=A.lot AND Spec=A.Spec) DETAbort "
    '    strSQL = strSQL & "FROM "
    '    strSQL = strSQL & "(SELECT Tester,Spec,Lot FROM db_" & strProduct & ".tabmean_avg A "
    '    strSQL = strSQL & "WHERE update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
    '    strSQL = strSQL & "AND Spec LIKE 'F%' AND Lot LIKE '%Q' "
    '    strSQL = strSQL & "AND ("
    '    Dim strSearchBY As String = dtbSearchBy.TableName
    '    For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
    '        If nSearch <> dtbSearchBy.Rows.Count - 1 Then
    '            strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "' OR "
    '        Else
    '            strSQL = strSQL & "A." & strSearchBY & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBY) & "') "
    '        End If
    '    Next nSearch
    '    strSQL = strSQL & "GROUP BY Tester,Spec,Lot) AllLot "
    '    strSQL = strSQL & "INNER JOIN db_" & strProduct & ".tabdetail_header A ON A.Tester=AllLot.Tester AND A.Spec=AllLot.Spec AND A.Lot=AllLot.Lot "
    '    strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
    '    strSQL = strSQL & "GROUP BY A.Tester,A.Spec,A.Lot "
    '    strSQL = strSQL & "ORDER BY A.Tester,A.Lot;"
    '    Dim clsMySql As New CMySQL
    '    Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    '    GetLotData = dtbData
    'End Function


    Private Function GetLARSetting(ByVal strProduct As String) As DataTable
        Dim strSQL As String
        strSQL = "SELECT * "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabctr_larcontrolsetting;"
        Dim clsMySql As New CMySQL
        Dim dtbSetting As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
        GetLARSetting = dtbSetting
    End Function
End Class

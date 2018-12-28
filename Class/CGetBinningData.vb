
Imports MySql.Data.MySqlClient

Public Class CGetBinningData
    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

    Public Function GetBinningData(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal nSearchOption As enumSearchOption) As DataTable
        Dim dtbBinningSetting As DataTable = GetBinningSetting(strProduct)
        If dtbBinningSetting.Rows.Count = 0 Then Return Nothing
        Dim clsParam As New CParameterRTTCMapping(m_mySqlConn)
        Dim dtbParam As DataTable = clsParam.GetParamByProduct(strProduct, True)
        Select Case nSearchOption
            Case enumSearchOption.eSearchByTester
                GetBinningData = GetBinningDataByTester(strProduct, dtbBinningSetting, dtbParam, dtbSearchBy, dtStart, dtEnd)
            Case enumSearchOption.eSearchBySpec
                GetBinningData = GetBinningDataBySpec(strProduct, dtbBinningSetting, dtbParam, dtbSearchBy, dtStart, dtEnd)
            Case enumSearchOption.eSearchByLot
                GetBinningData = GetBinningDataByLot(strProduct, dtbBinningSetting, dtbParam, dtbSearchBy, dtStart, dtEnd)
            Case enumSearchOption.eSearchByMachineType
                GetBinningData = GetBinningDataByMachineType(strProduct, dtbBinningSetting, dtbParam, dtbSearchBy, dtStart, dtEnd)
            Case Else
                GetBinningData = Nothing
        End Select

    End Function

    Private Function GetBinningSetting(ByVal strProduct As String) As DataTable
        Dim strSQL As String = "SELECT * FROM db_" & strProduct & ".tabctr_databinning;"
        Dim clsMySQL As New CMySQL
        Dim dtbSetting As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        GetBinningSetting = dtbSetting
    End Function

    Private Function GetBinningDataByTester(ByVal strProduct As String, ByVal dtbBinningSetting As DataTable, ByVal dtbParam As DataTable, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim strBinningParam As String = dtbBinningSetting.Rows(0).Item("BinningParam")
        Dim dblBinningTarget As Double = dtbBinningSetting.Rows(0).Item("BinningTarget")
        Dim strBinningSpec As String = dtbBinningSetting.Rows(0).Item("BinningSpec")

        dtStart = Format(dtStart, "yyyy-MM-dd 07:00:00")
        dtEnd = Format(dtEnd, "yyyy-MM-dd 07:00:00")
        Dim lngDay As Long = DateDiff(DateInterval.Day, dtStart, dtEnd)
        Dim dtbX2LotByDate As New DataTable
        Dim dtStartBydate As DateTime = dtStart
        Dim dtEndBydate As DateTime = dtStart

        Dim dtbBinningDataByDate As New DataTable
        For nDay As Integer = 0 To lngDay - 1
            dtStartBydate = dtEndBydate
            dtEndBydate = dtStartBydate.AddDays(1)
            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "'" & Format(dtStartBydate, "dd-MMM-yyyy") & "' Date_time," '& " To " & Format(dtEndBydate, "dd-MMM-yyyy HH:mm:ss") & "' Date_time,"
            strSQL = strSQL & "'" & strProduct & "' Product,"
            strSQL = strSQL & "A.Tester,"
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "COUNT(tag_id) Total,"
            strSQL = strSQL & "AVG(B." & strBinningParam & ") '" & strBinningParam & "',"
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam)
                If strCorrelateParam <> "" Then
                    Dim dtrParam() As DataRow = dtbParam.Select("Param_rttc='" & strCorrelateParam & "'")
                    Dim strParamDisplay As String = dtrParam(0).Item("Param_display")
                    strSQL = strSQL & "AVG(B." & strCorrelateParam & ") '" & strParamDisplay & "',"
                End If
            Next nParam
            If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
            strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStartBydate, "yyyyMMddHHmmss") & "' and '" & Format(dtEndBydate, "yyyyMMddHHmmss") & "') "
            strSQL = strSQL & " AND ("
            Dim strSearchBy As String = dtbSearchBy.TableName
            For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nSearch <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "') "
                End If
            Next nSearch
            strSQL = strSQL & "AND LEFT(A.Spec,1)='" & strBinningSpec & "' "
            Dim dblBinningMin As Double = dtbBinningSetting.Rows(0).Item("BinningMin")
            Dim dblBinningMax As Double = dtbBinningSetting.Rows(0).Item("BinningMax")
            strSQL = strSQL & "AND " & strBinningParam & ">=" & dblBinningMin & " AND " & strBinningParam & "<=" & dblBinningMax & " "
            strSQL = strSQL & "AND " & strBinningParam & " IS NOT NULL "
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam).ToString
                If strCorrelateParam <> "" Then
                    Dim dblCorrelateMin As Double = dtbBinningSetting.Rows(0).Item("ParaMin" & nParam)
                    Dim dblCorrelateMax As Double = dtbBinningSetting.Rows(0).Item("ParaMax" & nParam)
                    strSQL = strSQL & "AND ((" & strCorrelateParam & ">=" & dblCorrelateMin & " AND " & strCorrelateParam & "<=" & dblCorrelateMax & ") "
                    strSQL = strSQL & "OR " & strCorrelateParam & " IS NULL) "
                End If
            Next nParam
            For nParam As Integer = 0 To 9
                Dim strFilterParam As String = dtbBinningSetting.Rows(0).Item("FilterPara" & nParam).ToString
                If strFilterParam <> "" Then
                    Dim dblFilterMin As Double = dtbBinningSetting.Rows(0).Item("FilterMin" & nParam)
                    Dim dblFilterMax As Double = dtbBinningSetting.Rows(0).Item("FilterMax" & nParam)
                    strSQL = strSQL & "AND ((" & strFilterParam & ">=" & dblFilterMin & " AND " & strFilterParam & "<=" & dblFilterMax & ") "
                    strSQL = strSQL & "OR " & strFilterParam & " IS NULL) "
                End If
            Next nParam
            strSQL = strSQL & "GROUP BY A.Tester,A.Spec "
            strSQL = strSQL & "ORDER BY A.Tester,A.Spec;"
            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
            dtbBinningDataByDate.Merge(dtbData, dblBinningTarget)
        Next nDay

        GetBinningDataByTester = dtbBinningDataByDate
    End Function

    Private Function GetBinningDataBySpec(ByVal strProduct As String, ByVal dtbBinningSetting As DataTable, ByVal dtbParam As DataTable, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim strBinningParam As String = dtbBinningSetting.Rows(0).Item("BinningParam")
        Dim dblBinningTarget As Double = dtbBinningSetting.Rows(0).Item("BinningTarget")
        Dim strBinningSpec As String = dtbBinningSetting.Rows(0).Item("BinningSpec")

        dtStart = Format(dtStart, "yyyy-MM-dd 07:00:00")
        dtEnd = Format(dtEnd, "yyyy-MM-dd 07:00:00")
        Dim lngDay As Long = DateDiff(DateInterval.Day, dtStart, dtEnd)
        Dim dtbX2LotByDate As New DataTable
        Dim dtStartBydate As DateTime = dtStart
        Dim dtEndBydate As DateTime = dtStart

        Dim dtbBinningDataByDate As New DataTable
        For nDay As Integer = 0 To lngDay - 1
            dtStartBydate = dtEndBydate
            dtEndBydate = dtStartBydate.AddDays(1)

            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "'" & Format(dtStartBydate, "dd-MMM-yyyy") & "' Date_time," '& " To " & Format(dtEndBydate, "dd-MMM-yyyy HH:mm:ss") & "' Date_time,"
            strSQL = strSQL & "'" & strProduct & "' Product,"
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "COUNT(tag_id) Total,"
            strSQL = strSQL & "AVG(B." & strBinningParam & ") '" & strBinningParam & "',"
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam)
                If strCorrelateParam <> "" Then
                    Dim dtrParam() As DataRow = dtbParam.Select("Param_rttc='" & strCorrelateParam & "'")
                    Dim strParamDisplay As String = dtrParam(0).Item("Param_display")
                    strSQL = strSQL & "AVG(B." & strCorrelateParam & ") '" & strParamDisplay & "',"
                End If
            Next nParam
            If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
            strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStartBydate, "yyyyMMddHHmmss") & "' and '" & Format(dtEndBydate, "yyyyMMddHHmmss") & "') "
            strSQL = strSQL & " AND ("
            Dim strSearchBy As String = dtbSearchBy.TableName
            For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nSearch <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "') "
                End If
            Next nSearch
            strSQL = strSQL & "AND LEFT(A.Spec,1)='" & strBinningSpec & "' "
            Dim dblBinningMin As Double = dtbBinningSetting.Rows(0).Item("BinningMin")
            Dim dblBinningMax As Double = dtbBinningSetting.Rows(0).Item("BinningMax")
            strSQL = strSQL & "AND " & strBinningParam & ">=" & dblBinningMin & " AND " & strBinningParam & "<=" & dblBinningMax & " "
            strSQL = strSQL & "AND " & strBinningParam & " IS NOT NULL "
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam).ToString
                If strCorrelateParam <> "" Then
                    Dim dblCorrelateMin As Double = dtbBinningSetting.Rows(0).Item("ParaMin" & nParam)
                    Dim dblCorrelateMax As Double = dtbBinningSetting.Rows(0).Item("ParaMax" & nParam)
                    strSQL = strSQL & "AND ((" & strCorrelateParam & ">=" & dblCorrelateMin & " AND " & strCorrelateParam & "<=" & dblCorrelateMax & ") "
                    strSQL = strSQL & "OR " & strCorrelateParam & " IS NULL) "
                End If
            Next nParam
            For nParam As Integer = 0 To 9
                Dim strFilterParam As String = dtbBinningSetting.Rows(0).Item("FilterPara" & nParam).ToString
                If strFilterParam <> "" Then
                    Dim dblFilterMin As Double = dtbBinningSetting.Rows(0).Item("FilterMin" & nParam)
                    Dim dblFilterMax As Double = dtbBinningSetting.Rows(0).Item("FilterMax" & nParam)
                    strSQL = strSQL & "AND ((" & strFilterParam & ">=" & dblFilterMin & " AND " & strFilterParam & "<=" & dblFilterMax & ") "
                    strSQL = strSQL & "OR " & strFilterParam & " IS NULL) "
                End If
            Next nParam
            strSQL = strSQL & "GROUP BY A.Spec "
            strSQL = strSQL & "ORDER BY A.Spec;"
            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
            dtbBinningDataByDate.Merge(EditBinningData(dtbData, dblBinningTarget))
        Next nDay

        GetBinningDataBySpec = dtbBinningDataByDate
    End Function

    Private Function GetBinningDataByLot(ByVal strProduct As String, ByVal dtbBinningSetting As DataTable, ByVal dtbParam As DataTable, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim strBinningParam As String = dtbBinningSetting.Rows(0).Item("BinningParam")
        Dim dblBinningTarget As Double = dtbBinningSetting.Rows(0).Item("BinningTarget")
        Dim strBinningSpec As String = dtbBinningSetting.Rows(0).Item("BinningSpec")

        dtStart = Format(dtStart, "yyyy-MM-dd 07:00:00")
        dtEnd = Format(dtEnd, "yyyy-MM-dd 07:00:00")
        Dim lngDay As Long = DateDiff(DateInterval.Day, dtStart, dtEnd)
        Dim dtbX2LotByDate As New DataTable
        Dim dtStartBydate As DateTime = dtStart
        Dim dtEndBydate As DateTime = dtStart

        Dim dtbBinningDataByDate As New DataTable
        For nDay As Integer = 0 To lngDay - 1
            dtStartBydate = dtEndBydate
            dtEndBydate = dtStartBydate.AddDays(1)

            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "'" & Format(dtStartBydate, "dd-MMM-yyyy") & "' Date_time," '& " To " & Format(dtEndBydate, "dd-MMM-yyyy HH:mm:ss") & "' Date_time,"
            strSQL = strSQL & "'" & strProduct & "' Product,"
            strSQL = strSQL & "LEFT(A.Lot,4) Wafer,"
            strSQL = strSQL & "A.Spec,"
            strSQL = strSQL & "COUNT(tag_id) Total,"
            strSQL = strSQL & "AVG(B." & strBinningParam & ") '" & strBinningParam & "',"
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam)
                If strCorrelateParam <> "" Then
                    Dim dtrParam() As DataRow = dtbParam.Select("Param_rttc='" & strCorrelateParam & "'")
                    Dim strParamDisplay As String = dtrParam(0).Item("Param_display")
                    strSQL = strSQL & "AVG(B." & strCorrelateParam & ") '" & strParamDisplay & "',"
                End If
            Next nParam
            If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
            strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStartBydate, "yyyyMMddHHmmss") & "' and '" & Format(dtEndBydate, "yyyyMMddHHmmss") & "') "
            strSQL = strSQL & " AND ("
            Dim strSearchBy As String = dtbSearchBy.TableName
            For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nSearch <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "') "
                End If
            Next nSearch
            strSQL = strSQL & "AND LEFT(A.Spec,1)='" & strBinningSpec & "' "
            Dim dblBinningMin As Double = dtbBinningSetting.Rows(0).Item("BinningMin")
            Dim dblBinningMax As Double = dtbBinningSetting.Rows(0).Item("BinningMax")
            strSQL = strSQL & "AND " & strBinningParam & ">=" & dblBinningMin & " AND " & strBinningParam & "<=" & dblBinningMax & " "
            strSQL = strSQL & "AND " & strBinningParam & " IS NOT NULL "
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam).ToString
                If strCorrelateParam <> "" Then
                    Dim dblCorrelateMin As Double = dtbBinningSetting.Rows(0).Item("ParaMin" & nParam)
                    Dim dblCorrelateMax As Double = dtbBinningSetting.Rows(0).Item("ParaMax" & nParam)
                    strSQL = strSQL & "AND ((" & strCorrelateParam & ">=" & dblCorrelateMin & " AND " & strCorrelateParam & "<=" & dblCorrelateMax & ") "
                    strSQL = strSQL & "OR " & strCorrelateParam & " IS NULL) "
                End If
            Next nParam
            For nParam As Integer = 0 To 9
                Dim strFilterParam As String = dtbBinningSetting.Rows(0).Item("FilterPara" & nParam).ToString
                If strFilterParam <> "" Then
                    Dim dblFilterMin As Double = dtbBinningSetting.Rows(0).Item("FilterMin" & nParam)
                    Dim dblFilterMax As Double = dtbBinningSetting.Rows(0).Item("FilterMax" & nParam)
                    strSQL = strSQL & "AND ((" & strFilterParam & ">=" & dblFilterMin & " AND " & strFilterParam & "<=" & dblFilterMax & ") "
                    strSQL = strSQL & "OR " & strFilterParam & " IS NULL) "
                End If
            Next nParam
            strSQL = strSQL & "GROUP BY LEFT(A.Lot,4),A.Spec "
            strSQL = strSQL & "ORDER BY A.Lot,A.Spec;"
            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
            dtbBinningDataByDate.Merge(dtbData, dblBinningTarget)
        Next nDay

        GetBinningDataByLot = dtbBinningDataByDate
    End Function

    Private Function GetBinningDataByMachineType(ByVal strProduct As String, ByVal dtbBinningSetting As DataTable, ByVal dtbParam As DataTable, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable

        Dim strBinningParam As String = dtbBinningSetting.Rows(0).Item("BinningParam")
        Dim dblBinningTarget As Double = dtbBinningSetting.Rows(0).Item("BinningTarget")
        Dim strBinningSpec As String = dtbBinningSetting.Rows(0).Item("BinningSpec")

        dtStart = Format(dtStart, "yyyy-MM-dd 07:00:00")
        dtEnd = Format(dtEnd, "yyyy-MM-dd 07:00:00")
        Dim lngDay As Long = DateDiff(DateInterval.Day, dtStart, dtEnd)
        Dim dtbX2LotByDate As New DataTable
        Dim dtStartBydate As DateTime = dtStart
        Dim dtEndBydate As DateTime = dtStart

        Dim dtbBinningDataByDate As New DataTable
        For nDay As Integer = 0 To lngDay - 1
            dtStartBydate = dtEndBydate
            dtEndBydate = dtStartBydate.AddDays(1)

            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "'" & Format(dtStartBydate, "dd-MMM-yyyy") & "' Date_time," '& " To " & Format(dtEndBydate, "dd-MMM-yyyy HH:mm:ss") & "' Date_time,"
            strSQL = strSQL & "'" & strProduct & "' Product,"
            strSQL = strSQL & "IF(RIGHT(A.Spec,1)='A','TypeUp','TypeDown') MachineType,"
            strSQL = strSQL & "COUNT(tag_id) Total,"
            strSQL = strSQL & "AVG(B." & strBinningParam & ") '" & strBinningParam & "',"
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam)
                If strCorrelateParam <> "" Then
                    Dim dtrParam() As DataRow = dtbParam.Select("Param_rttc='" & strCorrelateParam & "'")
                    Dim strParamDisplay As String = dtrParam(0).Item("Param_display")
                    strSQL = strSQL & "AVG(B." & strCorrelateParam & ") '" & strParamDisplay & "',"
                End If
            Next nParam
            If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
            strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStartBydate, "yyyyMMddHHmmss") & "' and '" & Format(dtEndBydate, "yyyyMMddHHmmss") & "') "
            strSQL = strSQL & " AND ("
            Dim strSearchBy As String = dtbSearchBy.TableName
            For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If dtbSearchBy.Rows(nSearch).Item("OptionIndex") = enumMachineType.eTypeUp Then
                    strSQL = strSQL & "A.Spec LIKE '%A' OR "
                Else
                    strSQL = strSQL & "A.Spec LIKE '%B' OR "
                End If
            Next nSearch
            If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 4) & ") "
            strSQL = strSQL & "AND A.Spec LIKE '" & strBinningSpec & "%' "
            Dim dblBinningMin As Double = dtbBinningSetting.Rows(0).Item("BinningMin")
            Dim dblBinningMax As Double = dtbBinningSetting.Rows(0).Item("BinningMax")
            strSQL = strSQL & "AND " & strBinningParam & ">=" & dblBinningMin & " AND " & strBinningParam & "<=" & dblBinningMax & " "
            strSQL = strSQL & "AND " & strBinningParam & " IS NOT NULL "
            For nParam As Integer = 0 To 4
                Dim strCorrelateParam As String = dtbBinningSetting.Rows(0).Item("Para" & nParam).ToString
                If strCorrelateParam <> "" Then
                    Dim dblCorrelateMin As Double = dtbBinningSetting.Rows(0).Item("ParaMin" & nParam)
                    Dim dblCorrelateMax As Double = dtbBinningSetting.Rows(0).Item("ParaMax" & nParam)
                    strSQL = strSQL & "AND ((" & strCorrelateParam & ">=" & dblCorrelateMin & " AND " & strCorrelateParam & "<=" & dblCorrelateMax & ") "
                    strSQL = strSQL & "OR " & strCorrelateParam & " IS NULL) "
                End If
            Next nParam
            For nParam As Integer = 0 To 9
                Dim strFilterParam As String = dtbBinningSetting.Rows(0).Item("FilterPara" & nParam).ToString
                If strFilterParam <> "" Then
                    Dim dblFilterMin As Double = dtbBinningSetting.Rows(0).Item("FilterMin" & nParam)
                    Dim dblFilterMax As Double = dtbBinningSetting.Rows(0).Item("FilterMax" & nParam)
                    strSQL = strSQL & "AND ((" & strFilterParam & ">=" & dblFilterMin & " AND " & strFilterParam & "<=" & dblFilterMax & ") "
                    strSQL = strSQL & "OR " & strFilterParam & " IS NULL) "
                End If
            Next nParam
            strSQL = strSQL & "GROUP BY RIGHT(A.Spec,1) "
            strSQL = strSQL & "ORDER BY A.Spec;"
            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
            dtbBinningDataByDate.Merge(dtbData)
        Next nDay

        GetBinningDataByMachineType = dtbBinningDataByDate
    End Function


    Private Function EditBinningData(ByVal dtbBinning As DataTable, ByVal strBinningTarget As String) As DataTable

        Dim nSpec As Integer = dtbBinning.Columns("Spec").Ordinal
        Dim strCondition As String = dtbBinning.Columns(nSpec - 1).ColumnName
        Dim dtrTypeA() As DataRow = dtbBinning.Select("Spec LIKE '*A'")
        Dim dtrTypeB() As DataRow = dtbBinning.Select("Spec LIKE '*B'")
        Dim dtbData As New DataTable
        For nCol As Integer = 0 To dtbBinning.Columns.Count - 1
            Dim strColName As String = dtbBinning.Columns(nCol).ColumnName
            If nCol <= nSpec Then
                dtbData.Columns.Add(strColName, dtbBinning.Columns(nCol).DataType)
            Else
                dtbData.Columns.Add(strColName & "_Up", dtbBinning.Columns(nCol).DataType)
                dtbData.Columns.Add(strColName & "_Dn", dtbBinning.Columns(nCol).DataType)
            End If
        Next nCol
        Dim dtrPrimary() As DataRow
        If dtrTypeA.Length > dtrTypeB.Length Then
            dtrPrimary = dtrTypeA
        Else
            dtrPrimary = dtrTypeB
        End If

        For nData As Integer = 0 To dtrPrimary.Length - 1
            Dim strSpec As String = Left(dtrPrimary(nData).Item("Spec"), 3)
            Dim strFilter As String = "[" & dtbBinning.Columns(nSpec - 1).ColumnName & "]='" & dtrPrimary(nData).Item(nSpec - 1) & "'"
            dtbData.Rows.Add()
            Dim dtrDataA() As DataRow = dtbBinning.Select(strFilter & " AND [Spec]='" & strSpec & "A'")
            Dim dtrDataB() As DataRow = dtbBinning.Select(strFilter & " AND [Spec]='" & strSpec & "B'")
            For nCol As Integer = 0 To dtbBinning.Columns.Count - 1
                Dim strColName As String = dtbBinning.Columns(nCol).ColumnName
                If nCol < nSpec Then
                    If dtrDataA.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrDataA(0).Item(strColName)
                    If dtrDataB.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrDataB(0).Item(strColName)
                ElseIf nCol = nSpec Then
                    dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = strSpec & "_" & strBinningTarget
                Else
                    If dtrDataA.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName & "_Up") = dtrDataA(0).Item(strColName)
                    If dtrDataB.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName & "_Dn") = dtrDataB(0).Item(strColName)
                End If
            Next nCol
        Next nData
        EditBinningData = dtbData
    End Function

End Class

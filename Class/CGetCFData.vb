Imports MySql.Data.MySqlClient

Public Class CGetCFData
    Private m_MySqlConn As MySqlConnection
    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetCFByEndOfHour(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal bShowDeltaGOS As Boolean, Optional ByVal strShoe As String = "") As DataTable

        If dtEnd > Now Then dtEnd = Now
        Dim strSQL As String = ""
        Dim strSearchBy As String = dtbSearch.TableName
        Dim bDivideShoe As Boolean = True
        If strSearchBy.ToUpper = "SPEC" Then
            bDivideShoe = False
        End If
        If dtEnd > Now Then dtEnd = Now

        strSQL = "SELECT "
        strSQL = strSQL & "A.test_time_bigint MaxTime,"
        strSQL = strSQL & "CAST(DATE_FORMAT(DATE_ADD(A.Test_time,INTERVAL 1 HOUR),'%Y-%m-%d %H:00:00') AS DATETIME) Date_Time,"
        strSQL = strSQL & "A." & strSearchBy & ","
        strSQL = strSQL & "IF(RIGHT(A.Spec,1)='A','Up','Down') MachineType,"
        If bShowDeltaGOS Then
            strSQL = strSQL & "0 CntTester,"
            strSQL = strSQL & "0 CntGOS,"
        End If
        If bDivideShoe Then
            strSQL = strSQL & "A.Shoe,"
        End If
        strSQL = strSQL & "A.MediaSN,"

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("param_Display").ToString
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
            If bParamAdd = True Then
                strSQL = strSQL & "B." & strParam & " '" & strParamDisplay & ".CFAdd',"
            End If
            If bParamMul = True Then
                strSQL = strSQL & "C." & strParam & " '" & strParamDisplay & ".CFMul',"
            End If
            If bParamAdd Or bParamMul Then
                strSQL = strSQL & "M." & strParam & " '" & strParamDisplay & ".CFMedia',"
                If bShowDeltaGOS Then
                    strSQL = strSQL & "(0.0 + '0') '" & strParamDisplay & ".DeltaGOS',"       'Use this to return type double 
                    strSQL = strSQL & "(0.0 + '0') '" & strParamDisplay & ".SigmaGOS',"     'Use this to return type double 
                End If
            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM "
        strSQL = strSQL & "(SELECT MAX(Test_time_bigint) AS Max_time,Tester,Shoe FROM db_" & strProduct & ".tabdetail_header "
        strSQL = strSQL & "WHERE (test_time_bigint between '" & Format(dtStart, "yyyyMMddHH0000") & "' AND '" & Format(dtEnd, "yyyyMMddHH5959") & "') "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "" & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "" & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        If bDivideShoe Then
            If strShoe <> "" Then
                strSQL = strSQL & "AND Shoe='" & strShoe & "' "
            End If
        End If
        strSQL = strSQL & "AND GradeName NOT LIKE 'REJECT LOW%' "
        strSQL = strSQL & "AND GradeName NOT LIKE 'FAIL_NO_READING%' "
        strSQL = strSQL & "AND GradeName NOT LIKE 'FAIL_MRRCHECK%' "
        strSQL = strSQL & "AND GradeName NOT LIKE 'FAIL_CALL%' "
        strSQL = strSQL & "AND GradeName NOT LIKE 'FAIL-CALL%' "
        strSQL = strSQL & "AND GradeName<>'' "
        If strProduct.Contains("V2002") > 0 Or strProduct.Contains("DCT_SDET") Then
            strSQL = strSQL & "AND Shoe='1' "
        End If
        strSQL = strSQL & "GROUP BY DATE_FORMAT(Test_time,'%Y%m%d%k')," & strSearchBy & " "
        If bDivideShoe Then
            strSQL = strSQL & ",Shoe "
        End If
        strSQL = strSQL & ") T "
        strSQL = strSQL & "INNER JOIN db_" & strProduct & ".tabdetail_header A ON A.test_time_bigint=T.Max_time AND A.Tester=T.Tester AND A.Shoe=T.Shoe "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd B USING(tag_id) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfmul C USING(tag_id) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_media M ON DiskSN=LEFT(A.MediaSN,LENGTH(A.MediaSN)-3) "
        strSQL = strSQL & "ORDER BY A." & strSearchBy & ",A.Shoe,A.Test_time_bigint;"
        Dim clsMySQL As New CMySQL
        Dim dtbData As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MySqlConn)
        If dtbData.Rows.Count = 0 Then Return Nothing

        Dim dtbDataFillDate As DataTable = FillBlankDateTime(dtbData, dtStart, dtEnd, dtbSearch, bDivideShoe)
        If bDivideShoe Then
            GetCFByEndOfHour = CombineTableAllShoe(dtbDataFillDate, dtStart, dtEnd)
            GetCFByEndOfHour.Columns.Remove("MaxTime")
        Else
            GetCFByEndOfHour = dtbDataFillDate
            GetCFByEndOfHour.Columns.Remove("MaxTime")
        End If

    End Function

    'Public Function GetCFByEndOfHour(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal bShowDeltaGOS As Boolean, Optional ByVal strShoe As String = "") As DataTable

    '    Dim strSQL As String = ""
    '    Dim strSearchBy As String = dtbSearch.TableName
    '    Dim bDivideShoe As Boolean = True
    '    If strSearchBy.ToUpper = "SPEC" Then
    '        bDivideShoe = False
    '    End If
    '    If dtEnd > Now Then dtEnd = Now

    '    strSQL = "SELECT "
    '    strSQL = strSQL & "MAX(A.test_time_bigint) MaxTime,"
    '    strSQL = strSQL & "CAST(DATE_FORMAT(DATE_ADD(A.Test_time,INTERVAL 1 HOUR),'%Y-%m-%d %H:00:00') AS DATETIME) Date_Time,"
    '    strSQL = strSQL & "A." & strSearchBy & ","
    '    strSQL = strSQL & "IF(RIGHT(A.Spec,1)='A','Up','Down') MachineType,"
    '    If bShowDeltaGOS Then
    '        strSQL = strSQL & "0 CntTester,"
    '        strSQL = strSQL & "0 CntGOS,"
    '    End If
    '    If bDivideShoe Then
    '        strSQL = strSQL & "A.Shoe,"
    '    End If
    '    strSQL = strSQL & "A.MediaSN,"

    '    For nParam As Integer = 0 To dtbParam.Rows.Count - 1
    '        Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
    '        Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("param_Display").ToString
    '        Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
    '        Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
    '        If bParamAdd = True Then
    '            strSQL = strSQL & "(SELECT AVG(N." & strParam & ") FROM db_" & strProduct & ".tabdetail_header M "
    '            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd N USING(tag_id) "
    '            strSQL = strSQL & "WHERE M.test_time_bigint=MAX(A.test_time_bigint) "
    '            strSQL = strSQL & "AND M." & strSearchBy & "=A." & strSearchBy & " AND M.Shoe=A.Shoe) """ & strParamDisplay & ".CFAdd"","
    '        End If
    '        If bParamMul = True Then
    '            strSQL = strSQL & "(SELECT AVG(N." & strParam & ") FROM db_" & strProduct & ".tabdetail_header M "
    '            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfmul N USING(tag_id) "
    '            strSQL = strSQL & "WHERE M.test_time_bigint=MAX(A.test_time_bigint) "
    '            strSQL = strSQL & "AND M." & strSearchBy & "=A." & strSearchBy & " AND M.Shoe=A.Shoe) """ & strParamDisplay & ".CFMul"","
    '        End If
    '        If bParamAdd Or bParamMul Then
    '            strSQL = strSQL & "(SELECT " & strParam & " FROM db_" & strProduct & ".tabfactor_media WHERE DiskSN=LEFT(A.MediaSN,LENGTH(A.MediaSN)-3)) """ & strParamDisplay & ".CFMedia"","
    '            If bShowDeltaGOS Then
    '                strSQL = strSQL & "(0.0 + '0') """ & strParamDisplay & ".DeltaGOS"","       'Use this to return type double 
    '                strSQL = strSQL & "(0.0 + '0') """ & strParamDisplay & ".SigmaGOS"","     'Use this to return type double 
    '            End If
    '        End If
    '    Next nParam
    '    If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
    '    strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A "
    '    strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStart.AddHours(-73), "yyyyMMddHH0000") & "' AND '" & Format(dtEnd, "yyyyMMddHH5959") & "') "
    '    strSQL = strSQL & "AND ("
    '    For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
    '        If nSearch <> dtbSearch.Rows.Count - 1 Then
    '            strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
    '        Else
    '            strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
    '        End If
    '    Next nSearch
    '    If bDivideShoe Then
    '        If strShoe <> "" Then
    '            strSQL = strSQL & "AND A.Shoe='" & strShoe & "' "
    '        End If
    '    End If
    '    strSQL = strSQL & "AND A.GradeName NOT LIKE 'REJECT LOW%' "
    '    strSQL = strSQL & "AND A.GradeName NOT LIKE 'FAIL_NO_READING%' "
    '    strSQL = strSQL & "AND A.GradeName NOT LIKE 'FAIL_MRRCHECK%' "
    '    strSQL = strSQL & "AND A.GradeName NOT LIKE 'FAIL_CALL%' "
    '    strSQL = strSQL & "AND A.GradeName NOT LIKE 'FAIL-CALL%' "
    '    If InStr(strProduct, "V2002") > 0 Then
    '        strSQL = strSQL & "AND A.Shoe='1' "
    '    End If
    '    strSQL = strSQL & "GROUP BY DATE_FORMAT(A.Test_time,'%Y%m%d%k'),A." & strSearchBy & " "
    '    If bDivideShoe Then
    '        strSQL = strSQL & ",A.Shoe "
    '    End If
    '    strSQL = strSQL & "ORDER BY A." & strSearchBy & ",A.Shoe,A.Test_time_bigint;"
    '    Dim clsMySQL As New CMySQL
    '    Dim dtbData As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MySqlConn)
    '    If dtbData.Rows.Count = 0 Then Return Nothing

    '    Dim dtbDataFillDate As DataTable = FillBlankDateTime(dtbData, dtStart, dtEnd, dtbSearch, bDivideShoe)
    '    If bDivideShoe Then
    '        GetCFByEndOfHour = CombineTableAllShoe(dtbDataFillDate, dtStart, dtEnd)
    '        'GetCFByEndOfHour.Columns.Remove("MaxTime")
    '    Else
    '        GetCFByEndOfHour = dtbDataFillDate
    '        GetCFByEndOfHour.Columns.Remove("MaxTime")
    '    End If

    'End Function

    Private Function FillBlankDateTime(ByVal dtbData As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal dtbSearch As DataTable, ByVal bDivideShoe As Boolean)
        If dtEnd > Now Then dtEnd = Now
        Dim nHourDiff As Integer = DateDiff(DateInterval.Hour, dtStart, dtEnd)
        Dim dtbDataFillDate As DataTable
        dtbDataFillDate = dtbData.Clone
        Dim strSearchBy As String = dtbSearch.TableName
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            Dim strSearch As String = dtbSearch.Rows(nSearch).Item(strSearchBy)
            Dim strFilterShoe1 As String = "[" & strSearchBy & "]='" & strSearch & "'"
            Dim strFilterShoe2 As String = "[" & strSearchBy & "]='" & strSearch & "'"
            If bDivideShoe Then
                strFilterShoe1 = strFilterShoe1 & " AND [Shoe]='1'"
                strFilterShoe2 = strFilterShoe2 & " AND [Shoe]='2'"
            End If
            Dim dtrShoe1() As DataRow = dtbData.Select(strFilterShoe1, "Date_time ASC")
            Dim dtrShoe2() As DataRow = dtbData.Select(strFilterShoe2, "Date_time ASC")
            Dim dtrDefault1 As DataRow = Nothing
            Dim dtrDefault2 As DataRow = Nothing
            If dtrShoe1.Length > 0 Then dtrDefault1 = dtrShoe1(0)
            If dtrShoe2.Length > 0 Then dtrDefault2 = dtrShoe2(0)
            For nHour As Integer = 0 To nHourDiff
                Dim dtTime As DateTime = dtStart.AddHours(nHour)

                Dim strTime As String = Format(dtTime, "yyyy-MM-dd HH:00:00")
                Dim strSelectShoe1 As String = "[Date_time]='" & strTime & "' AND [" & strSearchBy & "]='" & strSearch & "'"
                Dim strSelectShoe2 As String = "[Date_time]='" & strTime & "' AND [" & strSearchBy & "]='" & strSearch & "'"
                If bDivideShoe Then
                    strSelectShoe1 = strSelectShoe1 & " AND [Shoe]='1'"
                    strSelectShoe2 = strSelectShoe2 & " AND [Shoe]='2'"
                End If
                Dim dtrSelectShoe1() As DataRow = dtbData.Select(strSelectShoe1)
                Dim dtrSelectShoe2() As DataRow = dtbData.Select(strSelectShoe2)
                Dim strMediaSN As String = ""
                If dtrSelectShoe1.Length > 0 Then
                    dtbDataFillDate.Rows.Add(dtrSelectShoe1(0).ItemArray)
                    dtrDefault1 = dtrSelectShoe1(0)
                    strMediaSN = dtrSelectShoe1(0).Item("MediaSN")
                Else
                    If Not dtrDefault1 Is Nothing Then
                        dtbDataFillDate.Rows.Add(dtrDefault1.ItemArray)
                        dtbDataFillDate.Rows(dtbDataFillDate.Rows.Count - 1).Item("Date_time") = CDate(strTime)
                        If strMediaSN = "" Then strMediaSN = dtrDefault1.Item("MediaSN")
                    End If
                End If

                If dtrSelectShoe2.Length > 0 Then
                    dtbDataFillDate.Rows.Add(dtrSelectShoe2(0).ItemArray)
                    dtrDefault2 = dtrSelectShoe2(0)
                    strMediaSN = dtrSelectShoe2(0).Item("MediaSN")
                Else
                    If Not dtrDefault2 Is Nothing Then
                        dtbDataFillDate.Rows.Add(dtrDefault2.ItemArray)
                        dtbDataFillDate.Rows(dtbDataFillDate.Rows.Count - 1).Item("Date_time") = CDate(strTime)
                        If strMediaSN = "" Then strMediaSN = dtrDefault2.Item("MediaSN")
                    End If
                End If
                If dtrSelectShoe1.Length > 0 Or dtrSelectShoe2.Length > 0 Then
                    dtbDataFillDate.Rows(dtbDataFillDate.Rows.Count - 1).Item("MediaSN") = strMediaSN
                End If
            Next nHour
        Next nSearch
        FillBlankDateTime = dtbDataFillDate
    End Function

    Private Function CombineTableAllShoe(ByVal dtbCFByHourData As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        If dtbCFByHourData.Rows.Count = 0 Then Return dtbCFByHourData

        dtbCFByHourData.DefaultView.RowFilter = "Shoe='1'"
        Dim dtbShoe1 As DataTable = dtbCFByHourData.DefaultView.ToTable

        dtbCFByHourData.DefaultView.RowFilter = "Shoe='2'"
        Dim dtbShoe2 As DataTable = dtbCFByHourData.DefaultView.ToTable

        Dim dtbData As New DataTable
        Dim nShoe As Integer = dtbCFByHourData.Columns("Shoe").Ordinal
        For nCol As Integer = 1 To dtbCFByHourData.Columns.Count - 1
            Dim strColName As String = dtbCFByHourData.Columns(nCol).ColumnName
            If nCol < nShoe Or InStr(strColName, "CFMedia") > 0 Or InStr(strColName, "MediaSN") > 0 Or InStr(strColName, "DeltaGOS") > 0 Or InStr(strColName, "SigmaGOS") > 0 Then
                dtbData.Columns.Add(strColName, dtbCFByHourData.Columns(strColName).DataType)
            ElseIf nCol > nShoe Then
                dtbData.Columns.Add(strColName & ".Shoe1", dtbCFByHourData.Columns(strColName).DataType)
                dtbData.Columns.Add(strColName & ".Shoe2", dtbCFByHourData.Columns(strColName).DataType)
                dtbShoe1.Columns(strColName).ColumnName = strColName & ".Shoe1"
                dtbShoe2.Columns(strColName).ColumnName = strColName & ".Shoe2"
            End If
        Next nCol

        dtbShoe1.Columns.Remove("Shoe")
        dtbShoe2.Columns.Remove("Shoe")
        dtbData.Merge(dtbShoe1)
        Dim dcPrim(1) As DataColumn
        dcPrim(0) = dtbData.Columns("Date_time")
        dcPrim(1) = dtbData.Columns("Tester")
        dtbData.PrimaryKey = dcPrim
        dtbData.Merge(dtbShoe2)

        'Dim dtrDataAll() As DataRow
        'If dtrShoe1.Length > dtrShoe2.Length Then
        '    dtrDataAll = dtrShoe1
        'Else
        '    dtrDataAll = dtrShoe2
        'End If

        'For nRow As Integer = 0 To dtrDataAll.Length - 1
        '    Dim strSelect As String = ""
        '    For nCol As Integer = 1 To nShoe - 1
        '        Dim strColName As String = dtbCFByHourData.Columns(nCol).ColumnName
        '        Dim strValue As String = dtrDataAll(nRow).Item(nCol)
        '        strSelect = strSelect & "[" & strColName & "]='" & strValue & "' AND "
        '    Next nCol
        '    Dim dtrSelectShoe1() As DataRow = dtbCFByHourData.Select(strSelect & "Shoe='1'")
        '    Dim dtrSelectShoe2() As DataRow = dtbCFByHourData.Select(strSelect & "Shoe='2'")

        '    dtbData.Rows.Add()

        '    For nData As Integer = 1 To dtbCFByHourData.Columns.Count - 1
        '        Dim strColName As String = dtbCFByHourData.Columns(nData).ColumnName
        '        If nData > nShoe And InStr(strColName, "CFMedia") = 0 And InStr(strColName, "MediaSN") = 0 And InStr(strColName, "DeltaGOS") = 0 And InStr(strColName, "SigmaGOS") = 0 Then
        '            If dtrSelectShoe1.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName & ".Shoe1") = dtrSelectShoe1(0).Item(strColName)
        '            If dtrSelectShoe2.Length > 0 Then dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName & ".Shoe2") = dtrSelectShoe2(0).Item(strColName)
        '        ElseIf nData < nShoe Or InStr(strColName, "CFMedia") > 0 Or InStr(strColName, "MediaSN") > 0 Or InStr(strColName, "DeltaGOS") > 0 Then
        '            If dtrSelectShoe1.Length > 0 And dtrSelectShoe2.Length > 0 Then
        '                If dtrSelectShoe1(0).Item("MediaSN") <> dtrSelectShoe2(0).Item("MediaSN") Then
        '                    Dim nTimeShoe1 As Long = dtrSelectShoe1(0).Item("MaxTime")
        '                    Dim nTimeShoe2 As Long = dtrSelectShoe2(0).Item("MaxTime")
        '                    If nTimeShoe1 > nTimeShoe2 Then
        '                        dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrSelectShoe1(0).Item(strColName)
        '                    Else
        '                        dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrSelectShoe2(0).Item(strColName)
        '                    End If
        '                Else
        '                    dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrSelectShoe1(0).Item(strColName)
        '                End If
        '            ElseIf dtrSelectShoe1.Length = 0 And dtrSelectShoe2.Length > 0 Then
        '                dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrSelectShoe2(0).Item(strColName)
        '            ElseIf dtrSelectShoe1.Length > 0 And dtrSelectShoe2.Length = 0 Then
        '                dtbData.Rows(dtbData.Rows.Count - 1).Item(strColName) = dtrSelectShoe1(0).Item(strColName)
        '            End If
        '        End If
        '    Next nData
        'Next nRow
        CombineTableAllShoe = dtbData
    End Function

    Private Function GetDeltaGOSAVGByDay(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String
        Dim strSearchBy As String = dtbSearch.TableName
        strSQL = "SELECT "
        strSQL = strSQL & "DATE(test_time) Date_time,"
        strSQL = strSQL & "COUNT(DISTINCT Tester) CntTester,"
        strSQL = strSQL & "COUNT(tag_id) CntGOS,"
        strSQL = strSQL & strSearchBy & ","
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("param_display").ToString
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
            If bParamAdd Or bParamMul Then
                strSQL = strSQL & "AVG(" & strParam & ") """ & strParamDisplay & ".DeltaGOS"","
                strSQL = strSQL & "STD(" & strParam & ") """ & strParamDisplay & ".SigmaGOS"","
            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabfactor_deltagos A "
        strSQL = strSQL & "WHERE (A.test_time_bigint between '" & Format(dtStart, "yyyyMMddHHmmss") & "' AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        strSQL = strSQL & "GROUP BY DATE(test_time),A." & strSearchBy & " "
        strSQL = strSQL & "ORDER BY A." & strSearchBy & ",A.Test_time_bigint;"
        Dim clsMySQL As New CMySQL
        Dim dtbDeltaGOSByDay As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MySqlConn)
        GetDeltaGOSAVGByDay = dtbDeltaGOSByDay
    End Function

    Public Function GetCFAVGByDate(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, Optional ByVal strShoe As String = "") As DataTable
        Dim dtDateStart As DateTime = dtStart.Date
        Dim dtDateEnd As DateTime = dtEnd.Date
        dtDateEnd = dtDateEnd.AddHours(23)
        dtDateEnd = dtDateEnd.AddMinutes(59)
        dtDateEnd = dtDateEnd.AddSeconds(59)
        Dim dtbCFByHour As DataTable = GetCFByEndOfHour(strProduct, dtbSearch, dtbParam, dtDateStart, dtDateEnd, True)
        Dim nDateDiff As Integer = DateDiff(DateInterval.Day, dtDateStart, dtDateEnd)
        Dim nMediaIndex As Integer = dtbCFByHour.Columns("MediaSN").Ordinal
        Dim dtbDataByDate As DataTable
        dtbDataByDate = dtbCFByHour.Clone
        dtbDataByDate.Columns.RemoveAt(nMediaIndex)

        Dim dtbDeltaGOSByDay As DataTable = GetDeltaGOSAVGByDay(strProduct, dtbSearch, dtbParam, dtDateStart, dtDateEnd)
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            For nDate As Integer = 0 To nDateDiff
                Dim dtSumDate As DateTime = dtDateStart.AddDays(nDate)
                Dim strSearch As String = dtbSearch.Rows(nSearch).Item(dtbSearch.TableName)
                Dim strFilter As String = "[Date_time]>='" & Format(dtSumDate, "yyyy-MM-dd 00:00:00") & "' AND [Date_time]<='" & Format(dtSumDate, "yyyy-MM-dd 23:59:59") & "' "
                strFilter = strFilter & "AND [" & dtbSearch.TableName & "]='" & strSearch & "' "
                Dim dtrFilter() As DataRow = dtbCFByHour.Select(strFilter)
                Dim dtbTemp As DataTable
                dtbTemp = dtbCFByHour.Clone
                For nData As Integer = 0 To dtrFilter.Length - 1
                    dtbTemp.Rows.Add(dtrFilter(nData).ItemArray)
                Next nData
                If dtbTemp.Rows.Count > 0 Then
                    Dim strFilterGOS As String = "[Date_time]='" & Format(dtSumDate, "yyyy-MM-dd 00:00:00") & "' "
                    strFilterGOS = strFilterGOS & "AND [" & dtbSearch.TableName & "]='" & strSearch & "' "
                    Dim dtrDeltaGOS() As DataRow = dtbDeltaGOSByDay.Select(strFilterGOS)

                    dtbDataByDate.Rows.Add()
                    For nCol As Integer = 0 To dtbCFByHour.Columns.Count - 1
                        Dim strColName As String = dtbCFByHour.Columns(nCol).ColumnName
                        If nCol < nMediaIndex And InStr(strColName, "Cnt") = 0 Then
                            dtbDataByDate.Rows(dtbDataByDate.Rows.Count - 1).Item(strColName) = dtbTemp.Rows(0).Item(strColName)
                        ElseIf nCol > nMediaIndex And InStr(strColName, "DeltaGOS") = 0 And InStr(strColName, "SigmaGOS") = 0 And InStr(strColName, "Cnt") = 0 Then
                            dtbTemp.Columns.Add("AVG." & strColName, Type.GetType("System.Double"), "AVG([" & strColName & "])")
                            dtbDataByDate.Rows(dtbDataByDate.Rows.Count - 1).Item(strColName) = dtbTemp.Rows(0).Item("AVG." & strColName)
                            dtbTemp.Columns.Remove("AVG." & strColName)
                        ElseIf InStr(strColName, "DeltaGOS") > 0 Or InStr(strColName, "SigmaGOS") > 0 Or InStr(strColName, "Cnt") > 0 Then
                            If dtrDeltaGOS.Length = 0 Then
                                dtbDataByDate.Rows(dtbDataByDate.Rows.Count - 1).Item(strColName) = System.DBNull.Value
                            Else
                                dtbDataByDate.Rows(dtbDataByDate.Rows.Count - 1).Item(strColName) = dtrDeltaGOS(0).Item(strColName)
                            End If
                        End If
                    Next nCol
                End If
            Next nDate
        Next nSearch
        GetCFAVGByDate = dtbDataByDate
    End Function

    Public Function GetCFHistoryByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, Optional ByVal strShoe As String = "") As DataTable
        Dim strSQL As String = ""
        Dim clsMySQL As New CMySQL
        Dim dtbCFHistory As New DataTable

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1

            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
            If bParamAdd Or bParamMul Then
                strSQL = "SELECT "
                strSQL = strSQL & "A.ChangeTime,"
                strSQL = strSQL & "A.Tester,"
                strSQL = strSQL & "A.Shoe,"
                strSQL = strSQL & "B.Param_rttc,"
                strSQL = strSQL & "A.CFValue "
                If bParamAdd Then
                    strSQL = strSQL & "FROM cf_" & strProduct & ".tabhistory_cfadd A "
                    strSQL = strSQL & "LEFT JOIN db_parameter.tabparamapping B USING(paramID) "
                ElseIf bParamMul Then
                    strSQL = strSQL & "FROM cf_" & strProduct & ".tabhistory_cfmul A "
                    strSQL = strSQL & "LEFT JOIN db_parameter.tabparamapping B USING(paramID) "
                End If
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "(A.ChangeTime between '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' and '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
                strSQL = strSQL & "AND ("
                Dim strSearchBy As String = dtbSearch.TableName
                For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
                    If nTester <> dtbSearch.Rows.Count - 1 Then
                        strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
                    Else
                        strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
                    End If
                Next nTester
                If strShoe <> "" Then
                    strSQL = strSQL & "AND A.Shoe='" & strShoe & "' "
                End If
                strSQL = strSQL & "AND B.Param_rttc='" & strParam & "' "
                strSQL = strSQL & "ORDER BY Tester,shoe,param_rttc,ChangeTime;"

                Dim dtbData As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MySqlConn)
                dtbCFHistory.Merge(dtbData)
            End If
        Next nParam

        GetCFHistoryByTester = dtbCFHistory
    End Function

    Public Function GetCFReportByType(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String = "SELECT CFDate,"
        strSQL = strSQL & "IF(MachineType=0,'Up','Down') MachineType,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("Param_Display").ToString
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
            If bParamAdd = True Then
                strSQL = strSQL & "(A." & strParam & ") """ & strParamDisplay & ".CFAdd"","
            End If
            If bParamMul = True Then
                strSQL = strSQL & "(B." & strParam & ") """ & strParamDisplay & ".CFMul"","
            End If
            If bParamAdd Or bParamMul Then
                strSQL = strSQL & "(C." & strParam & ") """ & strParamDisplay & ".CFMedia"","
            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabdaily_cfbytype_add A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdaily_cfbytype_mul B USING(CFDate,MachineType) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdaily_cfmedia C USING(CFDate,MachineType) "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "(A.CFDate between '" & Format(dtStart, "yyyy-MM-dd") & "' AND '" & Format(dtEnd, "yyyy-MM-dd") & "') "
        strSQL = strSQL & "AND ("
        Dim strSearchBy As String = "MachineType"
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item("OptionIndex") & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item("OptionIndex") & "') "
            End If
        Next nSearch
        strSQL = strSQL & "ORDER BY A.CFDate,MachineType;"
        Dim clsMysql As New CMySQL
        GetCFReportByType = clsMysql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

    Public Function GetDailyCFReportByType(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String = "SELECT CFDate,"
        strSQL = strSQL & "IF(MachineType=0,'Up','Down') MachineType,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamDisplay As String = dtbParam.Rows(nParam).Item("Param_Display").ToString
            Dim bParamAdd As Boolean = dtbParam.Rows(nParam).Item("Param_add")
            Dim bParamMul As Boolean = dtbParam.Rows(nParam).Item("Param_mul")
            If bParamAdd = True Then
                strSQL = strSQL & "AVG(A." & strParam & ") """ & strParamDisplay & ".CFAdd"","
            End If
            If bParamMul = True Then
                strSQL = strSQL & "AVG(B." & strParam & ") """ & strParamDisplay & ".CFMul"","
            End If
            If bParamAdd Or bParamMul Then
                strSQL = strSQL & "AVG(C." & strParam & ") """ & strParamDisplay & ".CFMedia"","
            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabdaily_cfbytype_add A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdaily_cfbytype_mul B USING(CFDate,MachineType) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdaily_cfmedia C USING(CFDate,MachineType) "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "(A.CFDate between '" & Format(dtStart, "yyyy-MM-dd 00:00:00") & "' AND '" & Format(dtEnd, "yyyy-MM-dd 23:59:59") & "') "
        strSQL = strSQL & "AND ("
        Dim strSearchBy As String = "MachineType"
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item("OptionIndex") & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item("OptionIndex") & "') "
            End If
        Next nSearch
        strSQL = strSQL & "GROUP BY DATE_FORMAT(A.CFDate,'%Y%m%d'),MachineType "
        strSQL = strSQL & "ORDER BY A.CFDate,MachineType;"
        Dim clsMysql As New CMySQL
        GetDailyCFReportByType = clsMysql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

End Class

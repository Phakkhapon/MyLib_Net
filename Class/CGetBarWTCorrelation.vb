
Imports MySql.Data.MySqlClient

Public Class CGetBarWTCorrelation
    Private m_mySqlConn As MySqlConnection
    Private m_dtbSection As DataTable
    Private m_dtbSldSection As DataTable
    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
        m_dtbSection = InitSectionTable()
        m_dtbSldSection = InitSldSectionTable()

    End Sub

    Public Function GetBWTCorrelation(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, _
        ByVal dtStart As DateTime, ByVal dtEnd As DateTime, Optional ByVal bIsMapLotName As Boolean = False, Optional ByVal bGetCF As Boolean = False) As DataTable

        Dim dtbData As DataTable = GetCorrelationData(strProduct, dtbSearchBy, dtStart, dtEnd)
        If Not dtbData Is Nothing Then
            Dim dcSF As DataColumn = dtbData.Columns.Add("SF_Section")
            dcSF.SetOrdinal(dtbData.Columns("Hga_SN").Ordinal + 1)
            'Dim strFilter As String = "[MEW.Delta]>1.5 OR [MEW.Delta]<-1.5 OR [MEW.Delta] IS NULL"
            'Dim dtrFilter() As DataRow = dtbData.Select(strFilter)
            'For nRemove As Integer = 0 To dtrFilter.Length - 1
            '    dtbData.Rows.Remove(dtrFilter(nRemove))
            'Next nRemove
            dtbData.Columns.Remove("STD_AvgDiff")
            dtbData.Columns.Remove("STD_Slope")
            dtbData.Columns.Remove("STD_STDVDiff")
            dtbData.Columns.Remove("STD_RSQ")
            For nData As Integer = 0 To dtbData.Rows.Count - 1
                Dim strHGA_SN As String = dtbData.Rows(nData).Item("Hga_SN")
                dtbData.Rows(nData).Item("SF_Section") = GetSliderFabSection(strHGA_SN)
            Next nData
        End If
        GetBWTCorrelation = dtbData
    End Function

    Private Function GetCorrelationData(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, _
        ByVal dtStart As DateTime, ByVal dtEnd As DateTime, Optional ByVal bIsMapLotName As Boolean = False, Optional ByVal bGetCF As Boolean = False) As DataTable
        If InStr(strProduct, "BWT") = 1 Then
            Dim dtbData As New DataTable
            Dim clsProduct As New CParameterRTTCMapping(m_mySqlConn)
            Dim dtbXLotProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductXLot)
            dtbXLotProduct.Merge(clsProduct.GetProductList(enuProductType.enuProductNPL))
            For nXLot As Integer = 0 To dtbXLotProduct.Rows.Count - 1
                Dim strProductXLot As String = dtbXLotProduct.Rows(nXLot).Item("Product")
                Dim strSQL As String = ""
                strSQL = "SELECT "
                strSQL = strSQL & "'" & strProductXLot & "' MTProduct,"
                strSQL = strSQL & "A.Test_Time ""TestTime.BWT"","
                strSQL = strSQL & "C.Test_time ""TestTime.MT"","
                strSQL = strSQL & "A.Tester ""Tester.BWT"","
                strSQL = strSQL & "C.Tester ""Tester.MT"","
                strSQL = strSQL & "A.Lot ""Lot.BWT"","
                strSQL = strSQL & "C.Lot ""Lot.MT"","
                strSQL = strSQL & "IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>=75,'XLot2',IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>0,'XLot1','SisterLot')) 'LotType',"
                strSQL = strSQL & "A.Assy Device,"
                strSQL = strSQL & "A.Spec ""Spec.BWT"","
                strSQL = strSQL & "C.Spec ""Spec.MT"","
                strSQL = strSQL & "A.Hga_sn,"
                strSQL = strSQL & "C.MediaSN,"
                strSQL = strSQL & "C.TrackID,"
                strSQL = strSQL & "C.BarNo,"
                strSQL = strSQL & "C.TrayID,"
                strSQL = strSQL & "C.CGALot,"
                strSQL = strSQL & "C.CGANo,"
                strSQL = strSQL & "C.SliderSite,"
                strSQL = strSQL & "B.MEW6T ""MEW6T.BWT"","
                strSQL = strSQL & "D.WEW ""WEW.MT"","
                strSQL = strSQL & "D.WEW-B.MEW6T ""MEW.Delta"","
                If bGetCF Then
                    strSQL = strSQL & "K.WEW ""WEW.CF"","
                    strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabfactor_media M WHERE M.DiskSN=LEFT(C.MediaSN,LENGTH(C.MediaSN)-3)) ""WEW.Media"","
                End If

                'strSQL = strSQL & "(SELECT WEW FROM db_" & strProductXLot & ".tabdetail_header M LEFT JOIN db_" & strProductXLot & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(C.Test_time_bigint) AND Hga_sn=C.Hga_sn) ""WEW.MT"","
                'strSQL = strSQL & "(SELECT WEW FROM db_" & strProductXLot & ".tabdetail_header M LEFT JOIN db_" & strProductXLot & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(C.Test_time_bigint) AND Hga_sn=C.Hga_sn)-B.MEW6T ""MEW.Delta"","

                strSQL = strSQL & "D.TdV,"
                strSQL = strSQL & "D.TdP,"
                strSQL = strSQL & "D.TdTyp,"
                strSQL = strSQL & "D.RevOwr,"
                strSQL = strSQL & "D.SNRTotal,"
                strSQL = strSQL & "B.para1 STD_AvgDiff,"
                strSQL = strSQL & "B.para2 STD_Slope,"
                strSQL = strSQL & "B.para3 STD_STDVDiff,"
                strSQL = strSQL & "B.para4 STD_RSQ "
                strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
                If bGetCF Then
                    strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd K USING(tag_id) "
                End If
                strSQL = strSQL & ","
                strSQL = strSQL & "db_" & strProductXLot & ".tabdetail_header C LEFT JOIN db_" & strProductXLot & ".tabfactor_value D USING(tag_id) "

                Dim strSearchBy As String = dtbSearchBy.TableName
                strSQL = strSQL & "WHERE (A.Test_Time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
                strSQL = strSQL & "AND ("
                For nTester As Integer = 0 To dtbSearchBy.Rows.Count - 1
                    If nTester <> dtbSearchBy.Rows.Count - 1 Then
                        strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "' OR "
                    Else
                        strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "') "
                    End If
                Next nTester
                strSQL = Left(strSQL, Len(strSQL) - 1) & " "
                strSQL = strSQL & "AND C.Test_time_bigint=(SELECT Max(Test_time_bigint) FROM db_" & strProductXLot & ".tabdetail_header M "
                strSQL = strSQL & "WHERE "
                If bIsMapLotName Then
                    strSQL = strSQL & "Lot=A.Lot AND "
                End If
                strSQL = strSQL & "Hga_sn=A.Hga_sn) "
                If bIsMapLotName Then
                    strSQL = strSQL & "AND (A.Lot=C.Lot OR REPLACE(A.Lot,'&','M')=REPLACE(C.Lot,'&','M')) "
                End If
                strSQL = strSQL & "AND A.Hga_SN=C.Hga_SN "
                strSQL = strSQL & "AND A.TestMode=0 "
                strSQL = strSQL & "GROUP BY A.Hga_SN "
                strSQL = strSQL & "ORDER BY C.test_time_bigint,A.Tester;"
                Dim clsMySql As New CMySQL
                Dim dtbDataXLot As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
                dtbData.Merge(dtbDataXLot)
            Next nXLot
            GetCorrelationData = dtbData

        ElseIf InStr(strProduct, "XLOT") <> 0 Or InStr(strProduct, "NPL") <> 0 Then
            Dim strSQL As String = ""
            strSQL = "SELECT "
            strSQL = strSQL & "'" & strProduct & "' MTProduct,"
            strSQL = strSQL & "C.Test_Time ""TestTime.BWT"","
            strSQL = strSQL & "A.Test_time ""TestTime.MT"","
            strSQL = strSQL & "C.Tester ""Tester.BWT"","
            strSQL = strSQL & "A.Tester ""Tester.MT"","
            strSQL = strSQL & "C.Lot ""Lot.BWT"","
            strSQL = strSQL & "A.Lot ""Lot.MT"","
            strSQL = strSQL & "IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>=75,'XLot2',IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>0,'XLot1','SisterLot')) 'LotType',"
            strSQL = strSQL & "C.Assy Device,"
            strSQL = strSQL & "C.Spec ""Spec.BWT"","
            strSQL = strSQL & "A.Spec ""Spec.MT"","
            strSQL = strSQL & "A.Hga_sn,"
            strSQL = strSQL & "A.GradeName,"
            strSQL = strSQL & "A.MediaSN,"
            strSQL = strSQL & "A.TrackID,"
            strSQL = strSQL & "A.BarNo,"
            strSQL = strSQL & "A.TrayID,"
            strSQL = strSQL & "A.CGALot,"
            strSQL = strSQL & "A.CGANo,"
            strSQL = strSQL & "A.SliderSite,"
            strSQL = strSQL & "D.MEW6T ""MEW6T.BWT"","
            'strSQL = strSQL & "B.WEW ""WEW.MT"","
            'strSQL = strSQL & "B.WEW-D.MEW6T ""MEW.Delta"","
            strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1) ""WEW.MT"","
            strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1)-D.MEW6T ""MEW.Delta"","
            If bGetCF Then
                strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_cfadd N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1) ""WEW.CF"","
                strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabfactor_media M WHERE M.DiskSN=LEFT(A.MediaSN,LENGTH(A.MediaSN)-3)) ""WEW.Media"","

            End If
            strSQL = strSQL & "B.TdV,"
            strSQL = strSQL & "B.TdP,"
            strSQL = strSQL & "B.TdTyp,"
            strSQL = strSQL & "B.RevOwr,"
            strSQL = strSQL & "B.SNRTotal,"
            strSQL = strSQL & "D.para1 STD_AvgDiff,"
            strSQL = strSQL & "D.para2 STD_Slope,"
            strSQL = strSQL & "D.para3 STD_STDVDiff,"
            strSQL = strSQL & "D.para4 STD_RSQ "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id),"
            strSQL = strSQL & "db_BWT_BWT.tabdetail_header C LEFT JOIN db_BWT_BWT.tabfactor_value D USING(tag_id) "

            Dim strSearchBy As String = dtbSearchBy.TableName
            strSQL = strSQL & "WHERE (A.Test_Time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
            strSQL = strSQL & "AND ("
            For nTester As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nTester <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "') "
                End If
            Next nTester
            strSQL = Left(strSQL, Len(strSQL) - 1) & " "
            If bIsMapLotName Then
                'strSQL = strSQL & "AND A.Lot=C.Lot "
                strSQL = strSQL & "AND (A.Lot=C.Lot OR REPLACE(A.Lot,'&','M')=REPLACE(C.Lot,'&','M')) "
            End If
            strSQL = strSQL & "AND A.Hga_SN=C.Hga_SN "
            strSQL = strSQL & "AND A.TestMode=0 "
            strSQL = strSQL & "GROUP BY A.Hga_SN "
            strSQL = strSQL & "ORDER BY A.test_time_bigint,C.Tester;"

            Dim clsMySql As New CMySQL
            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
            GetCorrelationData = dtbData
        Else
            Return Nothing
        End If

    End Function

    ''Private Function GetRawCorrelationDataByLot(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, _
    ''ByVal dtStart As DateTime, ByVal dtEnd As DateTime, Optional ByVal bIsMapLotName As Boolean = False, Optional ByVal bGetCF As Boolean = False) As DataTable
    ''    If InStr(strProduct, "BWT") = 1 Then
    ''        Dim dtbData As New DataTable
    ''        Dim clsProduct As New CParameterRTTCMapping(m_mySqlConn)
    ''        Dim dtbXLotProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductXLot)
    ''        dtbXLotProduct.Merge(clsProduct.GetProductList(enuProductType.enuProductNPL))
    ''        For nXLot As Integer = 0 To dtbXLotProduct.Rows.Count - 1
    ''            Dim strProductXLot As String = dtbXLotProduct.Rows(nXLot).Item("Product")
    ''            Dim strSQL As String = ""
    ''            strSQL = "SELECT "
    ''            strSQL = strSQL & "'" & strProductXLot & "' MTProduct,"
    ''            strSQL = strSQL & "MAX(A.Test_Time) 'TestTime.BWT',"
    ''            strSQL = strSQL & "MAX(C.Test_time) 'TestTime.MT',"
    ''            strSQL = strSQL & "A.Tester 'Tester.BWT',"
    ''            strSQL = strSQL & "C.Tester 'Tester.MT',"
    ''            strSQL = strSQL & "A.Lot 'Lot.BWT',"
    ''            strSQL = strSQL & "C.Lot 'Lot.MT',"
    ''            strSQL = strSQL & "IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>=75,'XLot2',IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>0,'XLot1','SisterLot')) 'LotType',"
    ''            strSQL = strSQL & "A.Assy Device,"
    ''            strSQL = strSQL & "A.Spec 'Spec.BWT',"
    ''            strSQL = strSQL & "C.Spec 'Spec.MT',"
    ''            strSQL = strSQL & "A.Hga_sn,"
    ''            strSQL = strSQL & "C.MediaSN,"
    ''            strSQL = strSQL & "C.TrackID,"
    ''            strSQL = strSQL & "C.BarNo,"
    ''            strSQL = strSQL & "C.TrayID,"
    ''            strSQL = strSQL & "C.CGALot,"
    ''            strSQL = strSQL & "C.CGANo,"
    ''            strSQL = strSQL & "AVG(B.MEW6T) 'MEW6T.BWT',"
    ''            strSQL = strSQL & "AVG(D.WEW) 'WEW.MT',"
    ''            strSQL = strSQL & "AVG(D.WEW-B.MEW6T) 'MEW.Delta',"
    ''            If bGetCF Then
    ''                strSQL = strSQL & "AVG(K.WEW) 'WEW.CF',"
    ''                strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabfactor_media M WHERE M.DiskSN=LEFT(C.MediaSN,LENGTH(C.MediaSN)-3)) ""WEW.Media"","
    ''            End If
    ''            strSQL = strSQL & "AVG(D.TdV) 'TdV',"
    ''            strSQL = strSQL & "AVG(D.TdP) 'TdP',"
    ''            strSQL = strSQL & "AVG(D.TdTyp) 'TdTyp',"
    ''            strSQL = strSQL & "AVG(D.RevOwr) 'RevOwr',"
    ''            strSQL = strSQL & "AVG(D.SNRTotal) 'SNRTotal',"
    ''            strSQL = strSQL & "AVG(B.para1) 'STD_AvgDiff',"
    ''            strSQL = strSQL & "AVG(B.para2) 'STD_Slope',"
    ''            strSQL = strSQL & "AVG(B.para3) 'STD_STDVDiff',"
    ''            strSQL = strSQL & "AVG(B.para4) 'STD_RSQ' "
    ''            strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A "
    ''            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
    ''            If bGetCF Then
    ''                strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd K USING(tag_id) "
    ''            End If
    ''            strSQL = strSQL & ","
    ''            strSQL = strSQL & "db_" & strProductXLot & ".tabdetail_header C "
    ''            strSQL = strSQL & "LEFT JOIN db_" & strProductXLot & ".tabfactor_value D USING(tag_id) "

    ''            Dim strSearchBy As String = dtbSearchBy.TableName
    ''            strSQL = strSQL & "WHERE (A.Test_Time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
    ''            strSQL = strSQL & "AND ("
    ''            For nTester As Integer = 0 To dtbSearchBy.Rows.Count - 1
    ''                If nTester <> dtbSearchBy.Rows.Count - 1 Then
    ''                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "' OR "
    ''                Else
    ''                    strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "') "
    ''                End If
    ''            Next nTester
    ''            strSQL = Left(strSQL, Len(strSQL) - 1) & " "
    ''            'strSQL = strSQL & "AND C.Test_time_bigint=(SELECT Max(Test_time_bigint) FROM db_" & strProductXLot & ".tabdetail_header M "
    ''            'strSQL = strSQL & "WHERE "
    ''            'If bIsMapLotName Then
    ''            'strSQL = strSQL & "Lot=A.Lot AND "
    ''            'End If
    ''            'strSQL = strSQL & "Hga_sn=A.Hga_sn) "
    ''            If bIsMapLotName Then
    ''                strSQL = strSQL & "AND A.Lot=C.Lot "
    ''            End If
    ''            strSQL = strSQL & "AND A.Hga_SN=C.Hga_SN "
    ''            strSQL = strSQL & "AND A.TestMode=0 "
    ''            strSQL = strSQL & "GROUP BY A." & strSearchBy & ",A.Lot "
    ''            'strSQL = strSQL & "GROUP BY A.Hga_SN "
    ''            strSQL = strSQL & "ORDER BY C.test_time_bigint,A.Tester;"
    ''            Dim clsMySql As New CMySQL
    ''            Dim dtbDataXLot As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
    ''            dtbData.Merge(dtbDataXLot)
    ''        Next nXLot
    ''        GetRawCorrelationDataByLot = dtbData

    ''    ElseIf InStr(strProduct, "XLOT") <> 0 Or InStr(strProduct, "NPL") <> 0 Then
    ''        Dim strSQL As String = ""
    ''        strSQL = "SELECT "
    ''        strSQL = strSQL & "'" & strProduct & "' 'MTProduct',"
    ''        strSQL = strSQL & "COUNT(tag_id) TotalHGA,"
    ''        strSQL = strSQL & "COUNT(A.WEW-C.MEW6T>1.5 OR A.WEW-C.MEW6T<-1.5 OR C.MEW6T IS NULL OR A.WEW IS NULL) OutlierHead,"
    ''        'dtbSum.Columns.Add("%OutlierHead", Type.GetType("System.Double"), "OutlierHead/TotalHGA*100")
    ''        strSQL = strSQL & "A.Test_time ""TestTime.MT"","
    ''        strSQL = strSQL & "C.Test_Time ""TestTime.BWT"","
    ''        strSQL = strSQL & "C.Tester ""Tester.BWT"","
    ''        strSQL = strSQL & "A.Tester ""Tester.MT"","
    ''        strSQL = strSQL & "A.Spec ""Spec.MT"","
    ''        strSQL = strSQL & "A.Lot ""Lot.MT"","
    ''        strSQL = strSQL & "IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>=75,'XLot2',IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>0,'XLot1','SisterLot')) 'LotType',"
    ''        strSQL = strSQL & "C.Assy Device,"

    ''        strSQL = strSQL & "AVG(D.para1) STD_AvgDiff,"
    ''        strSQL = strSQL & "AVG(D.para2) STD_Slope,"
    ''        strSQL = strSQL & "AVG(D.para3) STD_STDVDiff,"
    ''        strSQL = strSQL & "AVG(D.para4) STD_RSQ,"
    ''        strSQL = strSQL & "AVG(D.MEW6T) 'MEW6T.BWT',"
    ''        strSQL = strSQL & "AVG(B.WEW) 'WEW.MT',"
    ''        strSQL = strSQL & "STDEV(D.MEW6T) 'MEW6T.BWT.Stdev',"
    ''        strSQL = strSQL & "STDEV(B.WEW) 'WEW.MT.Stdev',"
    ''        strSQL = strSQL & "AVG(A.WEW-C.MEW6T) 'MEW.Delta',"

    ''        strSQL = strSQL & "B.TdV,"
    ''        strSQL = strSQL & "B.TdP,"
    ''        strSQL = strSQL & "B.TdTyp,"
    ''        strSQL = strSQL & "B.RevOwr,"
    ''        strSQL = strSQL & "B.SNRTotal,"

    ''        dtbSum.Columns.Add("AVG.CF.MT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("AVG.Media.MT", GetType(Double))
    ''        dtbSum.Columns.Add("Stdev.CF.MT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("Slope.BWT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("Intercept.BWT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("RSQ.BWT.MT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("Slope.MT", Type.GetType("System.Double"))
    ''        dtbSum.Columns.Add("Intercept.MT", Type.GetType("System.Double"))

    ''        strSQL = strSQL & "'" & strProduct & "' MTProduct,"
    ''        strSQL = strSQL & "C.Test_Time ""TestTime.BWT"","
    ''        strSQL = strSQL & "A.Test_time ""TestTime.MT"","
    ''        strSQL = strSQL & "C.Tester ""Tester.BWT"","
    ''        strSQL = strSQL & "A.Tester ""Tester.MT"","
    ''        strSQL = strSQL & "C.Lot ""Lot.BWT"","
    ''        strSQL = strSQL & "A.Lot ""Lot.MT"","
    ''        strSQL = strSQL & "IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>=75,'XLot2',IF(CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(A.Lot,LENGTH(A.Lot)-1) AS SIGNED)>0,'XLot1','SisterLot')) 'LotType',"
    ''        strSQL = strSQL & "C.Assy Device,"
    ''        strSQL = strSQL & "C.Spec ""Spec.BWT"","
    ''        strSQL = strSQL & "A.Spec ""Spec.MT"","
    ''        strSQL = strSQL & "A.Hga_sn,"
    ''        strSQL = strSQL & "A.GradeName,"
    ''        strSQL = strSQL & "A.MediaSN,"
    ''        strSQL = strSQL & "A.TrackID,"
    ''        strSQL = strSQL & "A.BarNo,"
    ''        strSQL = strSQL & "A.TrayID,"
    ''        strSQL = strSQL & "A.CGALot,"
    ''        strSQL = strSQL & "A.CGANo,"
    ''        strSQL = strSQL & "D.MEW6T ""MEW6T.BWT"","
    ''        'strSQL = strSQL & "B.WEW ""WEW.MT"","
    ''        'strSQL = strSQL & "B.WEW-D.MEW6T ""MEW.Delta"","
    ''        strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1) ""WEW.MT"","
    ''        strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_value N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1)-D.MEW6T ""MEW.Delta"","
    ''        If bGetCF Then
    ''            strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabdetail_header M LEFT JOIN db_" & strProduct & ".tabfactor_cfadd N USING(tag_id) WHERE M.test_time_bigint=Max(A.Test_time_bigint) AND Hga_sn=A.Hga_sn LIMIT 0,1) ""WEW.CF"","
    ''            strSQL = strSQL & "(SELECT WEW FROM db_" & strProduct & ".tabfactor_media M WHERE M.DiskSN=LEFT(A.MediaSN,LENGTH(A.MediaSN)-3)) ""WEW.Media"","

    ''        End If
    ''        strSQL = strSQL & "B.TdV,"
    ''        strSQL = strSQL & "B.TdP,"
    ''        strSQL = strSQL & "B.TdTyp,"
    ''        strSQL = strSQL & "B.RevOwr,"
    ''        strSQL = strSQL & "B.SNRTotal,"
    ''        strSQL = strSQL & "D.para1 STD_AvgDiff,"
    ''        strSQL = strSQL & "D.para2 STD_Slope,"
    ''        strSQL = strSQL & "D.para3 STD_STDVDiff,"
    ''        strSQL = strSQL & "D.para4 STD_RSQ "
    ''        strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id),"
    ''        strSQL = strSQL & "db_BWT_BWT.tabdetail_header C LEFT JOIN db_BWT_BWT.tabfactor_value D USING(tag_id) "

    ''        Dim strSearchBy As String = dtbSearchBy.TableName
    ''        strSQL = strSQL & "WHERE (A.Test_Time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
    ''        strSQL = strSQL & "AND ("
    ''        For nTester As Integer = 0 To dtbSearchBy.Rows.Count - 1
    ''            If nTester <> dtbSearchBy.Rows.Count - 1 Then
    ''                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "' OR "
    ''            Else
    ''                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "') "
    ''            End If
    ''        Next nTester
    ''        strSQL = Left(strSQL, Len(strSQL) - 1) & " "
    ''        If bIsMapLotName Then
    ''            strSQL = strSQL & "AND A.Lot=C.Lot "
    ''        End If
    ''        strSQL = strSQL & "AND A.Hga_SN=C.Hga_SN "
    ''        strSQL = strSQL & "AND A.TestMode=0 "
    ''        strSQL = strSQL & "GROUP BY A.Hga_SN "
    ''        strSQL = strSQL & "ORDER BY A.test_time_bigint,C.Tester;"

    ''        Dim clsMySql As New CMySQL
    ''        Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
    ''        GetRawCorrelationDataByLot = dtbData
    ''    Else
    ''        Return Nothing
    ''    End If

    ''End Function

    Public Function GetSummaryCorrelationByLot(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal bAddSummaryByDate As Boolean) As DataTable

        Dim dtbData As DataTable = GetCorrelationData(strProduct, dtbSearchBy, dtStart, dtEnd, True, True)

        If Not dtbData Is Nothing Then
            Dim strSearchBy As String = dtbSearchBy.TableName & ".MT"
            Dim distinctLot As DataTable = dtbData.DefaultView.ToTable(True, New String() {strSearchBy, "Lot.BWT"})

            Dim dtbSum As New DataTable            'To store summary data
            dtbSum.Columns.Add("MTProduct")
            dtbSum.Columns.Add("TotalHGA", Type.GetType("System.Int32"))
            dtbSum.Columns.Add("OutlierHead", Type.GetType("System.Int32"))
            dtbSum.Columns.Add("%OutlierHead", Type.GetType("System.Double"), "OutlierHead/TotalHGA*100")
            dtbSum.Columns.Add("TestTime.MT", Type.GetType("System.DateTime"))
            dtbSum.Columns.Add("TestTime.BWT", Type.GetType("System.DateTime"))
            dtbSum.Columns.Add("Spec.MT")
            dtbSum.Columns.Add("Tester.BWT")
            dtbSum.Columns.Add("Lot.MT")
            dtbSum.Columns.Add("LotType")
            dtbSum.Columns.Add("Device")
            dtbSum.Columns.Add("SliderSite")
            'dtbSum.Columns.Add("Lot.BWT")
            If dtbSum.Columns(strSearchBy) Is Nothing Then dtbSum.Columns.Add(strSearchBy)
            dtbSum.Columns.Add("STD_AvgDiff", Type.GetType("System.Double"))
            dtbSum.Columns.Add("STD_Slope", Type.GetType("System.Double"))
            dtbSum.Columns.Add("STD_STDVDiff", Type.GetType("System.Double"))
            dtbSum.Columns.Add("STD_RSQ", Type.GetType("System.Double"))
            dtbSum.Columns.Add("MEW6T.BWT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("WEW.MT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("MEW6T.BWT.Stdev", Type.GetType("System.Double"))
            dtbSum.Columns.Add("WEW.MT.Stdev", Type.GetType("System.Double"))
            dtbSum.Columns.Add("MEW.Delta", Type.GetType("System.Double"), "WEW.MT-MEW6T.BWT")
            dtbSum.Columns.Add("AVG.CF.MT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("AVG.Media.MT", GetType(Double))
            dtbSum.Columns.Add("Stdev.CF.MT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("Slope.BWT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("Intercept.BWT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("RSQ.BWT.MT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("Slope.MT", Type.GetType("System.Double"))
            dtbSum.Columns.Add("Intercept.MT", Type.GetType("System.Double"))

            For nLot As Integer = 0 To distinctLot.Rows.Count - 1
                Dim strSearch As String = distinctLot.Rows(nLot).Item(strSearchBy)
                Dim strLot As String = distinctLot.Rows(nLot).Item("Lot.BWT")
                Dim strInitCondition As String = "[" & strSearchBy & "]='" & strSearch & "' AND [Lot.BWT]='" & strLot & "'"

                dtbData.DefaultView.RowFilter = strInitCondition
                Dim dtbDataBySpecByLot As DataTable = dtbData.DefaultView.ToTable
                Dim strNoOutlier As String = "[MEW.Delta]>=-1.5 AND [MEW.Delta]<=1.5 AND [MEW.Delta] IS NOT NULL"
                Dim strOutlier As String = "([MEW.Delta]>1.5 OR [MEW.Delta]<-1.5 OR [MEW.Delta] IS NULL)"

                dtbDataBySpecByLot.DefaultView.RowFilter = strNoOutlier
                Dim dtbNoOutlier As DataTable = dtbDataBySpecByLot.DefaultView.ToTable
                Dim drOutlier() As DataRow = dtbDataBySpecByLot.Select(strOutlier)
                If dtbNoOutlier.Rows.Count > 0 Then
                    Dim strTesterBWT As String = dtbNoOutlier.Rows(0).Item("Tester.BWT")
                    AssignValueToSummaryTable(strSearchBy, strSearch, strLot, strTesterBWT, dtbSum, dtbNoOutlier, drOutlier.Length)
                    dtbNoOutlier.Dispose()
                End If

            Next nLot
            If bAddSummaryByDate And dtbSum.Rows.Count > 0 Then
                Dim strFilterByDate As String = "[MEW.Delta]>=-1.5 AND [MEW.Delta]<=1.5 AND [MEW.Delta] IS NOT NULL"
                dtbData.DefaultView.RowFilter = strFilterByDate
                Dim dtbByDate As DataTable = dtbData.DefaultView.ToTable
                If dtbByDate.Rows.Count > 0 Then
                    Dim strTesterBWT As String = "All Tester"
                    dtbSum.Rows.Add()
                    dtbSum.Rows(dtbSum.Rows.Count - 1).Item("MTProduct") = "SummaryByDate"
                    AssignValueToSummaryTable(strSearchBy, "SummaryByDate", "All", strTesterBWT, dtbSum, dtbByDate, dtbData.Rows.Count - dtbByDate.Rows.Count)
                End If

                Try
                    Dim strFilterByDateXLot1 As String = "[MEW.Delta]>=-1.5 AND [MEW.Delta]<=1.5 AND [MEW.Delta] IS NOT NULL"
                    strFilterByDateXLot1 = strFilterByDateXLot1 & " AND SUBSTRING([Lot.MT],LEN([Lot.MT])-1,2)<75"
                    dtbData.DefaultView.RowFilter = strFilterByDateXLot1
                    Dim dtbByDateXLot1 As DataTable = dtbData.DefaultView.ToTable
                    Dim strOutlier As String = "([MEW.Delta]>1.5 OR [MEW.Delta]<-1.5 OR [MEW.Delta] IS NULL) AND SUBSTRING([Lot.MT],LEN([Lot.MT])-1,2)<75"
                    Dim drOutlierXLot1() As DataRow = dtbData.Select(strOutlier)
                    If dtbByDateXLot1.Rows.Count > 0 Then
                        Dim strTesterBWT As String = "BWTXLot1"
                        dtbSum.Rows.Add()
                        dtbSum.Rows(dtbSum.Rows.Count - 1).Item("MTProduct") = "SummaryByDate.XLot1"
                        AssignValueToSummaryTable(strSearchBy, "SummaryByDate.XLot1", "XLot1", strTesterBWT, dtbSum, dtbByDateXLot1, drOutlierXLot1.Length)
                    End If
                Catch exXLot1 As Exception

                End Try
                Try
                    Dim strFilterByDateXLot2 As String = "[MEW.Delta]>=-1.5 AND [MEW.Delta]<=1.5 AND [MEW.Delta] IS NOT NULL"
                    strFilterByDateXLot2 = strFilterByDateXLot2 & " AND SUBSTRING([Lot.MT],LEN([Lot.MT])-1,2)>=75"
                    dtbData.DefaultView.RowFilter = strFilterByDateXLot2
                    Dim dtbByDateXLot2 As DataTable = dtbData.DefaultView.ToTable
                    Dim strOutlier As String = "([MEW.Delta]>1.5 OR [MEW.Delta]<-1.5 OR [MEW.Delta] IS NULL) AND SUBSTRING([Lot.MT],LEN([Lot.MT])-1,2)>=75"
                    Dim drOutlierXLot2() As DataRow = dtbData.Select(strOutlier)
                    If dtbByDateXLot2.Rows.Count > 0 Then
                        Dim strTesterBWT As String = "BWTXLot2"
                        dtbSum.Rows.Add()
                        dtbSum.Rows(dtbSum.Rows.Count - 1).Item("MTProduct") = "SummaryByDate.XLot2"
                        AssignValueToSummaryTable(strSearchBy, "SummaryByDate.XLot2", "XLot2", strTesterBWT, dtbSum, dtbByDateXLot2, drOutlierXLot2.Length)
                    End If
                Catch exXLot2 As Exception

                End Try
            End If
            GetSummaryCorrelationByLot = dtbSum
        Else
            GetSummaryCorrelationByLot = Nothing
        End If

    End Function

    Private Sub AssignValueToSummaryTable(ByVal strSearchBy As String, ByVal strSearch As String, ByVal strLot As String, ByVal strTesterBWT As String, ByVal dtbSummary As DataTable, ByVal dtbData As DataTable, ByVal nOutlier As Integer)
        Dim clsAnalyze As New CDataAnalyzer
        Dim sLinear As CDataAnalyzer.SLinearParameter = clsAnalyze.CalculateLinearRegression("MEW6T.BWT", "WEW.MT", dtbData)
        Dim sLinear2 As CDataAnalyzer.SLinearParameter = clsAnalyze.CalculateLinearRegression("WEW.MT", "MEW6T.BWT", dtbData)

        dtbSummary.Rows.Add()
        Dim nIndex As Integer = dtbSummary.Rows.Count - 1
        dtbSummary.Rows(nIndex).Item("MTProduct") = dtbData.Rows(0).Item("MTProduct")
        dtbSummary.Rows(nIndex).Item("TotalHGA") = dtbData.Rows.Count + nOutlier
        dtbSummary.Rows(nIndex).Item("OutlierHead") = nOutlier
        dtbSummary.Rows(nIndex).Item("TestTime.MT") = dtbData.Compute("MAX([TestTime.MT])", "")
        dtbSummary.Rows(nIndex).Item("TestTime.BWT") = dtbData.Compute("MAX([TestTime.BWT])", "")
        dtbSummary.Rows(nIndex).Item("Spec.MT") = dtbData.Rows(0).Item("Spec.MT")
        dtbSummary.Rows(nIndex).Item("Tester.BWT") = strTesterBWT
        dtbSummary.Rows(nIndex).Item(strSearchBy) = strSearch
        dtbSummary.Rows(nIndex).Item("Lot.MT") = strLot
        dtbSummary.Rows(nIndex).Item("LotType") = dtbData.Rows(0).Item("LotType")
        dtbSummary.Rows(nIndex).Item("Device") = dtbData.Rows(0).Item("Device")
        dtbSummary.Rows(nIndex).Item("SliderSite") = dtbData.Rows(0).Item("SliderSite")
        dtbSummary.Rows(nIndex).Item("STD_AvgDiff") = dtbData.Rows(0).Item("STD_AvgDiff")
        dtbSummary.Rows(nIndex).Item("STD_Slope") = dtbData.Rows(0).Item("STD_Slope")
        dtbSummary.Rows(nIndex).Item("STD_STDVDiff") = dtbData.Rows(0).Item("STD_STDVDiff")
        dtbSummary.Rows(nIndex).Item("STD_RSQ") = dtbData.Rows(0).Item("STD_RSQ")
        dtbSummary.Rows(nIndex).Item("MEW6T.BWT") = dtbData.Compute("AVG([MEW6T.BWT])", "")
        dtbSummary.Rows(nIndex).Item("WEW.MT") = dtbData.Compute("AVG([WEW.MT])", "")
        dtbSummary.Rows(nIndex).Item("MEW6T.BWT.Stdev") = dtbData.Compute("STDEV([MEW6T.BWT])", "")
        dtbSummary.Rows(nIndex).Item("WEW.MT.Stdev") = dtbData.Compute("STDEV([WEW.MT])", "")
        dtbSummary.Rows(nIndex).Item("AVG.CF.MT") = dtbData.Compute("AVG([WEW.CF])", "")
        dtbSummary.Rows(nIndex).Item("Stdev.CF.MT") = dtbData.Compute("STDEV([WEW.CF])", "")
        dtbSummary.Rows(nIndex).Item("AVG.Media.MT") = dtbData.Compute("AVG([WEW.Media])", "")
        dtbSummary.Rows(nIndex).Item("Slope.BWT") = sLinear.dblSlope
        dtbSummary.Rows(nIndex).Item("Intercept.BWT") = sLinear.dblIntercept
        dtbSummary.Rows(nIndex).Item("RSQ.BWT.MT") = sLinear.dblRSqr
        dtbSummary.Rows(nIndex).Item("Slope.MT") = sLinear2.dblSlope
        dtbSummary.Rows(nIndex).Item("Intercept.MT") = sLinear2.dblIntercept
    End Sub

    Public Function GetBWTReferenceByWafer(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        GetBWTReferenceByWafer = Nothing
        If strProduct.Contains("_SDET") Then
            Dim strSearchBy As String = dtbSearchBy.TableName

            Dim strProductMain As String = strProduct.Split("_")(0)
            Dim strProductXLot As String = strProductMain & "_XLOT_DCT_SDET"
            Dim strProductDCT As String = strProductMain & "_DCT"
            Dim strProductBWT As String = "BWT_BWT"

            Dim clsRTTCParam As New CParameterRTTCMapping(m_mySqlConn)
            Dim dtbProduct As DataTable = clsRTTCParam.GetProductList(enuProductType.enuProductAll)
            If dtbProduct.Select("Product='" & strProductXLot & "'").Length > 0 And dtbProduct.Select("Product='" & strProductDCT & "'").Length > 0 And dtbProduct.Select("Product='" & strProductBWT & "'").Length > 0 Then
                Dim strSQL As String = "SELECT "
                strSQL = strSQL & "MAX(A.Update_time) EndLotTime,"
                strSQL = strSQL & "LEFT(A.Lot,4) Wafer,"
                strSQL = strSQL & "SUM(A.TotalHGA) ""TotalHGAs.SDET"","
                strSQL = strSQL & "(SELECT SUM(M.TotalHGA) FROM db_" & strProductDCT & ".tabmean_avg M "
                strSQL = strSQL & "LEFT JOIN db_" & strProductDCT & ".tabmean_n N USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(A.Lot,4)=LEFT(M.Lot,4)) ""TotalHGAs.DCT"","
                If strSearchBy <> "Wafer" And strSearchBy <> "OptionIndex" Then strSQL = strSQL & "A." & strSearchBy & ","
                strSQL = strSQL & "SUM(A.WEW)/SUM(B.WEW) ""WEW.SDET"","
                strSQL = strSQL & "(SELECT SUM(M.WEW)/SUM(N.WEW) FROM db_" & strProductDCT & ".tabmean_avg M "
                strSQL = strSQL & "LEFT JOIN db_" & strProductDCT & ".tabmean_n N USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(A.Lot,4)=LEFT(M.Lot,4)) ""WEW.DCT"","
                strSQL = strSQL & "(SELECT SUM(C.WEW)/SUM(D.WEW) FROM db_" & strProductXLot & ".tabmean_avg C "
                strSQL = strSQL & "LEFT JOIN db_" & strProductXLot & ".tabmean_n D USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(C.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(C.Lot,LENGTH(C.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(C.Lot,LENGTH(C.Lot)-1) AS SIGNED)>0) ""WEW.XLot1"","
                strSQL = strSQL & "(SELECT SUM(E.WEW)/SUM(F.WEW) FROM db_" & strProductXLot & ".tabmean_avg E "
                strSQL = strSQL & "LEFT JOIN db_" & strProductXLot & ".tabmean_n F USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(E.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(E.Lot,LENGTH(E.Lot)-1) AS SIGNED)>=75) ""WEW.XLot2"","
                strSQL = strSQL & "(SELECT SUM(G.MEW6T)/SUM(H.MEW6T) FROM db_" & strProductBWT & ".tabmean_avg G "
                strSQL = strSQL & "LEFT JOIN db_" & strProductBWT & ".tabmean_n H USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(G.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(G.Lot,LENGTH(G.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(G.Lot,LENGTH(G.Lot)-1) AS SIGNED)>0) ""MEW.BWT.XLot1"","
                strSQL = strSQL & "(SELECT SUM(J.MEW6T)/SUM(K.MEW6T) FROM db_" & strProductBWT & ".tabmean_avg J "
                strSQL = strSQL & "LEFT JOIN db_" & strProductBWT & ".tabmean_n K USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE LEFT(J.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(J.Lot,LENGTH(J.Lot)-1) AS SIGNED)>=75) ""MEW.BWT.XLot2"" "
                strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_avg A "
                strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabmean_n B USING(tester,spec,lot,shoe) "
                strSQL = strSQL & "WHERE "
                strSQL = strSQL & "A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
                strSQL = strSQL & "AND ("
                For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                    Dim strSearchValue As String = dtbSearchBy.Rows(nSearch).Item(strSearchBy)
                    If strSearchBy = "Wafer" Then
                        strSQL = strSQL & "A.Lot LIKE '" & strSearchValue & "%' OR "
                    ElseIf strSearchBy = "OptionIndex" Then
                        If strSearchValue = "0" Then
                            strSQL = strSQL & "RIGHT(A.Spec,1)='A' OR "
                        Else
                            strSQL = strSQL & "RIGHT(A.Spec,1)='B' OR "
                        End If
                    Else
                        strSQL = strSQL & "A." & strSearchBy & "='" & strSearchValue & "' OR "
                    End If
                Next nSearch
                If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 3) & ") "
                strSQL = strSQL & "GROUP BY LEFT(A.Lot,4) "
                strSQL = strSQL & "ORDER BY MAX(A.Update_Time);"
                Dim clsMySql As New CMySQL
                Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
                Dim strSelect As String = "MEW.BWT.XLot2 IS NULL"
                Dim drSelect() As DataRow = dtbData.Select(strSelect)
                For nRemove As Integer = 0 To drSelect.Length - 1
                    dtbData.Rows.Remove(drSelect(nRemove))
                Next
                dtbData.Columns.Add("TotalHGAs", Type.GetType("System.Int32"), "ISNULL(TotalHGAs.SDET,0)+ISNULL(TotalHGAs.DCT,0)")
                dtbData.Columns.Add("WEW.Sister", Type.GetType("System.Double"), "(ISNULL(TotalHGAs.SDET*WEW.SDET,0)+ISNULL(TotalHGAs.DCT*WEW.DCT,0))/(ISNULL(TotalHGAs.SDET,0)+ISNULL(TotalHGAs.DCT,0))")

                dtbData.Columns("TotalHGAs").SetOrdinal(dtbData.Columns("TotalHGAs.DCT").Ordinal + 1)
                dtbData.Columns("WEW.Sister").SetOrdinal(dtbData.Columns("WEW.DCT").Ordinal + 1)
                GetBWTReferenceByWafer = dtbData
            End If
        ElseIf strProduct.Contains("BWT_BWT") Then
            Dim strSearchBy As String = dtbSearchBy.TableName
            Dim dtbWaferRef As New DataTable

            Dim clsRTTCParam As New CParameterRTTCMapping(m_mySqlConn)
            Dim dtbProduct As DataTable = clsRTTCParam.GetProductList(enuProductType.enuProductAll)
            For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
                Dim strProductMain As String = dtbProduct.Rows(nProduct).Item("Product").Split("_")(0)
                Dim strProductXLot As String = strProductMain & "_XLOT_DCT_SDET"
                Dim strProductDCT As String = strProductMain & "_DCT"
                Dim strProductBWT As String = "BWT_BWT"
                If dtbProduct.Select("Product='" & strProductXLot & "'").Length > 0 And dtbProduct.Select("Product='" & strProductDCT & "'").Length > 0 And dtbProduct.Select("Product='" & strProductBWT & "'").Length > 0 Then
                    Dim strSQL As String = "SELECT '" & dtbProduct.Rows(nProduct).Item("Product") & "' MTProduct,"
                    strSQL = strSQL & "MAX(A.Update_time) EndLotTime,"
                    strSQL = strSQL & "LEFT(A.Lot,4) Wafer,"
                    strSQL = strSQL & "SUM(A.TotalHGA) ""TotalHGAs.SDET"","
                    strSQL = strSQL & "(SELECT SUM(M.TotalHGA) FROM db_" & strProductDCT & ".tabmean_avg M "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductDCT & ".tabmean_n N USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(A.Lot,4)=LEFT(M.Lot,4)) ""TotalHGAs.DCT"","
                    If strSearchBy <> "Wafer" And strSearchBy <> "OptionIndex" Then strSQL = strSQL & "A." & strSearchBy & ","
                    strSQL = strSQL & "SUM(A.WEW)/SUM(B.WEW) ""WEW.SDET"","
                    strSQL = strSQL & "(SELECT SUM(M.WEW)/SUM(N.WEW) FROM db_" & strProductDCT & ".tabmean_avg M "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductDCT & ".tabmean_n N USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(A.Lot,4)=LEFT(M.Lot,4)) ""WEW.DCT"","
                    strSQL = strSQL & "(SELECT SUM(C.WEW)/SUM(D.WEW) FROM db_" & strProductXLot & ".tabmean_avg C "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductXLot & ".tabmean_n D USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(C.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(C.Lot,LENGTH(C.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(C.Lot,LENGTH(C.Lot)-1) AS SIGNED)>0) ""WEW.XLot1"","
                    strSQL = strSQL & "(SELECT SUM(E.WEW)/SUM(F.WEW) FROM db_" & strProductXLot & ".tabmean_avg E "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductXLot & ".tabmean_n F USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(E.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(E.Lot,LENGTH(E.Lot)-1) AS SIGNED)>=75) ""WEW.XLot2"","
                    strSQL = strSQL & "(SELECT SUM(G.MEW6T)/SUM(H.MEW6T) FROM db_" & strProductBWT & ".tabmean_avg G "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductBWT & ".tabmean_n H USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(G.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(G.Lot,LENGTH(G.Lot)-1) AS SIGNED)<75 AND CAST(SUBSTR(G.Lot,LENGTH(G.Lot)-1) AS SIGNED)>0) ""MEW.BWT.XLot1"","
                    strSQL = strSQL & "(SELECT SUM(J.MEW6T)/SUM(K.MEW6T) FROM db_" & strProductBWT & ".tabmean_avg J "
                    strSQL = strSQL & "LEFT JOIN db_" & strProductBWT & ".tabmean_n K USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE LEFT(J.Lot,4)=LEFT(A.Lot,4) AND CAST(SUBSTR(J.Lot,LENGTH(J.Lot)-1) AS SIGNED)>=75) ""MEW.BWT.XLot2"" "
                    strSQL = strSQL & "FROM db_BWT_BWT.tabmean_avg A "
                    strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabmean_n B USING(tester,spec,lot,shoe) "
                    strSQL = strSQL & "WHERE "
                    strSQL = strSQL & "A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
                    strSQL = strSQL & "AND ("
                    For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
                        Dim strSearchValue As String = dtbSearchBy.Rows(nSearch).Item(strSearchBy)
                        If strSearchBy = "Wafer" Then
                            strSQL = strSQL & "A.Lot LIKE '" & strSearchValue & "%' OR "
                        ElseIf strSearchBy = "OptionIndex" Then
                            If strSearchValue = "0" Then
                                strSQL = strSQL & "RIGHT(A.Spec,1)='A' OR "
                            Else
                                strSQL = strSQL & "RIGHT(A.Spec,1)='B' OR "
                            End If
                        Else
                            strSQL = strSQL & "A." & strSearchBy & "='" & strSearchValue & "' OR "
                        End If
                    Next nSearch
                    If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 3) & ") "
                    strSQL = strSQL & "GROUP BY LEFT(A.Lot,4) "
                    strSQL = strSQL & "ORDER BY MAX(A.Update_Time);"
                    Dim clsMySql As New CMySQL
                    Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
                    Dim strSelect As String = "MEW.BWT.XLot2 IS NULL"
                    Dim drSelect() As DataRow = dtbData.Select(strSelect)
                    For nRemove As Integer = 0 To drSelect.Length - 1
                        dtbData.Rows.Remove(drSelect(nRemove))
                    Next
                    dtbData.Columns.Add("TotalHGAs", Type.GetType("System.Int32"), "ISNULL(TotalHGAs.SDET,0)+ISNULL(TotalHGAs.DCT,0)")
                    dtbData.Columns.Add("WEW.Sister", Type.GetType("System.Double"), "(ISNULL(TotalHGAs.SDET*WEW.SDET,0)+ISNULL(TotalHGAs.DCT*WEW.DCT,0))/(ISNULL(TotalHGAs.SDET,0)+ISNULL(TotalHGAs.DCT,0))")

                    dtbData.Columns("TotalHGAs").SetOrdinal(dtbData.Columns("TotalHGAs.DCT").Ordinal + 1)
                    dtbData.Columns("WEW.Sister").SetOrdinal(dtbData.Columns("WEW.DCT").Ordinal + 1)
                    dtbWaferRef.Merge(dtbData)
                End If
            Next nProduct
            GetBWTReferenceByWafer = dtbWaferRef
        End If
    End Function

    'Public Function GetSummaryCorrelationByLotV2(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal bAddSummaryByDate As Boolean) As DataTable


    '    Dim strSQL As String = ""
    '    strSQL = "SELECT "
    '    If InStr(strProduct, "BWT") = 1 Then
    '        Dim dtbData As New DataTable
    '        Dim clsProduct As New CParameterRTTCMapping(m_mySqlConn)
    '        Dim dtbXLotProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductXLot)
    '        dtbXLotProduct.Merge(clsProduct.GetProductList(enuProductType.enuProductNPL))

    '    Else
    '        strSQL = "SELECT * FROM db_" & strProduct & ".tabdetail_header A "
    '        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
    '        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd C USING(tag_id) "
    '        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabfactor_cfadd E USING(tag_id),"
    '        strSQL = strSQL & "db_BWT_BWT.tabdetail_header F "
    '        strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabfactor_value G USING(tag_id) "
    '        strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabfactor_cfadd H USING(tag_id) "
    '        strSQL = strSQL & "LEFT JOIN db_BWT_BWT.tabfactor_cfadd I USING(tag_id) "

    '        Dim strSearchBy As String = dtbSearchBy.TableName
    '        strSQL = strSQL & "WHERE (A.Test_Time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "'  AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
    '        strSQL = strSQL & "AND ("
    '        For nTester As Integer = 0 To dtbSearchBy.Rows.Count - 1
    '            If nTester <> dtbSearchBy.Rows.Count - 1 Then
    '                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "' OR "
    '            Else
    '                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearchBy.Rows(nTester).Item(strSearchBy) & "') "
    '            End If
    '        Next nTester
    '        strSQL = strSQL & "AND A.Lot=F.Lot AND A.hga_sn=F.hga_sn "
    '        strSQL = Left(strSQL, Len(strSQL) - 1) & " "
    '    End If
    'End Function

    Private Function GetSliderFabSection(ByVal strHGA_SN As String) As String
        GetSliderFabSection = ""
        strHGA_SN = strHGA_SN.ToUpper
        Dim strSide As String = strHGA_SN.Substring(4, 1)
        Dim strBarNo As String = strHGA_SN.Substring(5, 2)

        Dim strSection As String = ""
        Dim drSection As DataRow = m_dtbSection.Rows.Find(strSide)
        If drSection IsNot Nothing Then
            strSection = drSection.Item("Section")
        End If

        Dim strFirstDigit As String = strBarNo(0)
        Dim strSecondDigit As String = strBarNo(1)

        Dim nNumber As Integer = 0

        If IsNumeric(strFirstDigit) Then
            nNumber = strFirstDigit
        Else
            nNumber = Asc(strFirstDigit) - Asc("A") + 10
        End If

        Dim nRealBar As Integer = nNumber * 10 + CInt(strSecondDigit)
        Dim nDivRealBar As Integer = nRealBar \ 51

        Dim drSFSection As DataRow = m_dtbSldSection.Rows.Find(New String() {strSection, nDivRealBar})
        If drSFSection IsNot Nothing Then
            GetSliderFabSection = drSFSection.Item("SF_Section")
        End If
    End Function

    Private Function InitSldSectionTable() As DataTable
        Dim dtbSldSectionTable As New DataTable
        Dim dcPrimary(1) As DataColumn
        dcPrimary(0) = dtbSldSectionTable.Columns.Add("Side")
        dcPrimary(1) = dtbSldSectionTable.Columns.Add("ModeRealBar")
        dtbSldSectionTable.PrimaryKey = dcPrimary
        dtbSldSectionTable.Columns.Add("SF_Section")
        dtbSldSectionTable.Rows.Add("1", "0", "1_B")
        dtbSldSectionTable.Rows.Add("1", "1", "1_A")
        dtbSldSectionTable.Rows.Add("2", "0", "3_B")
        dtbSldSectionTable.Rows.Add("2", "1", "3_A")
        dtbSldSectionTable.Rows.Add("3", "0", "5_B")
        dtbSldSectionTable.Rows.Add("3", "1", "5_A")
        dtbSldSectionTable.Rows.Add("4", "0", "7_B")
        dtbSldSectionTable.Rows.Add("4", "1", "7_A")
        dtbSldSectionTable.Rows.Add("5", "0", "9_B")
        dtbSldSectionTable.Rows.Add("5", "1", "9_A")
        dtbSldSectionTable.Rows.Add("6", "0", "A_F")
        dtbSldSectionTable.Rows.Add("6", "1", "A_E")
        dtbSldSectionTable.Rows.Add("6", "2", "A_D")
        dtbSldSectionTable.Rows.Add("6", "3", "A_C")
        dtbSldSectionTable.Rows.Add("6", "4", "A_B")
        dtbSldSectionTable.Rows.Add("6", "5", "A_A")
        dtbSldSectionTable.Rows.Add("7", "0", "C_F")
        dtbSldSectionTable.Rows.Add("7", "1", "C_E")
        dtbSldSectionTable.Rows.Add("7", "2", "C_D")
        dtbSldSectionTable.Rows.Add("7", "3", "C_C")
        dtbSldSectionTable.Rows.Add("7", "4", "C_B")
        dtbSldSectionTable.Rows.Add("7", "5", "C_A")
        dtbSldSectionTable.Rows.Add("8", "0", "E_F")
        dtbSldSectionTable.Rows.Add("8", "1", "E_E")
        dtbSldSectionTable.Rows.Add("8", "2", "E_D")
        dtbSldSectionTable.Rows.Add("8", "3", "E_C")
        dtbSldSectionTable.Rows.Add("8", "4", "E_B")
        dtbSldSectionTable.Rows.Add("8", "5", "E_A")
        dtbSldSectionTable.Rows.Add("9", "0", "G_F")
        dtbSldSectionTable.Rows.Add("9", "1", "G_E")
        dtbSldSectionTable.Rows.Add("9", "2", "G_D")
        dtbSldSectionTable.Rows.Add("9", "3", "G_C")
        dtbSldSectionTable.Rows.Add("9", "4", "G_B")
        dtbSldSectionTable.Rows.Add("9", "5", "G_A")
        dtbSldSectionTable.Rows.Add("0", "0", "J_B")
        dtbSldSectionTable.Rows.Add("0", "1", "J_A")
        dtbSldSectionTable.Rows.Add("A", "0", "L_B")
        dtbSldSectionTable.Rows.Add("A", "1", "L_A")
        dtbSldSectionTable.Rows.Add("B", "0", "N_B")
        dtbSldSectionTable.Rows.Add("B", "1", "N_A")
        dtbSldSectionTable.Rows.Add("C", "0", "Q_B")
        dtbSldSectionTable.Rows.Add("C", "1", "Q_A")
        dtbSldSectionTable.Rows.Add("D", "0", "S_B")
        dtbSldSectionTable.Rows.Add("D", "1", "S_A")
        InitSldSectionTable = dtbSldSectionTable

    End Function

    Private Function InitSectionTable() As DataTable
        Dim dtbSectionTable As New DataTable
        Dim dcPrimary(0) As DataColumn
        dcPrimary(0) = dtbSectionTable.Columns.Add("Side")
        dtbSectionTable.Columns.Add("Section")

        dtbSectionTable.PrimaryKey = dcPrimary
        dtbSectionTable.Rows.Add("1", "1")
        dtbSectionTable.Rows.Add("2", "1")
        dtbSectionTable.Rows.Add("3", "2")
        dtbSectionTable.Rows.Add("4", "2")
        dtbSectionTable.Rows.Add("5", "3")
        dtbSectionTable.Rows.Add("6", "3")
        dtbSectionTable.Rows.Add("7", "4")
        dtbSectionTable.Rows.Add("8", "4")
        dtbSectionTable.Rows.Add("9", "5")
        dtbSectionTable.Rows.Add("0", "5")
        dtbSectionTable.Rows.Add("A", "6")
        dtbSectionTable.Rows.Add("B", "6")
        dtbSectionTable.Rows.Add("C", "7")
        dtbSectionTable.Rows.Add("D", "7")
        dtbSectionTable.Rows.Add("E", "8")
        dtbSectionTable.Rows.Add("F", "8")
        dtbSectionTable.Rows.Add("G", "9")
        dtbSectionTable.Rows.Add("H", "9")
        dtbSectionTable.Rows.Add("J", "0")
        dtbSectionTable.Rows.Add("K", "0")
        dtbSectionTable.Rows.Add("L", "A")
        dtbSectionTable.Rows.Add("M", "A")
        dtbSectionTable.Rows.Add("N", "B")
        dtbSectionTable.Rows.Add("P", "B")
        dtbSectionTable.Rows.Add("Q", "C")
        dtbSectionTable.Rows.Add("R", "C")
        dtbSectionTable.Rows.Add("S", "D")
        dtbSectionTable.Rows.Add("T", "D")
        InitSectionTable = dtbSectionTable
    End Function

End Class

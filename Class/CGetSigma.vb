Imports MySql.Data.MySqlClient

Public Class CGetSigma

    Private m_MySqlConn As MySqlConnection

    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetSigma2Lot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim strSearchBy As String = dtbSearch.TableName
        Dim dtbSigmaData As DataTable = Nothing
        If strSearchBy.ToUpper = "TESTER" Then
            dtbSigmaData = GetSigma2LotByTester(strProduct, dtbSearch, dtbParam, dtStart, dtEnd)
            dtbSigmaData = CalculateDelta(strSearchBy, dtbSigmaData, dtbParam)
        ElseIf strSearchBy.ToUpper = "LOT" Then
            dtbSigmaData = GetSigma2LotByLot(strProduct, dtbSearch, dtbParam, dtStart, dtEnd)
        ElseIf strSearchBy.ToUpper = "SPEC" Then
            dtbSigmaData = GetSigma2LotBySpec(strProduct, dtbSearch, dtbParam, dtStart, dtEnd)
        End If
        GetSigma2Lot = dtbSigmaData
    End Function

    Private Function GetSigma2LotByTester(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim clsParam As New CParameterRTTCMapping(m_MySqlConn)
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

        Dim strBinA As String = "Defect" & nIDPassBin1
        Dim strBinAll As String = "Defect" & nIDPassBin1 & "+Defect" & nIDPassBin2 & "+Defect" & nIDPassBin3 & "+Defect" & nIDPassBin4

        Dim strSearchBy As String = dtbSearch.TableName

        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "A.Update_time Date_time,"
        strSQL = strSQL & "A.Tester,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalHGA.Lot"","
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot) ""TotalHGA.Tester"","
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='1') ""TotalHGA.Shoe1"","
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='2') ""TotalHGA.Shoe2"","

        strSQL = strSQL & "(SELECT SUM(" & strBinA & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalPass.Lot.BinA"","
        strSQL = strSQL & "(SELECT SUM(" & strBinA & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='1') ""TotalPass.Tester.Shoe1.BinA"","
        strSQL = strSQL & "(SELECT SUM(" & strBinA & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='2') ""TotalPass.Tester.Shoe2.BinA"","

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam_rttc As String = dtbParam.Rows(nParam).Item("Param_rttc")
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")

            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigmabylot WHERE Spec=A.Spec AND Lot=A.Lot AND BinType=False) """ & strParaDisplay & ".Stdev.Lot.BinA"","
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigma WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='1' AND BinType=False) """ & strParaDisplay & ".Stdev.Tester.Shoe1.BinA"","
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigma WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='2' AND BinType=False) """ & strParaDisplay & ".Stdev.Tester.Shoe2.BinA"","

        Next nParam

        strSQL = strSQL & "(SELECT SUM(" & strBinAll & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalPass.Lot.BinAll"","
        strSQL = strSQL & "(SELECT SUM(" & strBinAll & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='1') ""TotalPass.Tester.Shoe1.BinAll"","
        strSQL = strSQL & "(SELECT SUM(" & strBinAll & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='2') ""TotalPass.Tester.Shoe2.BinAll"","
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam_rttc As String = dtbParam.Rows(nParam).Item("Param_rttc")
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigmabylot WHERE Spec=A.Spec AND Lot=A.Lot AND BinType=True) """ & strParaDisplay & ".Stdev.Lot.BinAll"","
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigma WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='1' AND BinType=True) """ & strParaDisplay & ".Stdev.Tester.Shoe1.BinAll"","
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigma WHERE Tester=A.Tester AND Spec=A.Spec AND Lot=A.Lot AND Shoe='2' AND BinType=True) """ & strParaDisplay & ".Stdev.Tester.Shoe2.BinAll"","

        Next nParam

        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_sigma A "
        strSQL = strSQL & "WHERE A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        strSQL = strSQL & "AND (LEFT(A.Spec,1)='B' OR LEFT(A.Spec,1)='F' OR LEFT(A.Spec,1)='R' OR LEFT(A.Spec,1)='T') "
        strSQL = strSQL & "GROUP BY Tester,Spec,Lot "
        strSQL = strSQL & "ORDER BY Tester,Spec,Lot;"

        Dim clsMySql As New CMySQL
        Dim dtbSigmaBin As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
        GetSigma2LotByTester = dtbSigmaBin
    End Function

    Private Function GetSigma2LotByLot(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim clsParam As New CParameterRTTCMapping(m_MySqlConn)
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

        Dim strBinA As String = "Defect" & nIDPassBin1
        Dim strBinAll As String = "Defect" & nIDPassBin1 & "+Defect" & nIDPassBin2 & "+Defect" & nIDPassBin3 & "+Defect" & nIDPassBin4

        Dim strSearchBy As String = dtbSearch.TableName

        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "A.Update_time Date_time,"
        strSQL = strSQL & "A.Lot,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalHGA.Lot"","
        strSQL = strSQL & "(SELECT SUM(" & strBinA & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalPass.Lot.BinA"","

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam_rttc As String = dtbParam.Rows(nParam).Item("Param_rttc")
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")

            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigmabylot WHERE Spec=A.Spec AND Lot=A.Lot AND BinType=False) """ & strParaDisplay & ".Stdev.Lot.BinA"","
        Next nParam

        strSQL = strSQL & "(SELECT SUM(" & strBinAll & ") FROM db_" & strProduct & ".tabsummary_hgadefect WHERE Spec=A.Spec AND Lot=A.Lot) ""TotalPass.Lot.BinAll"","
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam_rttc As String = dtbParam.Rows(nParam).Item("Param_rttc")
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")
            strSQL = strSQL & "(SELECT " & strParam_rttc & " FROM db_" & strProduct & ".tabmean_sigmabylot WHERE Spec=A.Spec AND Lot=A.Lot AND BinType=True) """ & strParaDisplay & ".Stdev.Lot.BinAll"","
        Next nParam

        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1) & " "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabmean_sigma A "
        strSQL = strSQL & "WHERE A.Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        strSQL = strSQL & "AND (LEFT(A.Spec,1)='B' OR LEFT(A.Spec,1)='F' OR LEFT(A.Spec,1)='R' OR LEFT(A.Spec,1)='T') "
        strSQL = strSQL & "GROUP BY Spec,Lot "
        strSQL = strSQL & "ORDER BY Spec,Lot;"

        Dim clsMySql As New CMySQL
        Dim dtbSigmaBin As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
        GetSigma2LotByLot = dtbSigmaBin
    End Function

    Private Function GetSigma2LotBySpec(ByVal strProduct As String, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As DataTable
        Dim clsParam As New CParameterRTTCMapping(m_MySqlConn)
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

        Dim strBinA As String = "Defect" & nIDPassBin1
        Dim strBinAll As String = "Defect" & nIDPassBin1 & "+Defect" & nIDPassBin2 & "+Defect" & nIDPassBin3 & "+Defect" & nIDPassBin4

        Dim strSearchBy As String = dtbSearch.TableName

        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "MAX(A.Test_time) Date_time,"
        strSQL = strSQL & "A.Spec,"
        strSQL = strSQL & "(SELECT SUM(TotalHGA) FROM db_" & strProduct & ".tabsummary_hgadefect "
        strSQL = strSQL & "WHERE Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' AND Spec=A.Spec) ""TotalHGA.Spec"","
        strSQL = strSQL & "(SELECT SUM(" & strBinA & ") FROM db_" & strProduct & ".tabsummary_hgadefect "
        strSQL = strSQL & "WHERE Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' AND Spec=A.Spec) ""TotalPass.Spec.BinA"","
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam_rttc As String = dtbParam.Rows(nParam).Item("Param_rttc")
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")
            strSQL = strSQL & "STD(" & strParam_rttc & ") """ & strParaDisplay & ".Stdev.Spec.BinA"","
        Next nParam
        strSQL = strSQL & "(SELECT SUM(" & strBinAll & ") FROM db_" & strProduct & ".tabsummary_hgadefect "
        strSQL = strSQL & "WHERE Update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' AND Spec=A.Spec) ""TotalPass.Spec.BinAll"" "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabdetail_header A LEFT JOIN db_" & strProduct & ".tabfactor_value B USING(tag_id) "
        strSQL = strSQL & "WHERE A.test_time_bigint BETWEEN '" & Format(dtStart, "yyyyMMddHHmmss") & "' AND '" & Format(dtEnd, "yyyyMMddHHmmss") & "' "
        strSQL = strSQL & "AND ("
        For nSearch As Integer = 0 To dtbSearch.Rows.Count - 1
            If nSearch <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "A." & strSearchBy & "='" & dtbSearch.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        strSQL = strSQL & "AND (LEFT(A.Spec,1)='B' OR LEFT(A.Spec,1)='F' OR LEFT(A.Spec,1)='R' OR LEFT(A.Spec,1)='T') "
        Dim strSqlBinA As String = strSQL & "AND A.GradeName LIKE 'PASS_BIN1%' GROUP BY SPEC ORDER BY Spec;"
        Dim strSqlBinAll As String = strSQL & "AND A.GradeName LIKE 'PASS_BIN%' GROUP BY SPEC ORDER BY Spec;"

        Dim clsMySql As New CMySQL
        Dim dtsSigma As DataSet = clsMySql.CommandMySqlDataset(strSqlBinA & strSqlBinAll, m_MySqlConn)
        Dim dtbSigmaBinA As DataTable = dtsSigma.Tables(0)
        Dim dtbSigmaBinAll As DataTable = dtsSigma.Tables(1)

        Dim dtbDataAllBin As DataTable = dtbSigmaBinA.DefaultView.ToTable

        For nData As Integer = 0 To dtbDataAllBin.Rows.Count - 1
            Dim drRow As DataRow = dtbDataAllBin.Rows(nData)
            Dim strSpec As String = drRow.Item("Spec")
            For nParam As Integer = 0 To dtbParam.Rows.Count - 1
                Dim strParamDisplay = dtbParam.Rows(nParam).Item("param_display")
                Dim strColBinA As String = strParamDisplay & ".Stdev.Spec.BinA"
                Dim strColBinAll As String = strParamDisplay & ".Stdev.Spec.BinAll"
                If dtbDataAllBin.Columns(strColBinAll) Is Nothing Then dtbDataAllBin.Columns.Add(strColBinAll, Type.GetType("System.Double"))
                Dim drSelect() As DataRow = dtbSigmaBinAll.Select("Spec='" & strSpec & "'")
                dtbDataAllBin.Rows(nData).Item(strColBinAll) = drSelect(0).Item(strColBinA)
            Next nParam
        Next

        GetSigma2LotBySpec = dtbDataAllBin
    End Function

    Private Function CalculateDelta(ByVal strSearchBy As String, ByVal dtbData As DataTable, ByVal dtbParam As DataTable) As DataTable
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")
            Dim strColShoe1 As String = strParaDisplay & ".Stdev." & strSearchBy & ".Shoe1.BinA"
            Dim strColShoe2 As String = strParaDisplay & ".Stdev." & strSearchBy & ".Shoe2.BinA"
            Dim strColLot As String = strParaDisplay & ".Stdev.Lot.BinA"
            dtbData.Columns.Add("Delta." & strParaDisplay & ".Stdev.Shoe1.BinA", Type.GetType("System.Double"), strColShoe1 & "-" & strColLot)
            dtbData.Columns.Add("Delta." & strParaDisplay & ".Stdev.Shoe2.BinA", Type.GetType("System.Double"), strColShoe2 & "-" & strColLot)
        Next nParam
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParaDisplay As String = dtbParam.Rows(nParam).Item("param_Display")
            Dim strColShoe1 As String = strParaDisplay & ".Stdev." & strSearchBy & ".Shoe1.BinAll"
            Dim strColShoe2 As String = strParaDisplay & ".Stdev." & strSearchBy & ".Shoe2.BinAll"
            Dim strColLot As String = strParaDisplay & ".Stdev.Lot.BinAll"
            dtbData.Columns.Add("Delta." & strParaDisplay & ".Stdev.Shoe1.BinAll", Type.GetType("System.Double"), strColShoe1 & "-" & strColLot)
            dtbData.Columns.Add("Delta." & strParaDisplay & ".Stdev.Shoe2.BinAll", Type.GetType("System.Double"), strColShoe2 & "-" & strColLot)
        Next nParam
        CalculateDelta = dtbData
    End Function

End Class

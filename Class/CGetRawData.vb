Imports MySql.Data.MySqlClient

Public Class CGetRawData
    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

    Public Function GetRawDataByParamByTester(ByVal strProduct As String, ByVal dtbHeader As DataTable, ByVal dtbParam As DataTable, ByVal dtbSearchBy As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtStart As DateTime,
        ByVal dtEnd As DateTime, ByVal bIncludeCF As Boolean, Optional ByVal strShoe As String = "") As DataTable
        Dim strSQL As String

        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        For nHeader As Integer = 0 To dtbHeader.Rows.Count - 1
            Dim strHeader As String = dtbHeader.Rows(nHeader).Item("Header").ToString
            If strHeader.ToUpper = "BARNO" Then
                strSQL = strSQL & "Mid(a.HGA_SN,5,1) SideNo,"
                strSQL = strSQL & "Mid(HGA_SN,6,2) BarNo,"
                strSQL = strSQL & "Right(HGA_SN,1) LocNo,"
            Else
                strSQL = strSQL & "a." & strHeader & ","
            End If
        Next nHeader

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParaName As String = CStr(dtbParam.Rows(nParam).Item("param_rttc"))
            Dim strDisplayName As String = CStr(dtbParam.Rows(nParam).Item("param_display"))
            Dim strMachineCF As String '= dtbParam.Rows(nParam).Item("MachineCF")
            If IsDBNull(dtbParam.Rows(nParam).Item("MachineCF")) Then
                strMachineCF = ""
            Else
                strMachineCF = CStr(dtbParam.Rows(nParam).Item("MachineCF"))
            End If

            strSQL = strSQL & "B." & strParaName & " As """ & strDisplayName & ""","
            If bIncludeCF = True Then
                Dim bCFAdd As Boolean = dtbParam.Rows(nParam).Item("param_add")
                Dim bCFMul As Boolean = dtbParam.Rows(nParam).Item("param_mul")
                If bCFAdd Or strMachineCF <> "" Then
                    strSQL = strSQL & "C." & strParaName & " """ & strDisplayName & ".CFAdd"","
                End If
                If bCFMul = True Or strMachineCF <> "" Then
                    strSQL = strSQL & "D." & strParaName & " """ & strDisplayName & ".CFMul"","
                End If
                If bCFAdd Or bCFMul Then
                    strSQL = strSQL & "(SELECT " & strParaName & " FROM db_" & strProduct & ".tabfactor_media WHERE Media_PN=LEFT(A.MediaSN,3) AND Media_Group=SUBSTRING(A.MediaSN,4,1) AND Media_Serial=SUBSTRING(A.MediaSN,5,3)) """ & strDisplayName & ".CFMedia"","
                End If

            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1)
        'From and Condition
        strSQL = strSQL & " FROM db_" & strProduct & ".tabdetail_header A "
        strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_value B" & " USING (tag_id) "
        If bIncludeCF = True Then
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_cfadd C" & " USING (tag_id) "
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_cfmul D" & " USING (tag_id) "
        End If
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " (A.test_time_bigint between '" & Format(dtStart, "yyyyMMddHHmmss") & "' and '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
        strSQL = strSQL & " AND ("
        Dim strSearchBy As String = dtbSearchBy.TableName

        For nSearch As Integer = 0 To dtbSearchBy.Rows.Count - 1
            If nSearch <> dtbSearchBy.Rows.Count - 1 Then
                strSQL = strSQL & "a." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "' OR "
            Else
                strSQL = strSQL & "a." & strSearchBy & "='" & dtbSearchBy.Rows(nSearch).Item(strSearchBy) & "') "
            End If
        Next nSearch
        If dtbSliderSite IsNot Nothing Then
            If dtbSliderSite.Rows.Count > 0 Then strSQL = strSQL & " AND ("
            For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
                If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                    strSQL = strSQL & "a.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
                Else
                    strSQL = strSQL & "a.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
                End If
            Next nSliderSite
        End If
        If strShoe <> "" Then
            strSQL = strSQL & " AND Shoe ='" & strShoe & "'"
        End If
        strSQL = strSQL & " ORDER BY A.tester,A.test_time,spec,lot "
        strSQL = strSQL & "LIMIT 0,500000"

        Dim clsRawData As New CMySQL
        Dim dtbRawData As DataTable = clsRawData.CommandMySqlDataTable(strSQL, m_mySqlConn)

        GetRawDataByParamByTester = dtbRawData

    End Function

    ''' <summary>
    ''' Get Raw Data by Tester version 2 is using for auto mark delta mrr 
    ''' </summary>
    ''' <param name="strProduct"></param>
    ''' <param name="dtbHeader"></param>
    ''' <param name="dtbParam"></param>
    ''' <param name="dtStart"></param>
    ''' <param name="strTester"></param>
    ''' <param name="strDiskSN"></param>
    ''' <param name="strDiskSurface"></param>
    ''' <returns>datatable</returns>
    Public Function GetRawDataByParamByTesterV2(ByVal strProduct As String, ByVal dtbHeader As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, strTester As String,
        ByVal strDiskSN As String, strDiskSurface As String) As DataTable
        Dim strSQL As String = "SELECT * FROM db_"
        strSQL = strSQL & strProduct & ".tabdetail_header "
        strSQL = strSQL & " WHERE GradeName <> 'FAIL_SCRAP_DELTA_MRR' AND Spec not like 'C%'"
        strSQL = strSQL & " AND test_time_bigint >= '" & Format(dtStart, "yyyyMMddHHmmss") & "'"
        strSQL = strSQL & " AND MediaSN ='" & strDiskSN & "' AND MediaSurface Like '" & strDiskSurface & "%' AND Tester='" & strTester & "'"
        strSQL = strSQL & " ORDER BY test_time"
        strSQL = strSQL & " LIMIT 0,50000 ;"

        Dim clsRawData As New CMySQL
        Dim dtbRawData As DataTable = clsRawData.CommandMySqlDataTable(strSQL, m_mySqlConn)

        GetRawDataByParamByTesterV2 = dtbRawData

    End Function



    ''' <summary>
    ''' Get Raw Data by Tester lot spec shoed
    ''' </summary>
    ''' <param name="strProduct"></param>
    ''' <param name="dtbHeader"></param>
    ''' <param name="dtbParam"></param>
    ''' <param name="strTester"></param>
    ''' <param name="strLot"></param>
    ''' <param name="strSpec"></param>
    ''' <param name="dtStart"></param>
    ''' <param name="dtEnd"></param>
    ''' <param name="bIncludeCF"></param>
    ''' <param name="strShoe"></param>
    ''' <returns></returns>
    Public Function GetRawDataByTesterLotSpecShoe(ByVal strProduct As String, ByVal dtbHeader As DataTable, ByVal dtbParam As DataTable, ByVal strTester As String, ByVal strLot As String, ByVal strSpec As String, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal bIncludeCF As Boolean, Optional ByVal strShoe As String = "") As DataTable
        Dim strSQL As String

        strSQL = "SELECT "
        strSQL = strSQL & "'" & strProduct & "' ProductName,"
        For nHeader As Integer = 0 To dtbHeader.Rows.Count - 1
            Dim strHeader As String = dtbHeader.Rows(nHeader).Item("Header").ToString
            If strHeader.ToUpper = "BARNO" Then
                strSQL = strSQL & "Mid(a.HGA_SN,5,1) SideNo,"
                strSQL = strSQL & "Mid(HGA_SN,6,2) BarNo,"
                strSQL = strSQL & "Right(HGA_SN,1) LocNo,"
            Else
                strSQL = strSQL & "a." & strHeader & ","
            End If
        Next nHeader

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParaName As String = dtbParam.Rows(nParam).Item("param_rttc")
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            strSQL = strSQL & "B." & strParaName & " As """ & strDisplayName & ""","
            If bIncludeCF = True Then
                If dtbParam.Rows(nParam).Item("param_add") = True And dtbParam.Rows(nParam).Item("param_mul") = False Then
                    strSQL = strSQL & "C." & strParaName & " '" & strDisplayName & "_Add',"
                ElseIf dtbParam.Rows(nParam).Item("param_add") = False And dtbParam.Rows(nParam).Item("param_mul") = True Then
                    strSQL = strSQL & "D." & strParaName & " '" & strDisplayName & "_Mul',"
                ElseIf dtbParam.Rows(nParam).Item("param_add") = True And dtbParam.Rows(nParam).Item("param_mul") = True Then
                    strSQL = strSQL & "C." & strParaName & " '" & strDisplayName & "_Add',"
                    strSQL = strSQL & "D." & strParaName & " '" & strDisplayName & "_Mul',"
                End If
            End If
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, Len(strSQL) - 1)
        'From and Condition
        strSQL = strSQL & " FROM db_" & strProduct & ".tabdetail_header A "
        strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_value B" & " USING (tag_id) "
        If bIncludeCF = True Then
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_cfadd C" & " USING (tag_id) "
            strSQL = strSQL & " LEFT JOIN db_" & strProduct & ".tabfactor_cfmul D" & " USING (tag_id) "
        End If
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " (A.test_time_bigint between '" & Format(dtStart, "yyyyMMddHHmmss") & "' and '" & Format(dtEnd, "yyyyMMddHHmmss") & "') "
        strSQL = strSQL & " AND A.Tester='" & strTester & "' "
        strSQL = strSQL & " AND A.Spec='" & strSpec & "' "
        strSQL = strSQL & " AND A.Lot='" & strLot & "' "
        If strShoe <> "" Then
            strSQL = strSQL & " AND Shoe ='" & strShoe & "'"
        End If
        strSQL = strSQL & " ORDER BY A.tester,A.test_time,spec,lot "
        ' strSQL = strSQL & "LIMIT 0,500000"

        Dim clsRawData As New CMySQL
        Dim dtbRawData As DataTable = clsRawData.CommandMySqlDataTable(strSQL, m_mySqlConn)

        GetRawDataByTesterLotSpecShoe = dtbRawData
        'dtbRawData.Dispose()

    End Function




End Class

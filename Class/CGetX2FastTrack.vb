Imports MySql.Data.MySqlClient

Public Class CGetX2FastTrack
    Public Enum eMapMachine
        eV2002 = 0
        eDCT
    End Enum

    Private m_MyConn As MySqlConnection
    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MyConn = MySqlConn
    End Sub

    Public Function GetX2FastTrack(ByVal strMassProduct As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, _
    ByVal dtEnd As DateTime, ByVal eOption As enumSearchOption) As DataTable

        Dim strSQL As String = ""
        Dim strProduct As String = ""
        Dim strFastTrack_V2002 As String
        Dim strFastTrack_DCT As String
        Dim strFastTrack_DCT_SDET As String
        Dim strFastTrack_DUAL_SDET As String


        strProduct = Replace(strMassProduct.ToUpper, "_DCT", "")
        strProduct = Replace(strProduct, "_SDET", "")
        strProduct = Replace(strProduct, "_EH300", "")
        strProduct = Replace(strProduct, "_FASTTRACK", "")
        strProduct = Replace(strProduct, "_NPL", "")
        strProduct = Replace(strProduct, "_XLOT", "")
        strProduct = Replace(strProduct, "_DOE", "")
        strProduct = Replace(strProduct, "_V2002", "")
        strProduct = Replace(strProduct, "_DUAL", "")
        strFastTrack_DCT = strProduct & "_FASTTRACK_DCT"
        strFastTrack_V2002 = strProduct & "_FASTTRACK_V2002"
        strFastTrack_DCT_SDET = strProduct & "_FASTTRACK_DCT_SDET"
        strFastTrack_DUAL_SDET = strProduct & "_FASTTRACK_DUAL_SDET"
        ' If strMassProduct = strFastTrack_DCT Or strMassProduct = strFastTrack_V2002 Or strMassProduct = strFastTrack_DCT_SDET Then
        'Return Nothing
        'End If

        Dim dtbX2FasTrackDCT As DataTable = Nothing
        Dim dtbX2FastTrackV2002 As DataTable = Nothing
        Dim dtbX2FastTrackDCTSDET As DataTable = Nothing
        Dim dtbX2FastTrackDUALSDET As DataTable = Nothing
        'Try
        Dim clsParam As New CParameterRTTCMapping(m_MyConn)
        Dim clsMySQL As New CMySQL
        If clsParam.IsProductExist(strFastTrack_DCT) = True Then
            If eOption = enumSearchOption.eSearchByTester Then
                strSQL = GetSqlStringByTester(strMassProduct, strFastTrack_DCT, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            ElseIf eOption = enumSearchOption.eSearchByLot Then
                strSQL = GetSqlStringByLot(strMassProduct, strFastTrack_DCT, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            Else
                strSQL = GetSqlStringBySpec(strMassProduct, strFastTrack_DCT, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            End If
            dtbX2FasTrackDCT = clsMySQL.CommandMySqlDataTable(strSQL, m_MyConn)
        End If
        If clsParam.IsProductExist(strFastTrack_V2002) = True Then
            If eOption = enumSearchOption.eSearchByTester Then
                strSQL = GetSqlStringByTester(strMassProduct, strFastTrack_V2002, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            ElseIf eOption = enumSearchOption.eSearchByLot Then
                strSQL = GetSqlStringByLot(strMassProduct, strFastTrack_V2002, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            Else
                strSQL = GetSqlStringBySpec(strMassProduct, strFastTrack_V2002, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            End If
            dtbX2FastTrackV2002 = clsMySQL.CommandMySqlDataTable(strSQL, m_MyConn)
        End If
        If clsParam.IsProductExist(strFastTrack_DCT_SDET) = True Then
            If eOption = enumSearchOption.eSearchByTester Then
                strSQL = GetSqlStringByTester(strMassProduct, strFastTrack_DCT_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            ElseIf eOption = enumSearchOption.eSearchByLot Then
                strSQL = GetSqlStringByLot(strMassProduct, strFastTrack_DCT_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            Else
                strSQL = GetSqlStringBySpec(strMassProduct, strFastTrack_DCT_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            End If
            dtbX2FastTrackDCTSDET = clsMySQL.CommandMySqlDataTable(strSQL, m_MyConn)
        End If
        If clsParam.IsProductExist(strFastTrack_DUAL_SDET) = True Then
            If eOption = enumSearchOption.eSearchByTester Then
                strSQL = GetSqlStringByTester(strMassProduct, strFastTrack_DUAL_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            ElseIf eOption = enumSearchOption.eSearchByLot Then
                strSQL = GetSqlStringByLot(strMassProduct, strFastTrack_DUAL_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            Else
                strSQL = GetSqlStringBySpec(strMassProduct, strFastTrack_DUAL_SDET, dtbSearch, dtbSliderSite, dtbParam, dtStart, dtEnd)
            End If
            dtbX2FastTrackDCTSDET = clsMySQL.CommandMySqlDataTable(strSQL, m_MyConn)
        End If
        Dim dtbX2Fastrack As New DataTable
        If Not dtbX2FasTrackDCT Is Nothing Then
            dtbX2Fastrack.Merge(dtbX2FasTrackDCT)
        End If
        If Not dtbX2FastTrackV2002 Is Nothing Then
            dtbX2Fastrack.Merge(dtbX2FastTrackV2002)
        End If
        If Not dtbX2FastTrackDCTSDET Is Nothing Then
            dtbX2Fastrack.Merge(dtbX2FastTrackDCTSDET)
        End If
        If Not dtbX2FastTrackDUALSDET Is Nothing Then
            dtbX2Fastrack.Merge(dtbX2FastTrackDUALSDET)
        End If
        If dtbX2Fastrack.Rows.Count > 0 Then
            For nParam As Integer = 0 To dtbParam.Rows.Count - 1
                Dim strParam As String = dtbParam.Rows(nParam).Item("param_display").ToString
                Dim strParamFT As String = dtbParam.Rows(nParam).Item("paramFasttrack").ToString
                If strparamFT <> "" Then
                    Dim strDelMn As String = strParam & ".Del.Mn"
                    Dim strDelS1 As String = strParam & ".Del.S1"
                    Dim strDelS2 As String = strParam & ".Del.S2"
                    Dim nOrdinal As Integer = dtbX2Fastrack.Columns(strParam & ".S2").Ordinal
                    dtbX2Fastrack.Columns.Add(strDelMn, System.Type.GetType("System.Double"))
                    dtbX2Fastrack.Columns(strDelMn).Expression = "[" & strParam & ".Lot]-[" & strParamFT & ".FT]"
                    dtbX2Fastrack.Columns.Add(strDelS1, System.Type.GetType("System.Double"))
                    dtbX2Fastrack.Columns(strDelS1).Expression = "[" & strParam & ".S1]-[" & strParamFT & ".FT]"
                    dtbX2Fastrack.Columns.Add(strDelS2, System.Type.GetType("System.Double"))
                    dtbX2Fastrack.Columns(strDelS2).Expression = "[" & strParam & ".S2]-[" & strParamFT & ".FT]"
                    dtbX2Fastrack.Columns(strDelMn).SetOrdinal(nOrdinal + 1)
                    dtbX2Fastrack.Columns(strDelS1).SetOrdinal(nOrdinal + 2)
                    dtbX2Fastrack.Columns(strDelS2).SetOrdinal(nOrdinal + 3)
                End If
            Next nParam

        End If
        GetX2FastTrack = dtbX2Fastrack
        'Catch ex As Exception
        'MsgBox(ex.Message)
        'GetX2FastTrack = Nothing
        'End Try

    End Function

    Private Function GetSqlStringByTester(ByVal strMassProduct As String, ByVal strProductFasttrack As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, _
    ByVal dtEnd As DateTime) As String
        Dim strSql As String
        strSql = "SELECT '" & strMassProduct & "' ProductName,"
        strSql = strSql & "'" & strProductFasttrack & "' FastTrack,"
        strSql = strSql & "A.Update_Time,"
        strSql = strSql & "A.Tester,"
        strSql = strSql & "A.Spec,"
        strSql = strSql & "A.Lot,"
        strSql = strSql & "G.Lot FT_Lot,"
        strSql = strSql & "G.Tester FT_Tester,"
        strSql = strSql & "G.Update_time FT_DateTime,"
        strSql = strSql & "G.Spec FT_Spec,"
        strSql = strSql & "A.SliderSite,"
        strSql = strSql & "(SELECT K.WorkID FROM db_" & strProductFasttrack & ".tabdetail_header K WHERE K.Tester=G.Tester AND K.Lot=G.Lot AND K.Spec=G.Spec LIMIT 0,1)" & "FT_WorkID,"
        strSql = strSql & "(SELECT K.WorkID FROM db_" & strMassProduct & ".tabdetail_header K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec LIMIT 0,1)" & "Prime_WorkID,"
        strSql = strSql & "(SELECT K.Assy FROM db_" & strProductFasttrack & ".tabdetail_header K WHERE K.Tester=G.Tester AND K.Lot=G.Lot AND K.Spec=G.Spec LIMIT 0,1)" & "FT_Assy,"
        strSql = strSql & "(SELECT K.Assy FROM db_" & strMassProduct & ".tabdetail_header K WHERE K.Tester=A.Tester AND K.Lot=A.Lot AND K.Spec=A.Spec LIMIT 0,1)" & "Prime_Assy,"
        strSql = strSql & "RIGHT(A.Spec,1) HGA_Type,"
        strSql = strSql & "(SELECT sum(T.TotalHGA) FROM db_" & strMassProduct & ".tabmean_avg T WHERE A.Tester=T.Tester AND A.Lot=T.Lot) ""Total.Tester"","
        strSql = strSql & "(SELECT sum(T.TotalPass) FROM db_" & strMassProduct & ".tabmean_avg T WHERE A.Tester=T.Tester AND A.Lot=T.Lot) ""Pass.Tester"","
        strSql = strSql & "(SELECT sum(T.TotalPass)/ sum(T.TotalHGA)*100 FROM db_" & strMassProduct & ".tabmean_avg T WHERE A.Tester=T.Tester AND A.Lot=T.Lot) ""Yield.Tester"","

        If InStr(strMassProduct.ToUpper, "_SDET") Then
            strSql = strSql & "3600/(SELECT sum(D.CycleTime)/SUM(E.CycleTime) FROM db_" & strMassProduct & ".tabmean_avg D LEFT JOIN db_" & strMassProduct & ".tabmean_n E USING(tester,lot,spec,shoe) WHERE A.Tester=D.Tester AND A.Lot=D.Lot AND A.Spec=D.Spec) UPH,"
        Else
            strSql = strSql & "3600/(10+(SELECT sum(D.CycleTime)/SUM(E.CycleTime) FROM db_" & strMassProduct & ".tabmean_avg D LEFT JOIN db_" & strMassProduct & ".tabmean_n E USING(tester,lot,spec,shoe) WHERE A.Tester=D.Tester AND A.Lot=D.Lot AND A.Spec=D.Spec)) UPH,"
        End If

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamFT As String = dtbParam.Rows(nParam).Item("paramFasttrack").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            Dim bCFAdd As Boolean = dtbParam.Rows(nParam).Item("param_add")
            Dim bCFMul As Boolean = dtbParam.Rows(nParam).Item("param_Mul")
            If strParamFT <> "" Then
                strSql = strSql & "SUM(A." & strParam & ")/SUM(B." & strParam & ") """ & strDisplayName & ".Lot"","
                strSql = strSql & "SUM(G." & strParamFT & ")/SUM(H." & strParamFT & ") """ & strParamFT & ".FT"","
                strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg S "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='1' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".S1"","
                strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg S "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='2' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".S2"","
                If bCFAdd = True Then
                    strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_cfadd S "
                    strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_cfadd_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='1' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".CFAdd.S1"","
                    strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_cfadd S "
                    strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_cfadd_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='2' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".CFAdd.S2"","
                End If
                If bCFMul = True Then
                    strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_cfmul S "
                    strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_cfmul_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='1' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".CFMul.S1"","
                    strSql = strSql & "(SELECT SUM(S." & strParam & ")/SUM(T." & strParam & ") FROM db_" & strMassProduct & ".tabmean_cfmul S "
                    strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_cfmul_n T USING(tester,Lot,Spec,Shoe) WHERE S.Shoe='2' AND A.Tester=S.Tester AND A.Lot=S.Lot AND A.Spec=S.Spec) """ & strDisplayName & ".CFMul.S2"","
                End If
            End If
        Next nParam

        If Right(strSql, 1) = "," Then strSql = Left(strSql, Len(strSql) - 1)
        strSql = strSql & " FROM db_" & strMassProduct & ".tabmean_avg A "
        strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n B USING(tester,lot,spec,shoe),"
        strSql = strSql & "db_" & strProductFasttrack & ".tabmean_avg G "
        strSql = strSql & "LEFT JOIN db_" & strProductFasttrack & ".tabmean_n H USING(tester,lot,spec,shoe) "
        strSql = strSql & "WHERE "
        strSql = strSql & "(A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSql = strSql & "AND (CONCAT(A.Lot,'Q')=G.Lot OR CONCAT(LEFT(A.Lot,LENGTH(A.Lot)-1),'Q')=G.Lot OR CONCAT(A.Lot,'QQ')=G.Lot) "
        strSql = strSql & "AND LEFT(A.Spec,1)='R' "
        strSql = strSql & "AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite.Rows.Count > 0 Then strSql = strSql & " AND ("
        For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
            If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
            Else
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
            End If
        Next nSliderSite
        strSql = strSql & "GROUP BY A.Tester,A.Lot,A.Spec;"
        GetSqlStringByTester = strSql
    End Function

    Private Function GetSqlStringByLot(ByVal strMassProduct As String, ByVal strProductFasttrack As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, _
    ByVal dtEnd As DateTime) As String
        Dim strSql As String
        strSql = "SELECT '" & strMassProduct & "' ProductName,"
        strSql = strSql & "'" & strProductFasttrack & "' FastTrack,"
        strSql = strSql & "A.Update_Time,"
        strSql = strSql & "A.Spec,"
        strSql = strSql & "A.Lot,"
        strSql = strSql & "E.Lot FT_Lot,"
        strSql = strSql & "E.Update_time FT_DateTime,"
        strSql = strSql & "E.Spec FT_Spec,"
        strSql = strSql & "A.SliderSite,"
        strSql = strSql & "(SELECT K.WorkID FROM db_" & strProductFasttrack & ".tabdetail_header K WHERE K.Lot=E.Lot AND K.Spec=E.Spec LIMIT 0,1)" & "FT_WorkID,"
        strSql = strSql & "(SELECT K.WorkID FROM db_" & strMassProduct & ".tabdetail_header K WHERE K.Lot=A.Lot AND K.Spec=A.Spec LIMIT 0,1)" & "Prime_WorkID,"
        strSql = strSql & "(SELECT K.Assy FROM db_" & strProductFasttrack & ".tabdetail_header K WHERE K.Lot=E.Lot AND K.Spec=E.Spec LIMIT 0,1)" & "FT_Assy,"
        strSql = strSql & "(SELECT K.Assy FROM db_" & strMassProduct & ".tabdetail_header K WHERE K.Lot=A.Lot AND K.Spec=A.Spec LIMIT 0,1)" & "Prime_Assy,"
        strSql = strSql & "RIGHT(A.Spec,1) HGA_Type,"
        strSql = strSql & "(SELECT sum(C.TotalHGA) FROM db_" & strMassProduct & ".tabmean_avg C WHERE A.Lot=C.Lot) ""Total.Lot"","
        strSql = strSql & "(SELECT sum(C.TotalPass) FROM db_" & strMassProduct & ".tabmean_avg C WHERE A.Lot=C.Lot) ""Pass.Lot"","
        strSql = strSql & "(SELECT sum(C.TotalPass)/ sum(C.TotalHGA)*100 FROM db_" & strMassProduct & ".tabmean_avg C WHERE A.Lot=C.Lot) ""Yield.Lot"","

        If InStr(strMassProduct.ToUpper, "_SDET") Then
            strSql = strSql & "3600/(SELECT sum(D.CycleTime)/SUM(E.CycleTime) FROM db_" & strMassProduct & ".tabmean_avg D LEFT JOIN db_" & strMassProduct & ".tabmean_n E USING(tester,lot,spec,shoe) WHERE A.Lot=D.Lot) UPH,"
        Else
            strSql = strSql & "3600/(10+(SELECT sum(D.CycleTime)/SUM(E.CycleTime) FROM db_" & strMassProduct & ".tabmean_avg D LEFT JOIN db_" & strMassProduct & ".tabmean_n E USING(tester,lot,spec,shoe) WHERE A.Lot=D.Lot)) UPH,"
        End If

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamFT As String = dtbParam.Rows(nParam).Item("paramFasttrack").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            If strParamFT <> "" Then
                strSql = strSql & "SUM(C." & strParam & ")/SUM(D." & strParam & ") """ & strDisplayName & ".Lot"","
                strSql = strSql & "SUM(E." & strParamFT & ")/SUM(F." & strParamFT & ") """ & strParamFT & ".FT"","
                strSql = strSql & "(SELECT SUM(G." & strParam & ")/SUM(H." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg G "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n H USING(tester,Lot,Spec,Shoe) WHERE G.Shoe='1' AND A.Lot=G.Lot AND A.Spec=G.Spec) """ & strDisplayName & ".S1"","
                strSql = strSql & "(SELECT SUM(G." & strParam & ")/SUM(H." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg G "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n H USING(tester,Lot,Spec,Shoe) WHERE G.Shoe='2' AND A.Lot=G.Lot AND A.Spec=G.Spec) """ & strDisplayName & ".S2"","
            End If
        Next nParam

        If Right(strSql, 1) = "," Then strSql = Left(strSql, Len(strSql) - 1)
        strSql = strSql & " FROM db_" & strMassProduct & ".tabmean_avg A "
        strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n B USING(tester,lot,spec,shoe) "
        strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_avgbylot C USING(lot) "
        strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_nbylot D USING(lot),"
        strSql = strSql & "db_" & strProductFasttrack & ".tabmean_avg E "
        strSql = strSql & "LEFT JOIN db_" & strProductFasttrack & ".tabmean_n F USING(tester,lot,spec,shoe) "
        strSql = strSql & "WHERE "
        strSql = strSql & "(A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSql = strSql & "AND (CONCAT(A.Lot,'Q')=E.Lot OR CONCAT(LEFT(A.Lot,LENGTH(A.Lot)-1),'Q')=E.Lot OR CONCAT(A.Lot,'QQ')=E.Lot) "
        strSql = strSql & "AND LEFT(A.Spec,1)='R' "
        strSql = strSql & "AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite.Rows.Count > 0 Then strSql = strSql & " AND ("
        For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
            If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
            Else
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
            End If
        Next nSliderSite
        strSql = strSql & "GROUP BY A.Lot,A.Spec;"
        GetSqlStringByLot = strSql
    End Function

    Private Function GetSqlStringBySpec(ByVal strMassProduct As String, ByVal strProductFasttrack As String, ByVal dtbSearch As DataTable, ByVal dtbSliderSite As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, ByVal dtEnd As DateTime) As String
        Dim strSql As String
        strSql = "SELECT '" & strMassProduct & "' ProductName,"
        strSql = strSql & "'" & strProductFasttrack & "' FastTrack,"
        'strSql = strSql & "A.Update_Time,"
        strSql = strSql & "A.Spec,"
        'strSql = strSql & "A.Lot,"
        'strSql = strSql & "E.Lot FT_Lot,"
        'strSql = strSql & "E.Tester FT_Tester,"
        'strSql = strSql & "C.Update_time FT_DateTime,"
        strSql = strSql & "C.Spec FT_Spec,"
        strSql = strSql & "A.SliderSite,"
        strSql = strSql & "RIGHT(A.Spec,1) HGA_Type,"
        strSql = strSql & "SUM(A.TotalHGA) ""Total.Spec"","
        strSql = strSql & "SUM(A.TotalPass) ""Pass.Spec"","
        strSql = strSql & "SUM(A.TotalPass)/ SUM(A.TotalHGA)*100 ""Yield.Spec"","

        If InStr(strMassProduct.ToUpper, "_SDET") Then
            strSql = strSql & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        Else
            strSql = strSql & "3600/(SUM(A.CycleTime)/SUM(B.CycleTime)) UPH,"
        End If

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
            Dim strParamFT As String = dtbParam.Rows(nParam).Item("paramFasttrack").ToString
            Dim strDisplayName As String = dtbParam.Rows(nParam).Item("param_display")
            If strParamFT <> "" Then
                strSql = strSql & "SUM(A." & strParam & ")/SUM(B." & strParam & ") """ & strDisplayName & ".Lot"","
                strSql = strSql & "SUM(C." & strParamFT & ")/SUM(D." & strParamFT & ") """ & strParamFT & ".FT"","
                strSql = strSql & "(SELECT SUM(G." & strParam & ")/SUM(H." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg G "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n H USING(tester,Lot,Spec,Shoe) WHERE G.Shoe='1' AND A.Spec=G.Spec AND A.Lot=G.Lot) """ & strDisplayName & ".S1"","
                strSql = strSql & "(SELECT SUM(G." & strParam & ")/SUM(H." & strParam & ") FROM db_" & strMassProduct & ".tabmean_avg G "
                strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n H USING(tester,Lot,Spec,Shoe) WHERE G.Shoe='2' AND A.Spec=G.Spec AND A.Lot=G.Lot) """ & strDisplayName & ".S2"","
            End If
        Next nParam

        If Right(strSql, 1) = "," Then strSql = Left(strSql, Len(strSql) - 1)
        strSql = strSql & " FROM db_" & strMassProduct & ".tabmean_avg A "
        strSql = strSql & "LEFT JOIN db_" & strMassProduct & ".tabmean_n B USING(tester,lot,spec,shoe),"
        strSql = strSql & "db_" & strProductFasttrack & ".tabmean_avg C "
        strSql = strSql & "LEFT JOIN db_" & strProductFasttrack & ".tabmean_n D USING(tester,lot,spec,shoe) "
        strSql = strSql & "WHERE "
        strSql = strSql & "(A.update_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSql = strSql & "AND "
        strSql = strSql & "(CONCAT(A.Lot,'Q')=C.Lot OR CONCAT(LEFT(A.Lot,LENGTH(A.Lot)-1),'Q')=C.Lot OR CONCAT(A.Lot,'QQ')=C.Lot "
        strSql = strSql & "OR LEFT(C.Lot,LENGTH(C.Lot)-1)=A.Lot OR LEFT(C.Lot,LENGTH(C.Lot)-2)=A.Lot OR LEFT(C.Lot,LENGTH(C.Lot)-3)=A.Lot)"

        strSql = strSql & "AND LEFT(A.Spec,1)='R' "
        strSql = strSql & "AND LEFT(C.Spec,1)='F' "
        strSql = strSql & "AND ("
        Dim strSearchBy As String = dtbSearch.TableName
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "' OR "
            Else
                strSql = strSql & "A." & strSearchBy & "='" & dtbSearch.Rows(nTester).Item(strSearchBy) & "') "
            End If
        Next nTester
        If dtbSliderSite.Rows.Count > 0 Then strSql = strSql & " AND ("
        For nSliderSite As Integer = 0 To dtbSliderSite.Rows.Count - 1
            If nSliderSite <> dtbSliderSite.Rows.Count - 1 Then
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "' OR "
            Else
                strSql = strSql & "A.SliderSite='" & dtbSliderSite.Rows(nSliderSite).Item("SliderSite") & "') "
            End If
        Next nSliderSite
        'strSql = strSql & "AND (SELECT sum(C.TotalHGA) FROM db_" & strMassProduct & ".tabmean_avg C WHERE A.Lot=C.Lot)>200 "
        strSql = strSql & "GROUP BY A.Spec;"
        GetSqlStringBySpec = strSql
    End Function

End Class

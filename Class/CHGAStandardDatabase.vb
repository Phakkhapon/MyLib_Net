
Imports MySql.Data.MySqlClient

Public Class CHGAStandardDatabase
    Private m_MySqlConn As MySqlConnection
    Public Enum enuOrderParameterBy
        eByParamID = 0
        eByOrderID
        eByName
    End Enum

    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MySqlConn = MySqlConn
    End Sub

    Public Function GetHgaStandardData(ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal dtbSearch As DataTable, ByVal dtbParam As DataTable) As DataTable
        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "Date_time,"
        strSQL = strSQL & "Hga_SN,"
        strSQL = strSQL & "TrayName,"
        strSQL = strSQL & "SortingNo,"
        strSQL = strSQL & "SilverBuyoff,"
        strSQL = strSQL & "Tester,"
        strSQL = strSQL & "HeadType,"
        strSQL = strSQL & "ProductName,"
        strSQL = strSQL & "MediaName,"
        strSQL = strSQL & "GradeName,"
        strSQL = strSQL & "CGALotName,"
        strSQL = strSQL & "SliderPos,"
        strSQL = strSQL & "StdUsage,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            Dim strDisplay As String = drParam.Item("param_display")
            If drParam.Item("param_add") Or drParam.Item("param_mul") Then strSQL = strSQL & "para" & strParamID & " '" & strDisplay & "',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
        strSQL = strSQL & "LEFT JOIN std_standard.tabspec B USING(SpecID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabtray C USING(TrayID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabmedia D USING(MediaID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabgrade E USING(GradeID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabcgalot F USING(CGALotID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tablot G USING(LotID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabproduct H USING(ProductID) "

        strSQL = strSQL & "WHERE Date_Time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' AND ("
        For nTester As Integer = 0 To dtbSearch.Rows.Count - 1
            If nTester <> dtbSearch.Rows.Count - 1 Then
                strSQL = strSQL & "SpecName='" & Left(dtbSearch.Rows(nTester).Item(0), 3) & "' OR "
            Else
                strSQL = strSQL & "SpecName='" & Left(dtbSearch.Rows(nTester).Item(0), 3) & "');"
            End If
        Next nTester
        Dim clsMySql As New CMySQL
        GetHgaStandardData = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

    Public Function GetStandardDataByProductBySortingByTray(ByVal strProductID As String, ByVal strSortingNo As String, ByVal strSortingLot As String, ByVal arTrayList As ArrayList, ByVal dtbSTDParam As DataTable) As DataTable
        Dim strSQL As String = "SELECT "
        'strSQL = strSQL & "IsExpire,"
        strSQL = strSQL & "Date_time,"
        strSQL = strSQL & "Hga_SN,"
        strSQL = strSQL & "TrayName,"
        strSQL = strSQL & "A.SortingNo,"
        strSQL = strSQL & "'" & strSortingLot & "' AS 'SortingLot',"
        'strSQL = strSQL & "CAST(CONCAT(SortingNo,'-',"
        'strSQL = strSQL & "DATE_FORMAT((SELECT MIN(date_time)FROM std_standard.tabstandard_rawdata WHERE ProductID=A.ProductID AND SortingNo=A.SortingNo),'%y%m%d'),"
        'strSQL = strSQL & "Tester,'-',IFNULL(SilverBuyoff,'')) AS CHAR) 'SortingLot',"
        strSQL = strSQL & "SilverBuyoff,"
        strSQL = strSQL & "Tester,"
        strSQL = strSQL & "HeadType,"
        strSQL = strSQL & "MediaName,"
        strSQL = strSQL & "MediaSurface,"
        strSQL = strSQL & "GradeName,"
        strSQL = strSQL & "GradeRevName,"
        strSQL = strSQL & "CGALotName,"
        strSQL = strSQL & "SliderPos,"
        strSQL = strSQL & "StdUsage,"
        For nParam As Integer = 0 To dtbSTDParam.Rows.Count - 1
            Dim drParam As DataRow = dtbSTDParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            strSQL = strSQL & "(A.para" & strParamID & "*IFNULL(J.para" & strParamID & ",1)+IFNULL(I.para" & strParamID & ",0))"
            strSQL = strSQL & "*IFNULL(L.para" & strParamID & ",1)+IFNULL(K.para" & strParamID & ",0) '" & strRTTCParam & "',"


        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
        strSQL = strSQL & "LEFT JOIN std_standard.tabtray C USING(TrayID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabmedia D USING(MediaID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabgrade E USING(GradeID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabcgalot F USING(CGALotID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabgraderev G USING(GradeRevID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabproduct H USING(ProductID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting I ON A.ProductID=I.ProductID AND A.SortingNo=I.SortingNo AND I.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting J ON A.ProductID=J.ProductID AND A.SortingNo=J.SortingNo AND J.CFTypeID=1 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray K ON A.ProductID=K.ProductID AND A.SortingNo=K.SortingNo AND A.TrayID=K.TrayID AND K.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray L ON A.ProductID=L.ProductID AND A.SortingNo=L.SortingNo AND A.TrayID=L.TrayID AND L.CFTypeID=1 "
        strSQL = strSQL & "WHERE A.ProductID=" & strProductID & " "
        strSQL = strSQL & "AND A.SortingNo=" & strSortingNo & " "
        strSQL = strSQL & "AND ("
        For nTray As Integer = 0 To arTrayList.Count - 1
            Dim nTrayID As Integer = arTrayList.Item(nTray)
            strSQL = strSQL & "A.TrayID=" & nTrayID & " OR "
        Next nTray
        If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 3) & ") "
        strSQL = strSQL & "AND IF(LENGTH(TrayName)=8,(SliderPos>326 AND sliderPos<451),SliderPos<305);"
        'strSQL = strSQL & "LotName='" & strLotName & "';"
        Dim clsMySql As New CMySQL
        GetStandardDataByProductBySortingByTray = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

    'Public Function GetStandardDataBySpecBySorting(ByVal strSpecID As String, ByVal strLotName As String, ByVal dtbParam As DataTable) As DataTable
    '    Dim strSQL As String = "SELECT "
    '    strSQL = strSQL & "Date_time,"
    '    strSQL = strSQL & "Hga_SN,"
    '    strSQL = strSQL & "TrayName,"
    '    strSQL = strSQL & "SortingNo,"
    '    strSQL = strSQL & "SilverBuyoff,"
    '    strSQL = strSQL & "Tester,"
    '    strSQL = strSQL & "HeadType,"
    '    strSQL = strSQL & "MediaName,"
    '    strSQL = strSQL & "GradeName,"
    '    strSQL = strSQL & "CGALotName,"
    '    strSQL = strSQL & "SliderPos,"
    '    strSQL = strSQL & "StdUsage,"
    '    For nParam As Integer = 0 To dtbParam.Rows.Count - 1
    '        Dim drParam As DataRow = dtbParam.Rows(nParam)
    '        Dim strRTTCParam As String = drParam.Item("param_rttc")
    '        Dim strParamID As String = drParam.Item("ParamID")
    '        'Dim strDisplay As String = drParam.Item("param_display")
    '        'If drParam.Item("param_add").ToString = "True" Or drParam.Item("param_mul").ToString = "True" Then
    '        strSQL = strSQL & "(A.para" & strParamID & "+IFNULL(H.para" & strParamID & ",0)) '" & strRTTCParam & "',"
    '        'End If
    '    Next nParam
    '    If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
    '    strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabspec B USING(SpecID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabtray C USING(TrayID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabmedia D USING(MediaID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabgrade E USING(GradeID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabcgalot F USING(CGALotID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tablot G USING(LotID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabproduct H USING(ProductID) "
    '    strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjust K USING(SpecID,LotID) "
    '    'strSQL = strSQL & "WHERE SpecID=" & strSpecID & " AND "
    '    'strSQL = strSQL & "LotName='" & strLotName & "';"
    '    Dim clsMySql As New CMySQL
    '    GetStandardDataBySpecBySorting = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    'End Function

    Public Function GetSTDAdjustbySorting(ByVal strProductID As String, ByVal strSortingNo As String, ByVal dtbParam As DataTable) As DataTable
        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "A.ProductID,"
        strSQL = strSQL & "A.SortingNo,"
        'strSQL = strSQL & "LotName,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            strSQL = strSQL & "C.para" & strParamID & " '" & strRTTCParam & ".Mul',"
            strSQL = strSQL & "B.para" & strParamID & " '" & strRTTCParam & ".Add',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting B ON A.ProductID=B.ProductID AND A.SortingNo=B.SortingNo AND B.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting C ON A.ProductID=C.ProductID AND A.SortingNo=C.SortingNo AND C.CFTypeID=1 "
        strSQL = strSQL & "WHERE A.ProductID=" & strProductID & " "
        strSQL = strSQL & "AND A.SortingNo='" & strSortingNo & "' "
        strSQL = strSQL & "GROUP BY ProductID,SortingNo;"
        Dim clsMySql As New CMySQL
        GetSTDAdjustbySorting = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

    Public Function GetSTDAdjustbySortingByTray(ByVal strProductID As String, ByVal strSortingNo As String, ByVal arTrayList As ArrayList, ByVal dtbParam As DataTable) As DataTable
        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "A.ProductID,"
        strSQL = strSQL & "A.TrayID,"
        strSQL = strSQL & "TrayName,"
        strSQL = strSQL & "A.SortingNo,"
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strParamID As String = drParam.Item("ParamID")
            strSQL = strSQL & "C.para" & strParamID & " '" & strRTTCParam & ".Mul',"
            strSQL = strSQL & "B.para" & strParamID & " '" & strRTTCParam & ".Add',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray B ON A.ProductID=B.ProductID AND A.SortingNo=B.SortingNo AND A.TrayID=B.TrayID AND B.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray C ON A.ProductID=C.ProductID AND A.SortingNo=C.SortingNo AND A.TrayID=C.TrayID AND C.CFTypeID=1 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabtray D ON A.TrayID=D.TrayID "
        strSQL = strSQL & "WHERE A.ProductID=" & strProductID & " "
        strSQL = strSQL & "AND A.SortingNo='" & strSortingNo & "' "
        strSQL = strSQL & "AND ("
        For nTray As Integer = 0 To arTrayList.Count - 1
            Dim nTrayID As Integer = arTrayList.Item(nTray)
            strSQL = strSQL & "A.TrayID=" & nTrayID & " OR "
        Next nTray
        If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 3) & ") "
        strSQL = strSQL & "GROUP BY A.ProductID,A.SortingNo,A.TrayID;"
        Dim clsMySql As New CMySQL
        GetSTDAdjustbySortingByTray = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function

    'Specail product : Spyglass 
    Public Function GetStandardDataSpecialProductBySortingByTray(ByVal strProductID As String, ByVal strSortingNo As String, ByVal strSortingLot As String, ByVal arTrayList As ArrayList, ByVal dtbSTDParam As DataTable) As DataTable
        Dim strSQL As String = "SELECT "
        'strSQL = strSQL & "IsExpire,"
        strSQL = strSQL & "Date_time,"
        strSQL = strSQL & "Hga_SN,"
        strSQL = strSQL & "TrayName,"
        strSQL = strSQL & "A.SortingNo,"
        strSQL = strSQL & "'" & strSortingLot & "' AS 'SortingLot',"
        'strSQL = strSQL & "CAST(CONCAT(SortingNo,'-',"
        'strSQL = strSQL & "DATE_FORMAT((SELECT MIN(date_time)FROM std_standard.tabstandard_rawdata WHERE ProductID=A.ProductID AND SortingNo=A.SortingNo),'%y%m%d'),"
        'strSQL = strSQL & "Tester,'-',IFNULL(SilverBuyoff,'')) AS CHAR) 'SortingLot',"
        strSQL = strSQL & "SilverBuyoff,"
        strSQL = strSQL & "Tester,"
        strSQL = strSQL & "HeadType,"
        strSQL = strSQL & "MediaName,"
        strSQL = strSQL & "MediaSurface,"
        strSQL = strSQL & "GradeName,"
        strSQL = strSQL & "GradeRevName,"
        strSQL = strSQL & "CGALotName,"
        strSQL = strSQL & "SliderPos,"
        strSQL = strSQL & "StdUsage,"
        For nParam As Integer = 0 To dtbSTDParam.Rows.Count - 1
            Dim drParam As DataRow = dtbSTDParam.Rows(nParam)
            Dim strRTTCParam As String = drParam.Item("param_rttc")
            Dim strDisplayName As String = drParam.Item("param_display")
            Dim strParamID As String = drParam.Item("ParamID")
            strSQL = strSQL & "(A.para" & strParamID & "*IFNULL(J.para" & strParamID & ",1)+IFNULL(I.para" & strParamID & ",0))"
            strSQL = strSQL & "*IFNULL(L.para" & strParamID & ",1)+IFNULL(K.para" & strParamID & ",0) '" & strDisplayName & "'," '//strRTTCParam & "',"
        Next nParam
        If Right(strSQL, 1) = "," Then strSQL = Left(strSQL, strSQL.Length - 1) & " "
        strSQL = strSQL & "FROM std_standard.tabstandard_rawdata A "
        strSQL = strSQL & "LEFT JOIN std_standard.tabtray C USING(TrayID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabmedia D USING(MediaID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabgrade E USING(GradeID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabcgalot F USING(CGALotID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabgraderev G USING(GradeRevID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabproduct H USING(ProductID) "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting I ON A.ProductID=I.ProductID AND A.SortingNo=I.SortingNo AND I.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysorting J ON A.ProductID=J.ProductID AND A.SortingNo=J.SortingNo AND J.CFTypeID=1 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray K ON A.ProductID=K.ProductID AND A.SortingNo=K.SortingNo AND A.TrayID=K.TrayID AND K.CFTypeID=0 "
        strSQL = strSQL & "LEFT JOIN std_standard.tabstandard_adjustbysortingbytray L ON A.ProductID=L.ProductID AND A.SortingNo=L.SortingNo AND A.TrayID=L.TrayID AND L.CFTypeID=1 "
        strSQL = strSQL & "WHERE A.ProductID=" & strProductID & " "
        strSQL = strSQL & "AND A.SortingNo=" & strSortingNo & " "
        strSQL = strSQL & "AND ("
        For nTray As Integer = 0 To arTrayList.Count - 1
            Dim nTrayID As Integer = arTrayList.Item(nTray)
            strSQL = strSQL & "A.TrayID=" & nTrayID & " OR "
        Next nTray
        If Right(strSQL, 3) = "OR " Then strSQL = Left(strSQL, strSQL.Length - 3) & ") "
        strSQL = strSQL & "AND IF(LENGTH(TrayName)=8,(SliderPos>326 AND sliderPos<451),SliderPos<305);"
        'strSQL = strSQL & "LotName='" & strLotName & "';"
        Dim clsMySql As New CMySQL
        GetStandardDataSpecialProductBySortingByTray = clsMySql.CommandMySqlDataTable(strSQL, m_MySqlConn)
    End Function
    'Specail get stddata only Spyglass Product
   

 


End Class

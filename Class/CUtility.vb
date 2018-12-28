
'This class use for temperary job
Imports System.IO
Imports MySql.Data.MySqlClient

Public Class CUtility

    Public Function GetConnString(ByVal strServer As String, ByVal strUser As String, _
ByVal strPassword As String, ByVal strPort As String, Optional ByVal strDatabase As String = "") As String
        Dim SsqlcON As String
        SsqlcON = "server=" & strServer & ";"
        SsqlcON = SsqlcON & "uid=" & strUser & ";"
        SsqlcON = SsqlcON & "pwd=" & strPassword & ";"
        If strDatabase <> "" Then SsqlcON = SsqlcON & "database=" & strDatabase & ";"
        SsqlcON = SsqlcON & "port=" & strPort & ";"
        'SsqlcON = SsqlcON & "Mode=3;"
        SsqlcON = SsqlcON & "Connection Lifetime=15;"
        GetConnString = SsqlcON
    End Function

    Public Sub UpdateShowParameterByProduct()
        Dim mySqlConn As New MySqlConnection
        mySqlConn.ConnectionString = GetConnString("172.20.65.5", "rttc", "rttc", "3307")
        mySqlConn.Open()

        Dim clsParameter As New CParameterRTTCMapping(mySqlConn)
        Dim dtbProduct As DataTable = clsParameter.GetProductList(enuProductType.enuProductAll)
        Dim strSQL As String
        Dim clsSql As New CMySQL

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item("Product").ToString
            mySqlConn.Close()
            mySqlConn.Open()
            strSQL = "SELECT tag_id FROM db_" & strProduct & ".tabfactor_value LIMIT 0,100"
            Dim dtbDataCount As DataTable = clsSql.CommandMySqlDataTable(strSQL, mySqlConn)
            If dtbDataCount.Rows.Count > 99 Then
                Dim dtbParam As DataTable = clsParameter.GetParamByProduct(strProduct, True)
                For nParam As Integer = 0 To dtbParam.Rows.Count - 1
                    Dim strParam As String = dtbParam.Rows(nParam).Item("param_rttc").ToString
                    strSQL = "SELECT tag_id FROM db_" & strProduct & ".tabfactor_value "
                    strSQL = strSQL & "WHERE (NOT " & strParam & " IS NULL) LIMIT 0,1;"
                    Dim dtbTemp As DataTable = clsSql.CommandMySqlDataTable(strSQL, mySqlConn)
                    If dtbTemp.Rows.Count = 0 Then
                        strSQL = "UPDATE db_" & strProduct & ".tabparameterbyproduct SET isEnable=False "
                        strSQL = strSQL & "WHERE param_rttc='" & strParam & "';"
                    Else
                        strSQL = "UPDATE db_" & strProduct & ".tabparameterbyproduct SET isEnable=True "
                        strSQL = strSQL & "WHERE param_rttc='" & strParam & "';"
                    End If
                    clsSql.CommandMySqlDataTable(strSQL, mySqlConn)
                Next nParam
            End If
        Next nProduct
    End Sub

    Public Sub GetHistoryAdjustCF(ByVal strProduct As String, ByVal strParam As String, ByVal mySqlConn As MySqlConnection)

        Dim strSQL As String = ""

        Dim strStart As String = "2013-08-16 00:00:00"
        Dim strEnd As String = Format(Now, "yyyy-MM-dd HH:mm:ss")

        strSQL = strSQL & "WHERE param_rttc='" & strParam & "' "
        strSQL = strSQL & "AND Test_Time BETWEEN '" & strStart & "' AND '" & strEnd & "' "
        strSQL = strSQL & "ORDER BY tester,test_time;"
        Dim clsSql As New CMySQL
        Dim dtbData As DataTable
        dtbData = clsSql.CommandMySqlDataTable(strSQL, mySqlConn)

        For nData As Integer = 0 To dtbData.Rows.Count - 1
            Dim strID As String = dtbData.Rows(nData).Item("ID")
            strSQL = "SELECT * FROM db_" & strProduct & ".tabhistory_dataadjcf "
            strSQL = strSQL & "WHERE ID='" & strID & "' "
            strSQL = strSQL & "AND param_rttc='" & strParam & "' "
            strSQL = strSQL & "ORDER BY EndLotTime;"
            Dim dtbHis As DataTable
            dtbHis = clsSql.CommandMySqlDataTable(strSQL, mySqlConn)

            Dim strData As String = ""
            For nCol As Integer = 0 To dtbData.Columns.Count - 2
                strData = strData & dtbData.Rows(nData).Item(nCol).ToString & ","
            Next nCol
            Dim strFilepath As String = "C:\DataAdjust_" & strParam & ".csv"
            strData = strData & dtbData.Rows(nData).Item(dtbData.Columns.Count - 1).ToString & Environment.NewLine
            File.AppendAllText(strFilepath, strData)

            For nHis As Integer = 0 To dtbHis.Rows.Count - 1
                Dim strHis As String = ""
                For nCol As Integer = 0 To dtbHis.Columns.Count - 2
                    strHis = strHis & dtbHis.Rows(nHis).Item(nCol).ToString & ","
                Next nCol
                strHis = strHis & dtbHis.Rows(nHis).Item(dtbHis.Columns.Count - 1).ToString & Environment.NewLine
                File.AppendAllText(strFilepath, strHis)
            Next nHis
            File.AppendAllText(strFilepath, Environment.NewLine)
        Next nData
    End Sub

    Public Sub xxD(ByVal mySqlConn As MySqlConnection)
        Dim strSQL As String
        Dim strProduct As String = "db_trails_dct_sdet"
        Dim strFile() As String = File.ReadAllLines("R:\Karnt\Xlot_Sisterlot Raw Data.csv")
        Dim strResultPath As String = "R:\Result.csv"
        Dim strHeader As String = "Hga_SN,Test_time,Tester,Lot,Spec,MEW" & Environment.NewLine
        File.WriteAllText(strResultPath, strHeader)

        'For nFile As Integer = 0 To strFile.Length - 1
        strSQL = "SELECT A.hga_sn,A.test_time,A.tester,A.Lot,A.Spec,B.MEW6T FROM " & strProduct & ".tabdetail_header A "
        strSQL = strSQL & "LEFT JOIN " & strProduct & ".tabfactor_value B USING(test_time,tester) "
        strSQL = strSQL & "WHERE "
        Dim strHga_sn As String
        For nHga As Integer = 1302 To strFile.Length - 2
            strHga_sn = strFile(nHga).Split(",")(4)
            strSQL = strSQL & "hga_sn='" & strHga_sn & "' OR "
        Next nHga
        strSQL = strSQL & "hga_sn='" & strFile(strFile.Length - 1).Split(",")(4) & "';"

        Dim clsSql As New CMySQL
        Dim dtbDetail As DataTable = clsSql.CommandMySqlDataTable(strSQL, mySqlConn)

        For nData As Integer = 0 To dtbDetail.Rows.Count - 1
            strHga_sn = dtbDetail.Rows(nData).Item("Hga_sn").ToString
            Dim strTestTime As String = dtbDetail.Rows(nData).Item("Test_time").ToString
            Dim strTester As String = dtbDetail.Rows(nData).Item("Tester").ToString
            Dim strLot As String = dtbDetail.Rows(nData).Item("Lot").ToString
            Dim strSpec As String = dtbDetail.Rows(nData).Item("Spec").ToString
            Dim strMEW As String = dtbDetail.Rows(nData).Item("MEW6T").ToString
            Dim strWrite As String = strHga_sn & "," & strTestTime & "," & strTester & "," & strLot & "," & strSpec & "," & strMEW & Environment.NewLine
            File.AppendAllText(strResultPath, strWrite)
        Next nData
        'Next nFile

    End Sub

    Public Sub TransferDataToNewServer()

        Dim myConnOld As New MySqlConnection
        Dim myConnNew As New MySqlConnection

        myConnOld.ConnectionString = GetConnString("wdtbdbte02", "rttc_net", "rttc_net", "3306")
        myConnOld.Open()

        myConnNew.ConnectionString = GetConnString("wdtbtsd13", "rttc_net", "rttc_net", "3306")
        myConnNew.Open()

        Dim clsMySql As New CMySQL

        Dim clsDatabaseOld As New CDatabaseManage(myConnOld)
        Dim dtbDatabase As DataTable = clsDatabaseOld.GetDatabaseList()
        Dim clsDatabaseNew As New CDatabaseManage(myConnNew)

        Dim dtbTableSkip As New DataTable
        dtbTableSkip.Columns.Add("TableSkip")
        dtbTableSkip.Rows.Add("tabdetail_header")
        dtbTableSkip.Rows.Add("tabfactor_cfadd")
        dtbTableSkip.Rows.Add("tabfactor_cfmul")
        dtbTableSkip.Rows.Add("tabfactor_value")
        dtbTableSkip.Rows.Add("tabhistory_adjcf")
        dtbTableSkip.Rows.Add("tabhistory_dataadjcf")
        dtbTableSkip.Rows.Add("tabhistory_dataendlot")
        dtbTableSkip.Rows.Add("tabhistory_dataendlotbytester")
        dtbTableSkip.Rows.Add("tabhistory_sigmaendlotbytester")
        dtbTableSkip.Rows.Add("tabmean_avgbylot")
        dtbTableSkip.Rows.Add("tabmean_nbylot")
        dtbTableSkip.Rows.Add("tabmean_cfadd")
        dtbTableSkip.Rows.Add("tabmean_cfadd_n")
        dtbTableSkip.Rows.Add("tabmean_cfmul")
        dtbTableSkip.Rows.Add("tabmean_cfmul_n")
        dtbTableSkip.Rows.Add("tabmean_avg")
        dtbTableSkip.Rows.Add("tabmean_n")
        'dtbTableSkip.Rows.Add("tabsigmabyday")
        dtbTableSkip.Rows.Add("tabsummary_hgadefect")
        dtbTableSkip.Rows.Add("tabtester")
        dtbTableSkip.Rows.Add("testerhistory")
        dtbTableSkip.Rows.Add("tabuseraction")

        For nDB As Integer = 24 To 24 'dtbDatabase.Rows.Count - 1
StartCopy:

            Dim strDatabaseName As String = dtbDatabase.Rows(nDB).Item("Database").ToString
            If strDatabaseName <> "information_schema" And strDatabaseName <> "mysql" And strDatabaseName <> "test" Then
                Try
                    Dim strSQLDB As String = ""
                    'Try
                    '    strSQLDB = "DROP DATABASE " & strDatabaseName & ";"
                    '    clsMySql.CommandMySqlDataTable(strSQLDB, myConnNew)
                    'Catch exDrop As Exception
                    'End Try
                    'strSQLDB = "CREATE DATABASE " & strDatabaseName & ";"
                    'clsMySql.CommandMySqlDataTable(strSQLDB, myConnNew)

                    Dim dtbTable As DataTable = clsDatabaseOld.ShowTableInDatabase(strDatabaseName)

                    For nTable As Integer = 0 To dtbTable.Rows.Count - 1
                        Dim strTableName As String = dtbTable.Rows(nTable).Item(0).ToString
                        Dim strCreateTable As String = "SHOW CREATE TABLE " & strDatabaseName & "." & strTableName
                        Dim dtbCreateTable As DataTable = clsMySql.CommandMySqlDataTable(strCreateTable, myConnOld)
                        strSQLDB = dtbCreateTable.Rows(0).Item("Create Table").ToString
                        If myConnNew.State = ConnectionState.Closed Then myConnNew.Open()
                        myConnNew.ChangeDatabase(strDatabaseName)
                        'clsMySql.CommandMySqlDataTable(strSQLDB, myConnNew)
                        Dim dtrSkip() As DataRow = dtbTableSkip.Select("TableSkip='" & strTableName & "'")
                        If dtrSkip.Length = 0 Then

                            Dim strSQL As String = "SELECT * FROM " & strDatabaseName & "." & strTableName
                            Dim dtbData As DataTable = clsMySql.CommandMySqlDataTable(strSQL, myConnOld)
                            If dtbData.Rows.Count > 0 Then
                                For nData As Integer = 0 To dtbData.Rows.Count - 1
                                    Dim strInsert As String = "REPLACE INTO " & strDatabaseName & "." & strTableName & "("
                                    Dim strSelectValue As String = " SELECT "
                                    For nCol As Integer = 0 To dtbData.Columns.Count - 1
                                        Dim strCol As String = dtbData.Columns(nCol).ColumnName
                                        Dim strValue As String = dtbData.Rows(nData).Item(strCol).ToString
                                        If strValue = "True" Then
                                            strValue = "1"
                                        ElseIf strValue = "False" Then
                                            strValue = "0"
                                        End If
                                        If strValue <> "" Then
                                            If dtbData.Rows(nData).Item(strCol).GetType() Is System.Type.GetType("System.DateTime") Then
                                                strValue = Format(CDate(strValue), "yyyy-MM-dd HH:mm:ss")
                                            End If
                                        End If
                                        If nCol = dtbData.Columns.Count - 1 Then
                                            strInsert = strInsert & strCol & ")"
                                            strSelectValue = strSelectValue & "'" & strValue & "';"
                                        Else
                                            strInsert = strInsert & strCol & ","
                                            strSelectValue = strSelectValue & "'" & strValue & "',"
                                        End If
                                    Next nCol
                                    strInsert = strInsert & strSelectValue
                                    strInsert = Replace(strInsert, """", "")
                                    clsMySql.CommandMySqlDataTable(strInsert, myConnNew)
                                Next nData
                            End If
                        End If
                    Next nTable

                Catch ex As Exception
                    Dim clsCreateProduct As New CDatabaseManage(myConnNew)
                    clsCreateProduct.AddNewProduct(strDatabaseName)
                    GoTo StartCopy
                End Try
            End If
NextProduct:
        Next nDB

    End Sub

    Public Function GetTesterSoftwareVersion(ByVal MySqlConn As MySqlConnection) As DataTable
        Dim dtbTesterInfo As New DataTable
        dtbTesterInfo.Columns.Add("Test_time")
        dtbTesterInfo.Columns.Add("ProductName")
        dtbTesterInfo.Columns.Add("Tester")
        dtbTesterInfo.Columns.Add("WTrayVersion")
        dtbTesterInfo.Columns.Add("DOVERHMI")
        dtbTesterInfo.Columns.Add("DOVERSCRIPT")
        dtbTesterInfo.Columns("Tester").Unique = True

        Dim clsProduct As New CParameterRTTCMapping(MySqlConn)
        Dim dtbProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductAll)
        Dim clsMySql As New CMySQL
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strDBName As String = "db_" & dtbProduct.Rows(nProduct).Item("Product")
            If MySqlConn.State = ConnectionState.Closed Then MySqlConn.Open()
            MySqlConn.ChangeDatabase(strDBName)
            Dim strSQL As String
            strSQL = "SELECT "
            strSQL = strSQL & "MAX(test_time) Test_time,"
            strSQL = strSQL & "'" & dtbProduct.Rows(nProduct).Item("Product") & "' ProductName,"
            strSQL = strSQL & "Tester,"
            strSQL = strSQL & "(SELECT WTrayVersion FROM tabdetail_header WHERE test_time_bigint="
            strSQL = strSQL & "(SELECT MAX(Test_time_bigint) FROM tabdetail_header WHERE tester=A.tester) AND tester=A.Tester) WTrayVersion,"
            strSQL = strSQL & "(SELECT DOVERHMI FROM tabdetail_header WHERE test_time_bigint="
            strSQL = strSQL & "(SELECT MAX(Test_time_bigint) FROM tabdetail_header WHERE tester=A.tester) AND tester=A.Tester) DOVERHMI,"
            strSQL = strSQL & "(SELECT DOVERSCRIPT FROM tabdetail_header WHERE test_time_bigint="
            strSQL = strSQL & "(SELECT MAX(Test_time_bigint) FROM tabdetail_header WHERE tester=A.tester) AND tester=A.Tester) DOVERSCRIPT "
            strSQL = strSQL & "FROM tabdetail_header A "
            strSQL = strSQL & "GROUP BY tester "
            strSQL = strSQL & "ORDER BY tester;"
            Dim dtbInfo As DataTable = clsMySql.CommandMySqlDataTable(strSQL, MySqlConn)
            For nTester As Integer = 0 To dtbInfo.Rows.Count - 1
                Dim strSelect As String = "Tester='" & dtbInfo.Rows(nTester).Item("Tester") & "'"
                Dim drSelect() As DataRow = dtbTesterInfo.Select(strSelect)
                If drSelect.Length = 0 Then
                    dtbTesterInfo.Rows.Add(dtbInfo.Rows(nTester).ItemArray)
                Else
                    Dim dtOld As DateTime = drSelect(0).Item("Test_time")
                    Dim dtNew As DateTime = dtbInfo.Rows(nTester).Item("Test_time")
                    If dtNew > dtOld Then
                        drSelect(0).Item("Test_time") = dtNew
                        drSelect(0).Item("ProductName") = dtbInfo.Rows(nTester).Item("ProductName")
                        drSelect(0).Item("WTrayVersion") = dtbInfo.Rows(nTester).Item("WTrayVersion")
                        drSelect(0).Item("DOVERHMI") = dtbInfo.Rows(nTester).Item("DOVERHMI")
                        drSelect(0).Item("DOVERSCRIPT") = dtbInfo.Rows(nTester).Item("DOVERSCRIPT")
                    End If
                End If
            Next nTester
        Next nProduct
        GetTesterSoftwareVersion = dtbTesterInfo
    End Function

    Public Function GetInsertCommandStringTable(ByVal strTable As String, ByVal dtbTable As DataTable, ByVal bInsertCompleteFieldOnly As Boolean) As String
        Dim strSQL As String = ""
        For nRow As Integer = 0 To dtbTable.Rows.Count - 1
            Dim rsState As DataRowState = dtbTable.Rows(nRow).RowState
            If rsState <> DataRowState.Deleted Then
                Dim strSqlInsert As String = "INSERT INTO " & strTable & "("
                Dim strSqlValue As String = "VALUES("
                Dim strSqlDuplicate As String = "ON DUPLICATE KEY UPDATE "
                Dim nFieldCount As Integer = 0
                For nCol As Integer = 0 To dtbTable.Columns.Count - 1
                    Dim strSplit As String = ","
                    Dim strSplitUpdate As String = ","
                    If nCol = dtbTable.Columns.Count - 1 Then
                        strSplit = ") "
                        strSplitUpdate = ";"
                    End If
                    Dim strColName As String = dtbTable.Columns(nCol).ColumnName
                    Dim strValue As String = dtbTable.Rows(nRow).Item(nCol).ToString
                    If strValue <> "" Then
                        nFieldCount = nFieldCount + 1
                        strSqlInsert = strSqlInsert & strColName & strSplit
                        strSqlValue = strSqlValue & "'" & strValue & "'" & strSplit
                        strSqlDuplicate = strSqlDuplicate & strColName & "='" & strValue & "'" & strSplitUpdate
                    End If
                Next nCol
                If bInsertCompleteFieldOnly Then
                    If nFieldCount = dtbTable.Columns.Count Then strSQL = strSQL & strSqlInsert & strSqlValue & strSqlDuplicate
                Else
                    If nFieldCount > 0 Then strSQL = strSQL & strSqlInsert & strSqlValue & strSqlDuplicate
                End If
            End If
        Next nRow
        GetInsertCommandStringTable = strSQL
    End Function

    Public Function GetInsertCommandStringRow(ByVal strTable As String, ByVal drRow As DataRow, Optional ByVal bHasSubQuery As Boolean = True) As String
        Dim strSQL As String = ""

        Dim strSqlInsert As String = "INSERT INTO " & strTable & "("
        Dim strSqlValue As String = "SELECT "
        Dim strSqlDuplicate As String = "ON DUPLICATE KEY UPDATE "
        Dim strSqlDelete As String = "DELETE FROM " & strTable & " "
        Dim strWhereCause As String = "WHERE "

        For nCol As Integer = 0 To drRow.ItemArray.Length - 1
            Dim strColName As String = drRow.Table.Columns(nCol).ColumnName
            Dim strValue As String = ""
            If drRow.RowState = DataRowState.Deleted Then
                strValue = drRow.Item(nCol, DataRowVersion.Original).ToString
            ElseIf drRow.RowState = DataRowState.Detached Then
                strValue = drRow.Item(nCol, DataRowVersion.Default).ToString
            Else
                strValue = drRow.Item(nCol).ToString
            End If
            If drRow.Table.Columns(nCol).DataType Is GetType(DateTime) And strValue <> "" Then
                strValue = "'" & Format(CDate(strValue), "yyyy-MM-dd HH:mm:ss") & "'"
            ElseIf drRow.Table.Columns(nCol).DataType Is GetType(Boolean) Then
                If strValue = "" Then strValue = "False"
            ElseIf strColName.ToLower <> "failureid" Then
                strValue = "'" & strValue & "'"
            End If
            strSqlInsert = strSqlInsert & strColName & ","
            strSqlValue = strSqlValue & strValue & ","
            If strColName.ToLower <> "failureid" Then
                strSqlDuplicate = strSqlDuplicate & strColName & "=" & strValue & ","
            End If
            If strValue <> "" And strValue <> "''" Then strWhereCause = strWhereCause & strColName & "=" & strValue & " AND "
        Next nCol

        If Right(strSqlInsert, 1) = "," Then strSqlInsert = Left(strSqlInsert, strSqlInsert.Length - 1) & ") "
        If Right(strSqlValue, 1) = "," Then strSqlValue = Left(strSqlValue, strSqlValue.Length - 1) & " "
        If Right(strSqlDuplicate, 1) = "," Then strSqlDuplicate = Left(strSqlDuplicate, strSqlDuplicate.Length - 1)
        If Right(strWhereCause, 5) = " AND " Then strWhereCause = Left(strWhereCause, strWhereCause.Length - 5)
        If bHasSubQuery Then
            strSqlValue = strSqlValue & "FROM " & strTable & " "
        End If
        If drRow.RowState = DataRowState.Deleted Or drRow.RowState = DataRowState.Detached Then
            strSQL = strSqlDelete & strWhereCause
        Else
            strSQL = strSQL & strSqlInsert & strSqlValue & strSqlDuplicate
        End If
        GetInsertCommandStringRow = strSQL & ";"

    End Function

    'Public Function GetInsertCommandStringRow(ByVal strTable As String, ByVal drvRow As DataRowView) As String
    '    Dim strSQL As String = ""

    '    Dim strSqlInsert As String = "INSERT INTO " & strTable & "("
    '    Dim strSqlValue As String = "SELECT "
    '    Dim strSqlDuplicate As String = "ON DUPLICATE KEY UPDATE "
    '    If drvRow.Row.RowState = DataRowState.Deleted Then
    '        strSqlInsert = "DELETE FROM " & strTable & " "
    '        strSqlDuplicate = "WHERE "
    '    End If


    '    For nCol As Integer = 0 To drvRow.DataView.Table.Columns.Count - 1
    '        Dim strColName As String = drvRow.DataView.Table.Columns(nCol).ColumnName
    '        Dim strValue As String = ""
    '        If drvRow.Row.RowState = DataRowState.Deleted Then
    '            strValue = drvRow.Row.Item(nCol, DataRowVersion.Original).ToString
    '        Else
    '            strValue = drvRow.Row.Item(nCol).ToString
    '        End If
    '        'If strValue <> "" Then
    '        If drvRow.DataView.Table.Columns(nCol).DataType Is GetType(DateTime) Then
    '            strValue = "'" & Format(drvRow.Row.Item(nCol), "yyyy-MM-dd HH:mm:ss") & "'"
    '        ElseIf drvRow.DataView.Table.Columns(nCol).DataType Is GetType(Boolean) Then
    '            If strValue = "" Then strValue = "False"
    '        Else 'And strColName.ToLower <> "failureid" Then
    '            strValue = "'" & strValue & "'"
    '        End If
    '        strSqlInsert = strSqlInsert & strColName & ","
    '        strSqlValue = strSqlValue & strValue & ","
    '        strSqlDuplicate = strSqlDuplicate & strColName & "=" & strValue & ","
    '        'End If
    '    Next nCol

    '    If Right(strSqlInsert, 1) = "," Then strSqlInsert = Left(strSqlInsert, strSqlInsert.Length - 1) & ") "
    '    If Right(strSqlValue, 1) = "," Then strSqlValue = Left(strSqlValue, strSqlValue.Length - 1) & " FROM " & strTable & " "
    '    If Right(strSqlDuplicate, 1) = "," Then strSqlDuplicate = Left(strSqlDuplicate, strSqlDuplicate.Length - 1)

    '    If drvRow.Row.RowState = DataRowState.Deleted Then
    '        strSQL = strSqlInsert & strSqlDuplicate
    '        'ElseIf drRow.RowState = DataRowState.Added Then
    '        '    strSQL = strSqlInsert & strSqlValue
    '    Else
    '        strSQL = strSQL & strSqlInsert & strSqlValue & strSqlDuplicate
    '    End If
    '    GetInsertCommandStringRow = strSQL & ";"

    'End Function


    Public Function GetHistorySetting(ByVal strProduct As String, ByVal strPageClass As String, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal MySqlConn As MySqlConnection) As DataTable
        Dim strSQL As String = "SELECT "
        strSQL = strSQL & "ActionTime,"
        strSQL = strSQL & "ActionUser,"
        'strSQL = strSQL & "PageClass,"
        'strSQL = strSQL & "ProductName,"
        strSQL = strSQL & "SQLScript "
        strSQL = strSQL & "FROM ctr_user_rttc.tabuseraction "
        strSQL = strSQL & "WHERE ActionTime BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "AND PageClass='" & strPageClass & "' "
        strSQL = strSQL & "AND ProductName='" & strProduct & "' "
        strSQL = strSQL & "ORDER BY ActionTime DESC;"
        Dim clsMySql As New CMySQL
        GetHistorySetting = clsMySql.CommandMySqlDataTable(strSQL, MySqlConn)
    End Function

    Public Function GetInsertSettingHistorySQL(ByVal drRowData As DataRow, ByVal strPrimaryKey As String, ByVal strUser As String, ByVal strProduct As String, ByVal strPageClass As String) As String

        Dim strValueChange As String = drRowData.RowState.ToString & "-" & strPrimaryKey & "#"
        If drRowData.RowState <> DataRowState.Deleted Then
            For nCol As Integer = 0 To drRowData.Table.Columns.Count - 1
                Dim strColName As String = drRowData.Table.Columns(nCol).ColumnName
                Dim strOldValue As String = ""

                If drRowData.HasVersion(DataRowVersion.Original) Then
                    strOldValue = drRowData.Item(nCol, DataRowVersion.Original).ToString()
                End If

                Dim strNowValue As String = drRowData.Item(nCol, DataRowVersion.Current).ToString
                If strOldValue <> strNowValue Then
                    strValueChange = strValueChange & strColName & ":" & strOldValue & "-->" & strNowValue & ","
                End If
            Next nCol
        End If
 
        Dim strActionTime As String = Format(Now, "yyyy-MM-dd HH:mm:ss")
        Dim strLogSQL As String
        strLogSQL = "INSERT INTO ctr_user_rttc.tabuseraction("
        strLogSQL = strLogSQL & "ActionTime,"
        strLogSQL = strLogSQL & "ActionUser,"
        strLogSQL = strLogSQL & "PageClass,"
        strLogSQL = strLogSQL & "ProductName,"
        strLogSQL = strLogSQL & "SqlScript) VALUES("
        strLogSQL = strLogSQL & "'" & strActionTime & "',"
        strLogSQL = strLogSQL & "'" & strUser & "',"
        strLogSQL = strLogSQL & "'" & strPageClass & "',"
        strLogSQL = strLogSQL & "'" & strProduct & "',"
        strLogSQL = strLogSQL & "'" & strValueChange & "');"
        GetInsertSettingHistorySQL = strLogSQL
    End Function

End Class

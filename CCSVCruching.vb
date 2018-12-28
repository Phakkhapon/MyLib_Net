
Imports System.IO
Imports MySql.Data.MySqlClient
Public Class CCSVCruching

    Public Sub CrunchCSVData(ByVal strProduct As String, ByVal strFullFileName As String, ByVal dtbHeader As DataTable, ByVal dtbRTTCParam As DataTable, ByVal myRTTCConn As MySqlConnection)

        'Dim strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFolder & ";Extended Properties=""text;HDR=No;FMT=Delimited"";"
        'Dim dtbCSVData As New DataTable
        'Using Adp As New OleDb.OleDbDataAdapter("SELECT * FROM [" & strFileName & "]", strConn)
        '    Adp.Fill(dtbCSVData)
        'End Using
        'Edit Colunm name,
        Dim strData() As String = File.ReadAllLines(strFullFileName)
        Dim strHeader() As String = strData(0).Split(",")
        Dim nStart As Integer = 0
        If strHeader.Length = 0 Then
            nStart = 1
            strHeader = strData(1).Split(",")
        End If
        Dim dtbCSVData As New DataTable
        For nCol As Integer = 0 To strHeader.Length - 1
            Dim strColName As String = strHeader(nCol)
            dtbCSVData.Columns.Add(strColName)
        Next nCol
        For nLine As Int32 = nStart + 1 To strData.Length - 1
            Dim strValue() As String = strData(nLine).Split(",")
            dtbCSVData.Rows.Add(strValue)
        Next

        Dim dtbDetail As DataTable = dtbCSVData.DefaultView.ToTable
        Dim dtbValue As DataTable = dtbCSVData.DefaultView.ToTable
        Dim strHeaderDetail As String = "INSERT INTO db_" & strProduct & ".tabdetail_header(tag_id,test_time_bigint,"
        Dim strHeaderValue As String = "INSERT INTO db_" & strProduct & ".tabfactor_value(tag_id,test_time_bigint,test_time,tester,"
        For nCol As Integer = 0 To dtbCSVData.Columns.Count - 1
            Dim strColName As String = dtbCSVData.Columns(nCol).ColumnName
            Dim drSelect() As DataRow = dtbHeader.Select("[HeaderName]='" & strColName & "'")
            If drSelect.Length = 0 Then
                dtbDetail.Columns.Remove(strColName)
            Else
                strHeaderDetail = strHeaderDetail & strColName & ","
            End If
        Next
        If dtbDetail.Columns.Count = 0 Then
            Throw New Exception("Files:" & strFullFileName & " is not raw data file")
        End If
        For nCol As Integer = 0 To dtbCSVData.Columns.Count - 1
            Dim strColName As String = dtbCSVData.Columns(nCol).ColumnName
            Dim drSelect() As DataRow = Nothing
            drSelect = dtbRTTCParam.Select("[param_rttc]='" & strColName & "'")

            If drSelect.Length = 0 Then
                drSelect = dtbRTTCParam.Select("[paramMachine]='" & strColName & "'")
                If drSelect.Length = 0 Then
                    dtbValue.Columns.Remove(strColName)
                Else
                    Dim strParam As String = drSelect(0).Item("param_rttc")
                    If InStr(strParam, "para") = 1 Then
                        dtbValue.Columns(strColName).ColumnName = strParam
                        strHeaderValue = strHeaderValue & strColName & ","
                    Else
                        dtbValue.Columns.Remove(strColName)
                    End If
                End If
            Else
                strHeaderValue = strHeaderValue & strColName & ","
            End If
        Next

        If Right(strHeaderDetail, 1) = "," Then strHeaderDetail = Left(strHeaderDetail, Len(strHeaderDetail) - 1) & ") VALUES("
        If Right(strHeaderValue, 1) = "," Then strHeaderValue = Left(strHeaderValue, Len(strHeaderValue) - 1) & ") VALUES("

        Dim clsMySQL As New CMySQL
        For nData As Int32 = 0 To dtbDetail.Rows.Count - 1
            Dim strDetailValue As String = strHeaderDetail
            Dim strValue As String = strHeaderValue
            Dim drDetail As DataRow = dtbDetail.Rows(nData)
            Dim drValue As DataRow = dtbValue.Rows(nData)
            Dim strTester As String = drDetail.Item("tester")
            Dim strLot As String = drDetail.Item("Lot")
            Dim strSpec As String = drDetail.Item("Spec")
            drDetail.Item("test_time") = Format(CDate(drDetail.Item("test_time")), "yyyy-MM-dd HH:mm:ss")
            Dim strTestTime As String = drDetail.Item("test_time")
            Dim strTestTimeBigInt As String = Format(CDate(strTestTime), "yyyyMMddHHmmss")
            Dim strTagID As String = strTester & strTestTimeBigInt
            strDetailValue = strDetailValue & "'" & strTagID & "','" & strTestTimeBigInt & "',"
            strValue = strValue & "'" & strTagID & "','" & strTestTimeBigInt & "','" & strTestTime & "','" & strTester & "',"
            For nRow As Integer = 0 To drDetail.ItemArray.Length - 1
                strDetailValue = strDetailValue & "'" & drDetail.Item(nRow) & "',"
            Next
            For nRow As Integer = 0 To drValue.ItemArray.Length - 1
                strValue = strValue & "'" & drValue.Item(nRow) & "',"
            Next
            If Right(strDetailValue, 1) = "," Then strDetailValue = Left(strDetailValue, Len(strDetailValue) - 1) & ");"
            If Right(strValue, 1) = "," Then strValue = Left(strValue, Len(strValue) - 1) & ");"

            Dim strDateTmp() As String = strTestTime.Split(":")
            Dim strDateByHour As String = ""
            If strDateTmp.Length > 1 Then
                If CInt(strDateTmp(1)) < 30 Then
                    strDateByHour = strDateTmp(0) & ":00:00"
                Else
                    strDateByHour = strDateTmp(0) & ":30:00"
                End If
            End If

            Dim strTesterSQL As String = "REPLACE INTO db_" & strProduct & ".tabtester "
            strTesterSQL = strTesterSQL & "SELECT '" & strDateByHour & "','" & strTester & "','" & strLot & "','" & strSpec & "';"

            'clsMySQL.CommandNoQuery(strDetailValue & strValue, myRTTCConn)
        Next nData
    End Sub

End Class

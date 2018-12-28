Imports Microsoft.Office.Interop
Imports System.IO

Public Class CExcelManagment

    Public Enum enumAppend
        eVertical = 0
        eHorizontal
    End Enum

    Public Sub New()

    End Sub

    Private Function GetExcelColumnName(ByVal nCol As Integer) As String        'Start from 1
        Dim nCharA As Integer = Asc("A")
        Dim nMod As Integer = nCol Mod 26
        Dim nDiv As Integer = nCol \ 26

        Dim strLeft As String = ""
        Dim strRight As String = Chr(nMod + nCharA)
        If nDiv > 0 Then
            strLeft = Chr(nDiv + nCharA - 1)
        End If

        Dim strColName As String = strLeft & strRight
        GetExcelColumnName = strColName
    End Function

    Public Sub ExportDatatableToExcel(ByVal dtbData As DataTable, ByVal strFileName As String, ByVal bAppend As Boolean)
        Dim bExistFile As Boolean = False
        Dim appXL As Excel.Application = CreateObject("Excel.Application")
        Dim wbX1 As Excel.Workbook
        Dim raXL As Excel.Range

        If bAppend Then
            If File.Exists(strFileName) Then

                wbX1 = appXL.Workbooks.Open(strFileName)
                Dim bX As Boolean = wbX1.ReadOnly
                bExistFile = True
            Else
                wbX1 = appXL.Workbooks.Add
            End If
        Else
            If File.Exists(strFileName) Then
                File.Delete(strFileName)
            End If
            wbX1 = appXL.Workbooks.Add
        End If
        ' appXL.Visible = True
        Dim strMaxCol As String = GetExcelColumnName(dtbData.Columns.Count - 1)
        Dim bNewSheetAdd As Boolean = False
        Dim strSheetName As String = dtbData.TableName
        If Len(strSheetName) > 13 Then strSheetName = Left(strSheetName, 13)
        Dim shXL As Excel.Worksheet
        Try
            shXL = wbX1.Sheets(strSheetName)
        Catch ex As Exception
            bNewSheetAdd = True
            shXL = wbX1.Worksheets.Add()
            shXL.Name = strSheetName
        End Try

        If Not bExistFile Or bNewSheetAdd = True Then
            For nHeader As Integer = 0 To dtbData.Columns.Count - 1     'write header
                shXL.Range(GetExcelColumnName(nHeader) & "1").Value = dtbData.Columns(nHeader).ColumnName
            Next nHeader
        End If

        Dim nAppenAt As Integer = 2
        Dim strLastRow As String = dtbData.Rows(0).Item(0).ToString.ToUpper
        While shXL.Range("A" & nAppenAt).Text.ToString <> "" And Format(shXL.Range("A" & nAppenAt).Value, "dd-MMM-yyyy").ToUpper <> strLastRow
            nAppenAt = nAppenAt + 1
        End While


        shXL.Range("A1", strMaxCol & "1").Font.Bold = True
        shXL.Range("A1", strMaxCol & "1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


        For nRow As Integer = 0 To dtbData.Rows.Count - 1    'Write value
            Dim dtrData As DataRow = dtbData.Rows(nRow)
            shXL.Range("A" & nRow + nAppenAt, strMaxCol & (nRow + nAppenAt)).Value = dtrData.ItemArray
        Next nRow

        If bExistFile = False Then
            'For nSheet As Integer = 1 To wbX1.Worksheets.Count - 1
            '    Dim strSheet As String = "Sheet" & nSheet
            '    Dim Sheet As Excel.Worksheets = wbX1.Worksheets(strSheet)
            '    Sheet.Delete()
            'Next nSheet
            wbX1.SaveAs(strFileName)
        Else
            wbX1.Save()
        End If

        raXL = Nothing
        wbX1.Close()
        wbX1 = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub

    Public Sub ExportDatasetToExcel(ByVal dtsData As DataSet, ByVal strFileName As String, ByVal bAppend As Boolean)

        Dim bExistFile As Boolean = False
        Dim appXL As Excel.Application = CreateObject("Excel.Application")
        Dim wbX1 As Excel.Workbook
        Dim raXL As Excel.Range

        If bAppend Then
            If File.Exists(strFileName & ".xls") Then

                wbX1 = appXL.Workbooks.Open(strFileName & ".xls")
                Dim bX As Boolean = wbX1.ReadOnly
                bExistFile = True
            Else
                wbX1 = appXL.Workbooks.Add
            End If
        Else
            If File.Exists(strFileName) Then
                File.Delete(strFileName)
            End If
            wbX1 = appXL.Workbooks.Add
        End If
        ' appXL.Visible = True
        For nTable As Integer = 0 To dtsData.Tables.Count - 1
            Dim dtbData As DataTable = dtsData.Tables(nTable)
            Dim strMaxCol As String = GetExcelColumnName(dtbData.Columns.Count - 1)
            Dim bNewSheetAdd As Boolean = False
            Dim strSheetName As String = dtbData.TableName
            If Len(strSheetName) > 13 Then strSheetName = Left(strSheetName, 13)
            Dim shXL As Excel.Worksheet
            Try
                shXL = wbX1.Sheets(strSheetName)
            Catch ex As Exception
                bnewSheetAdd = True
                shXL = wbX1.Worksheets.Add()
                shXL.Name = strSheetName
            End Try

            'If bExistFile = True Then
            '    shXL = wbX1.Sheets(dtbData.TableName)
            'Else
            '    shXL = wbX1.Worksheets.Add()
            '    shXL.Name = dtbData.TableName
            'End If

            If Not bExistFile Or bnewSheetAdd = True Then
                For nHeader As Integer = 0 To dtbData.Columns.Count - 1     'write header
                    shXL.Range(GetExcelColumnName(nHeader) & "1").Value = dtbData.Columns(nHeader).ColumnName
                Next nHeader
            End If

            Dim nAppenAt As Integer = 2
            Dim strLastRow As String = dtbData.Rows(0).Item(0).ToString.ToUpper
            While shXL.Range("A" & nAppenAt).Text.ToString <> "" And Format(shXL.Range("A" & nAppenAt).Value, "dd-MMM-yyyy").ToUpper <> strLastRow
                nAppenAt = nAppenAt + 1
            End While


            shXL.Range("A1", strMaxCol & "1").Font.Bold = True
            shXL.Range("A1", strMaxCol & "1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


            For nRow As Integer = 0 To dtbData.Rows.Count - 1    'Write value
                Dim dtrData As DataRow = dtbData.Rows(nRow)
                shXL.Range("A" & nRow + nAppenAt, strMaxCol & (nRow + nAppenAt)).Value = dtrData.ItemArray
            Next nRow
        Next nTable

        If bExistFile = False Then
            Dim Sheet1 As Excel.Worksheet = wbX1.Worksheets("Sheet1")
            Sheet1.Delete()
            Dim Sheet2 As Excel.Worksheet = wbX1.Worksheets("Sheet2")
            Sheet2.Delete()
            Dim Sheet3 As Excel.Worksheet = wbX1.Worksheets("Sheet3")
            Sheet3.Delete()
            wbX1.SaveAs(strFileName)
        Else
            wbX1.Save()
        End If

        'appXL.SaveWorkspace(strFileName)
        'appXL.Visible = True
        'appXL.UserControl = True

        raXL = Nothing
        wbX1.Close()
        wbX1 = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub

    Public Sub SaveDataSetForMCDefect(ByVal dtsData As DataSet, ByVal strFileName As String, ByVal bAppend As Boolean, Optional ByVal nDirection As enumAppend = enumAppend.eVertical)

        Dim bExistFile As Boolean = False
        Dim appXL As Excel.Application = CreateObject("Excel.Application")
        Dim wbX1 As Excel.Workbook
        Dim raXL As Excel.Range

        If bAppend Then
            If File.Exists(strFileName & ".xls") Then

                wbX1 = appXL.Workbooks.Open(strFileName & ".xls")
                Dim bX As Boolean = wbX1.ReadOnly
                bExistFile = True
            Else
                wbX1 = appXL.Workbooks.Add
            End If
        Else
            If File.Exists(strFileName) Then
                File.Delete(strFileName)
            End If
            wbX1 = appXL.Workbooks.Add
        End If
        ' appXL.Visible = True
        For nTable As Integer = dtsData.Tables.Count - 1 To 0 Step -1
            Dim dtbData As New DataTable(dtsData.Tables(nTable).TableName)
            dtbData.Merge(dtsData.Tables(nTable))
            If bExistFile Then
                dtbData.Columns.Remove(dtbData.Columns(0).ColumnName)
            End If
            Dim strMaxCol As String = GetExcelColumnName(dtbData.Columns.Count - 1)

            Dim shXL As Excel.Worksheet
            If bExistFile = True Then
                shXL = wbX1.Sheets(dtbData.TableName)
            Else
                shXL = wbX1.Worksheets.Add()
                shXL.Name = dtbData.TableName
            End If
            Dim nAppenAt As Integer = 1
            If bExistFile = True Then
                Dim strDate As String = dtbData.Columns(0).ColumnName.ToUpper
                While shXL.Range(GetExcelColumnName(nAppenAt) & 1).Text.ToString <> "" And Format(shXL.Range(GetExcelColumnName(nAppenAt) & 1).Value, "dd-MMM-yyyy").ToUpper <> strDate
                    nAppenAt = nAppenAt + 1
                End While
            End If
            For nHeader As Integer = 0 To dtbData.Columns.Count - 1     'write header
                If bExistFile Then
                    shXL.Cells(1, nHeader + nAppenAt + 1) = dtbData.Columns(nHeader).ColumnName

                Else
                    shXL.Cells(1, nHeader + nAppenAt) = dtbData.Columns(nHeader).ColumnName

                End If
            Next nHeader

            shXL.Range("A1", GetExcelColumnName(nAppenAt) & "1").Font.Bold = True
            shXL.Range("A1", GetExcelColumnName(nAppenAt) & "1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            For nRow As Integer = 0 To dtbData.Rows.Count - 1    'Write value
                Dim dtrData As DataRow = dtbData.Rows(nRow)
                If bExistFile Then
                    shXL.Range(GetExcelColumnName(nAppenAt) & nRow + 2, GetExcelColumnName(nAppenAt + dtbData.Columns.Count - 1) & nRow + 2).Value = dtrData.ItemArray
                Else
                    shXL.Range("A" & nRow + 2, strMaxCol & nRow + 2).Value = dtrData.ItemArray
                End If
            Next nRow

        Next nTable
        If bExistFile = False Then
            Dim Sheet1 As Excel.Worksheet = wbX1.Worksheets("Sheet1")
            Sheet1.Delete()
            Dim Sheet2 As Excel.Worksheet = wbX1.Worksheets("Sheet2")
            Sheet2.Delete()
            Dim Sheet3 As Excel.Worksheet = wbX1.Worksheets("Sheet3")
            Sheet3.Delete()
            wbX1.SaveAs(strFileName)
        Else
            wbX1.Save()
        End If

        'appXL.SaveWorkspace(strFileName)
        'appXL.Visible = True
        'appXL.UserControl = True

        raXL = Nothing
        wbX1.Close()
        wbX1 = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub
End Class

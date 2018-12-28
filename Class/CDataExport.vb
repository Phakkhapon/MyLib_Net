Imports System.IO
Imports System.ComponentModel

Public Class CDataExport
    Private m_ProgressBar As Windows.Forms.ProgressBar
    Private Delegate Sub delegate_ProgressUpdate(ByVal paramValue As Integer, ByVal paramMax As Integer)
    Dim WithEvents bgw As New BackgroundWorker()

    Public Sub New()

    End Sub

    Public Function ReadCSV2Datatable(ByVal strCSVFilePath As String) As DataTable
        Dim strData() As String = File.ReadAllLines(strCSVFilePath)
        Dim dtbData As New DataTable(strCSVFilePath)
        Dim nStart As Integer = 0
        Dim strHeader() As String = Split(strData(nStart), ",")
        If strHeader.Length < 2 Then
            nStart = 1
            strHeader = Split(strData(nStart), ",")
        End If
        If strHeader.Length > 0 Then
            For nHeader As Integer = 0 To strHeader.Length - 1
                If dtbData.Columns(strHeader(nHeader)) Is Nothing Then
                    dtbData.Columns.Add(strHeader(nHeader))
                Else
                    dtbData.Columns.Add(strHeader(nHeader) & "_" & nHeader)
                End If
            Next nHeader
        End If

        For nData As Integer = nStart + 1 To strData.Length - 1
            Dim strValue() As String = Split(strData(nData), ",")
            If strValue.Length <= dtbData.Columns.Count And strData(nData) <> "" Then
                dtbData.Rows.Add(strValue)
            End If
        Next nData
        ReadCSV2Datatable = dtbData
    End Function


    Public Sub ExportDatatableToCSV(ByVal strFileSaveTo As String, ByVal dtbData As DataTable, ByVal bAppend As Boolean, Optional ByVal ProgressBar As Windows.Forms.ProgressBar = Nothing)

        If Not ProgressBar Is Nothing Then
            m_ProgressBar = ProgressBar
            m_ProgressBar.Visible = True
            m_ProgressBar.Maximum = dtbData.Rows.Count
            m_ProgressBar.Value = 0
            ' Dim bgw As New BackgroundWorker()
            AddHandler bgw.DoWork, AddressOf UpdateProgressbar
            bgw.RunWorkerAsync()

        End If
        Dim bWriteHeader As Boolean = True
        If bAppend = True And File.Exists(strFileSaveTo) = True Then
            bWriteHeader = False
        End If
        Dim strWriteHeader As String = ""
        For nCol As Integer = 0 To dtbData.Columns.Count - 1
            strWriteHeader = strWriteHeader & dtbData.Columns(nCol).ColumnName.Replace(",", ":") & ","
        Next nCol
        If Right(strWriteHeader, 1) = "," Then strWriteHeader = Left(strWriteHeader, strWriteHeader.Length - 1)
        Using swCSV As New IO.StreamWriter(strFileSaveTo, bAppend)
            If bWriteHeader Then swCSV.WriteLine(strWriteHeader)
            For nRow As Int64 = 0 To dtbData.Rows.Count - 1
                If Not m_ProgressBar Is Nothing Then
                    m_ProgressBar.Value = nRow
                    If Not bgw.IsBusy Then
                        bgw.RunWorkerAsync()
                    End If
                End If
                Dim drRow As DataRow = dtbData.Rows(nRow)
                Dim strDataTmp As String = ""
                For nCol As Integer = 0 To dtbData.Columns.Count - 1
                    strDataTmp = strDataTmp & drRow.Item(nCol).ToString.Replace(",", ":") & ","
                Next nCol
                If Right(strDataTmp, 1) = "," Then strDataTmp = Left(strDataTmp, strDataTmp.Length - 1)
                swCSV.WriteLine(strDataTmp)
            Next nRow
        End Using
        If Not m_ProgressBar Is Nothing Then m_ProgressBar.Visible = False
    End Sub

    Private Sub UpdateProgressbar() Handles bgw.DoWork
        If m_ProgressBar.InvokeRequired Then
            m_ProgressBar.BeginInvoke(New delegate_ProgressUpdate(AddressOf invokeMe_ProgressUpdate), m_ProgressBar.Value, m_ProgressBar.Maximum)
        Else
            invokeMe_ProgressUpdate(m_ProgressBar.Value, m_ProgressBar.Maximum)
        End If
    End Sub

    Private Sub invokeMe_ProgressUpdate(ByVal paramValue As Integer, ByVal paramMax As Integer)
        m_ProgressBar.Maximum = paramMax
        m_ProgressBar.Value = paramValue
        m_ProgressBar.Update()
    End Sub

End Class


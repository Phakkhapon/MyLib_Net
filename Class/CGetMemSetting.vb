Imports System.IO

Public Class CGetMemSetting
    Private m_strData() As String = Nothing
    Private m_strFileName As String = ""
    Private m_bIsWrite As Boolean = False
    Private m_dtsData As New DataSet

    Public Sub New(ByVal strFilePath As String)       'Read from file
        m_strFileName = strFilePath
        If File.Exists(strFilePath) Then
            m_strData = File.ReadAllLines(strFilePath)
            ConvertStringArrayToDataset(m_strData)
        End If
    End Sub

    Public Sub New(ByVal strData() As String, ByVal strFilePath As String)      'Read from array of data
        m_strFileName = strFilePath
        m_strData = strData
        ConvertStringArrayToDataset(m_strData)
    End Sub

    Public Sub New()
        ReDim m_strData(0)
        ConvertStringArrayToDataset(m_strData)
    End Sub

    Private Sub ConvertStringArrayToDataset(ByVal strData() As String)
        If strData Is Nothing Then Exit Sub
        Dim strSection As String = ""
        For nLine As Integer = 0 To strData.Length - 1
            Dim strLine As String = strData(nLine)
            If Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
                strSection = strLine.Replace("[", "").Replace("]", "")
                InitialSectionTable(strSection)
            Else
                If strSection <> "" Then
                    Dim strTemp() As String = Split(strLine, "=")
                    If strTemp.Length > 1 Then
                        Dim strKey As String = Trim(strTemp(0))
                        Dim strValue As String = Trim(Mid(strLine, strTemp(0).Length + 2))
                        If m_dtsData.Tables(strSection).Rows.Find(strKey) Is Nothing Then m_dtsData.Tables(strSection).Rows.Add(strKey, strValue)
                    End If
                End If
            End If
        Next nLine
    End Sub

    Private Sub InitialSectionTable(ByVal strSection As String)
        strSection = strSection.Replace("[", "").Replace("]", "")
        If m_dtsData.Tables(strSection) Is Nothing And strSection <> "" Then
            Dim dtbSection As New DataTable(strSection)
            Dim dcKey As DataColumn = dtbSection.Columns.Add("Key")
            dtbSection.Columns.Add("Value")
            Dim dcPrimKey(0) As DataColumn
            dcPrimKey(0) = dcKey
            dtbSection.PrimaryKey = dcPrimKey
            m_dtsData.Tables.Add(dtbSection)
        End If
    End Sub

    Public Function GetFileName() As String
        Return m_strFileName
    End Function

    Public Sub DeleteSection(ByVal strSection As String)
        If Not m_dtsData.Tables(strSection) Is Nothing Then
            m_dtsData.Tables.Remove(strSection)
        End If
    End Sub

    Public Function GetValueString(ByVal strSection As String, ByVal strKey As String, Optional ByVal strDefault As String = "") As String
        If m_dtsData.Tables(strSection) Is Nothing Then
            Return strDefault
        Else
            'Dim dv As New DataView(m_dtsData.Tables(strSection), Nothing, "Key", DataViewRowState.CurrentRows)
            'Dim drv() As DataRowView = dv.FindRows(strKey)
            'If drv.Length > 0 Then
            '    Dim strValue As String = drv(0).Item("Value")
            'End If

            Dim dtrValue As DataRow = m_dtsData.Tables(strSection).Rows.Find(strKey)
            If dtrValue IsNot Nothing Then
                Dim strValue As String = dtrValue.Item("Value").ToString
                If strValue = "" Then
                    Return strDefault
                Else
                    Return strValue
                End If
            Else
                Return strDefault
            End If
        End If
    End Function

    Public Function GetValueSection(ByVal strSection As String) As String()
        If m_dtsData.Tables(strSection) Is Nothing Then
            Return New String() {"[" & strSection & "]"}
        Else
            Dim dtbSection As DataTable = m_dtsData.Tables(strSection)
            Dim strValueSection(dtbSection.Rows.Count) As String
            For nData As Integer = 0 To dtbSection.Rows.Count
                If nData = 0 Then
                    strValueSection(0) = "[" & strSection & "]"
                Else
                    strValueSection(nData) = dtbSection.Rows(nData - 1).Item("Key").ToString & "=" & dtbSection.Rows(nData - 1).Item("Value").ToString
                End If
            Next
            Return strValueSection
        End If
    End Function

    Public Sub WriteValueSection(ByVal strSectionValue() As String, ByVal bAppend As Boolean)
        Dim strSection As String = strSectionValue(0)
        strSection = strSection.Replace("[", "").Replace("]", "")
        InitialSectionTable(strSection)
        If Not bAppend Then
            m_dtsData.Tables(strSection).Rows.Clear()
        End If

        For nLine As Integer = 1 To strSectionValue.Length - 1
            Dim strValue() As String = Split(strSectionValue(nLine), "=")
            If strValue.Length = 2 Then
                Dim drRow As DataRow = m_dtsData.Tables(strSection).Rows.Find(strValue(0))
                If drRow IsNot Nothing Then
                    drRow.Item("Value") = strValue(1)
                Else
                    m_dtsData.Tables(strSection).Rows.Add(strValue)
                End If
            End If
        Next nLine
    End Sub

    Public Function IsSection(ByVal strSection As String) As Boolean
        If m_dtsData.Tables(strSection) Is Nothing Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub ChangeSection(ByVal strSectionOld As String, ByVal strSectionNew As String)
        If strSectionNew.ToUpper <> strSectionOld.ToUpper And Not m_dtsData.Tables(strSectionOld) Is Nothing Then
            m_dtsData.Tables.Add(m_dtsData.Tables(strSectionOld).DefaultView.ToTable(strSectionNew))
            m_dtsData.Tables.Remove(strSectionOld)
        End If
    End Sub

    Public Sub WriteValueString(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String)

        InitialSectionTable(strSection)
        Dim dtbSection As DataTable = m_dtsData.Tables(strSection)
        Dim dtrRow As DataRow = dtbSection.Rows.Find(strKey)
        If dtrRow Is Nothing Then
            dtbSection.Rows.Add(strKey, strValue)
        Else
            dtrRow.Item("Value") = strValue
        End If
    End Sub

    Private Function ConvertDatasetToStringArray() As String()

        If m_dtsData.Tables.Count > 0 Then
            Dim strDataArray() As String = Nothing

            For nTable As Integer = 0 To m_dtsData.Tables.Count - 1
                Dim strSection As String = m_dtsData.Tables(nTable).TableName
                If strDataArray Is Nothing Then
                    ReDim Preserve strDataArray(0)
                Else
                    ReDim Preserve strDataArray(strDataArray.Length)
                End If
                strDataArray(strDataArray.Length - 1) = "[" & strSection & "]"
                For nData As Integer = 0 To m_dtsData.Tables(nTable).Rows.Count - 1
                    ReDim Preserve strDataArray(strDataArray.Length)
                    strDataArray(strDataArray.Length - 1) = m_dtsData.Tables(nTable).Rows(nData).Item("Key").ToString & "=" & m_dtsData.Tables(nTable).Rows(nData).Item("Value").ToString
                Next nData
            Next nTable
            Return strDataArray
        Else
            Return Nothing
        End If
    End Function

    Public Function GetINIData() As String()
        Return ConvertDatasetToStringArray()
    End Function

    Public Sub WriteBackAllString()
        Dim strData() As String = ConvertDatasetToStringArray()
        If m_bIsWrite = True And m_strFileName <> "" Then
            File.WriteAllLines(m_strFileName, strData)
        End If
    End Sub

    Public Sub WriteStringToFile(ByVal strFileName As String)
        Dim strData() As String = ConvertDatasetToStringArray()
        If strFileName <> "" And Not strData Is Nothing Then
            File.WriteAllLines(strFileName, strData)
        ElseIf strFileName <> "" And strData Is Nothing Then
            File.Delete(strFileName)
        End If
    End Sub

    Public ReadOnly Property GetDataset() As DataSet
        Get
            Return m_dtsData
        End Get
    End Property
    Protected Overrides Sub Finalize()
        'If m_bIsWrite = True And m_strFileName <> "" Then
        '    File.WriteAllLines(m_strFileName, m_strData)
        'End If
        m_dtsData.Dispose()
        MyBase.Finalize()
    End Sub
End Class

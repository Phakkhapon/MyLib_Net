Option Explicit On
Imports System.IO

Public Enum enumDataGroup
    enumHeader = 0
    enumData
    enumNorm
    enumGrade
End Enum

Public Class CWEXReader
    Private Const m_cstQuotation = """"
    Private Const HeaderDetail = 0
    Private Const DataHeader = 1
    Private Const HeaderParam = 2
    Private Const DataParamAvg = 3
    Private Const DataParamNrm = 4
    Private Const HeaderGrade = 5
    Private Const DataGrade = 6
    Private m_strFileName As String
    'Private  m_strGradeSetup(0) As String

    Private m_strCoreDataSetup1(6) As String
    Private m_strCoreDataSetup2(6) As String

    Private m_strHeaderArray(21) As String
    Private m_strAll() As String

    Public Sub New(ByVal strFileName As String)
        m_strFileName = strFileName
    End Sub

    Public Function ReadWexFile()
        Dim strWexFile As String = m_strFileName
        Dim sData As SRawData = Nothing
        If (File.Exists(strWexFile) = False) Then
            sData.bIsRead = False
            Return sData  'if file does't exist return data.bIsRead = false to function
        End If

        'Dim INIReader As New CIniFile(strWexFile)
        ''Set header data
        'data.sHeader.Product = INIReader.GetString(CombineResultW, Product, "")
        'data.sHeader.Station = INIReader.GetString(CombineResultW, Station, "")
        'data.sHeader.GradeName = INIReader.GetString(CombineResultW, GradeName, "")
        'data.sHeader.PartID = INIReader.GetString(CombineResultW, PartID, "")
        'data.sHeader.TestMode = INIReader.GetString(CombineResultW, TestMode, "")
        'data.sHeader.HeadSN = INIReader.GetString(CombineResultW, HeadSN, "")
        'data.sHeader.Alias_Product = INIReader.GetString(CombineResultW, Alias_Product, "")
        'data.sHeader.ShoeNo = INIReader.GetString(CombineResultW, ShoeW, "")
        'data.sHeader.SpinStand = INIReader.GetString(CombineResultW, SpinStandW, "")
        'data.sHeader.GradeRev = INIReader.GetString(CombineResultW, GradeinfoW, "")

        m_strAll = File.ReadAllLines(strWexFile, System.Text.UTF8Encoding.UTF8)
        'Dim strCoreData(6) As String
        For nCore As Integer = 0 To m_strCoreDataSetup1.Length - 1
            m_strCoreDataSetup1(nCore) = ""
        Next nCore

        'ReDim m_strGradeSetup(0) 'Set array to 0 dimension for initial

        'strCoreData(0) = strAll(1)  'Header detail
        'strCoreData(1) = strAll(2) 'header data detail
        'strCoreData(2) = strAll(4)  'data detail
        'strCoreData(3) = strAll(29)    'measure data
        'strCoreData(4) = strAll(30)       'normalize data
        m_strCoreDataSetup1 = GetCoreData("Setup1")
        m_strCoreDataSetup2 = GetCoreData("Setup2")
        If m_strCoreDataSetup1(3) = "" Then
            m_strCoreDataSetup1 = m_strCoreDataSetup2
        End If
        'If m_strCoreDataSetup1(3) = "" Then
        '    For nCore As Integer = 0 To m_strCoreDataSetup1.Length - 1    'clear data before read
        '        m_strCoreDataSetup1(nCore) = ""
        '    Next nCore
        '    GetCoreData("Setup2")
        'End If
        sData.bIsRead = True

        'GetRawData(strCoreData, sData)
        Return sData
    End Function

    Private Function GetCoreData(ByVal strSetup As String) As String()
        Dim strDataArray(6) As String
        For n As Integer = 0 To m_strAll.Length - 1
            m_strAll(n) = Replace(m_strAll(n), m_cstQuotation, "")
            Dim strDetail() As String = Split(m_strAll(n), ",")   'For Get Detail
            If strDetail.Length <= 3 And strDetail.Length > 1 Then
                If strDetail(HeaderDetail) = "Head #" Then
                    strDataArray(HeaderGrade) = m_strAll(n)
                ElseIf strDetail(0) = "-1" Then
                    strDataArray(DataGrade) = m_strAll(n)
                End If
            ElseIf strDetail.Length > 3 Then
                If strDetail(0) = "Start Time" Then
                    strDataArray(0) = m_strAll(n)
                    strDataArray(DataHeader) = Replace(m_strAll(n + 1), m_cstQuotation, "")
                ElseIf strDetail(0) = "ZName" And strDetail(5) = "Statistic Type" Then
                    strDataArray(HeaderParam) = m_strAll(n)
                ElseIf strDetail(1) = strSetup And strDetail(5) = "Avg" And (strDetail(6) = "0.0" Or strDetail(6) = "0") Then
                    strDataArray(DataParamAvg) = m_strAll(n)
                ElseIf strDetail(1) = strSetup And strDetail(5) = "Nrm" And (strDetail(6) = "0.0" Or strDetail(6) = "0") Then
                    strDataArray(DataParamNrm) = m_strAll(n)
                    'ElseIf strDetail(0) = "Grade Name" Then
                    '    For nGradeName As Integer = 0 To 25
                    '        If m_strAll(n + nGradeName) = "" Then
                    '            'sData.bIsRead = True
                    '            'GetRawData(m_strCoreData, sData)
                    '            'Return sData
                    '        End If
                    '        'ReDim Preserve m_strGradeSetup(nGradeName)
                    '        'm_strGradeSetup(nGradeName) = Replace(m_strAll(n + nGradeName), Quotation, "")
                    '    Next nGradeName
                End If
            End If
            If strDataArray(DataGrade) <> "" Then Exit For
        Next n
        Return strDataArray
    End Function

    Private Function GetRawData(ByVal strDataAll() As String, ByRef sData As SRawData)
        Dim strHeaderArray() As String = Split(strDataAll(1), ",")
        Dim strValueArray() As String = Split(strDataAll(3), ",")
        Dim strGradeArray() As String = Split(strDataAll(6), ",")
        With sData.sHeader
            If strHeaderArray.Length > 1 Then .dtStartTime = strHeaderArray(0)
            If strHeaderArray.Length > 1 Then .strMachine = strHeaderArray(1)
            If strHeaderArray.Length > 1 Then
                Dim strPartID() As String = Split(strHeaderArray(2), "/")
                sData.sHeader.strSpec = strPartID(0)
                sData.sHeader.strLotSN = strPartID(1)
                sData.sHeader.strAssySN = strPartID(2)
                sData.sHeader.strgr_config = strPartID(3)
            End If
            If strHeaderArray.Length > 1 Then .strOperatorID = strHeaderArray(3)
            If strValueArray.Length > 1 Then .strHeadSN = strValueArray(4)
            If strGradeArray.Length > 1 Then
                .strGradeName = strGradeArray(1)
                'If strGradeArray(2) = "Yes" Then
                '    .bIsPass = True
                'Else
                '    .bIsPass = False
                'End If
            End If
        End With

        Return True
    End Function

    Public Function ConvertToIniFormat() As String()

        Dim clsGetCombine As New CGetMemSetting(m_strAll, m_strFileName)
        Dim strMemINI(39) As String
        Dim strHead() As String = Split(m_strCoreDataSetup1(1), ",")
        Dim strDataHead() As String = Split(m_strCoreDataSetup1(2), ",")
        Dim strDataAvg() As String = Split(m_strCoreDataSetup1(3), ",")
        Dim strDataNrm() As String = Split(m_strCoreDataSetup1(4), ",")
        Dim strGrade() As String = Split(m_strCoreDataSetup1(6), ",")

        strMemINI(0) = "[Header]"
        strMemINI(1) = "StartTime=" & Format(CDate(strHead(0)), "yyyy-MM-dd HH:mm:ss")
        strMemINI(2) = "Station=" & strHead(1)
        strMemINI(3) = "Operator=" & strHead(3)
        strMemINI(4) = "Product=" & clsGetCombine.GetValueString("CombineResult", "Product")
        strMemINI(5) = "Alias=" & clsGetCombine.GetValueString("CombineResult", "Alias")
        strMemINI(6) = "TrayID=" & clsGetCombine.GetValueString("CombineResult", "HostTray")
        strMemINI(7) = "ProductPath="
        strMemINI(8) = "TestMode=" & clsGetCombine.GetValueString("CombineResult", "TestMode")
        'strMemINI(9) = "HeadSN=" & strDataAvg(4)  'Move to down to avoid error of no HeadSN data
        strMemINI(10) = "Machine=" & strHead(1)
        strMemINI(11) = "PartID=" & strHead(2)
        strMemINI(12) = "WorkType="
        strMemINI(13) = "MachineType="
        strMemINI(14) = "MonSpec="
        strMemINI(15) = "WTrayVersion=" & clsGetCombine.GetValueString("CombineResult", "WTrayVersion")
        strMemINI(16) = "GradeRev="
        strMemINI(17) = "Gradeinfo=" & clsGetCombine.GetValueString("CombineResult", "Gradeinfo")
        strMemINI(18) = "DISK Pack S/N=" & strHead(6)
        strMemINI(19) = "GradeName=" & strGrade(1)
        strMemINI(20) = "CartID=" & Split(Split(strHead(6), "/")(1), "_")(0)
        'Dim strShoe() As String = Split(strMemINI(18), "_")
        'strMemINI(21) = "ShoeNo=" & strShoe(strShoe.Length - 1)
        strMemINI(21) = "ShoeNo=" & clsGetCombine.GetValueString("CombineResult", "Shoe")

        strMemINI(22) = "GoodAfterAbort=" & clsGetCombine.GetValueString("CombineResult", "GoodAfterAbort")
        strMemINI(23) = "rHtrOhm=" & clsGetCombine.GetValueString("CombineResult", "rHtrOhm")
        strMemINI(24) = "rgLFTAA=" & clsGetCombine.GetValueString("CombineResult", "rgLFTAA")
        strMemINI(25) = "rTdType=" & clsGetCombine.GetValueString("CombineResult", "rTdType")
        strMemINI(26) = "rMWW=" & clsGetCombine.GetValueString("CombineResult", "rMWW")
        strMemINI(27) = "rTuMR_MRR=" & clsGetCombine.GetValueString("CombineResult", "rTuMR_MRR")
        strMemINI(28) = "rTdV=" & clsGetCombine.GetValueString("CombineResult", "rTdV")
        strMemINI(29) = "rMRRCheck=" & clsGetCombine.GetValueString("CombineResult", "rMRRCheck")
        strMemINI(30) = "rTdFreqH=" & clsGetCombine.GetValueString("CombineResult", "rTdFreqH")
        strMemINI(31) = "rTdAmpH=" & clsGetCombine.GetValueString("CombineResult", "rTdAmpH")
        strMemINI(32) = "rTrkPAmp=" & clsGetCombine.GetValueString("CombineResult", "rTrkPAmp")
        strMemINI(33) = "rPESAbort=" & clsGetCombine.GetValueString("CombineResult", "rPESAbort")
        strMemINI(34) = "rTC=" & clsGetCombine.GetValueString("CombineResult", "rTC")
        strMemINI(35) = "rWriterImpCheck=" & clsGetCombine.GetValueString("CombineResult", "rWriterImpCheck")
        strMemINI(36) = "rFaultCheck=" & clsGetCombine.GetValueString("CombineResult", "rFaultCheck")
        strMemINI(37) = "MEW_Abort=" & clsGetCombine.GetValueString("CombineResult", "MEW_Abort")
        strMemINI(38) = "GradeInfo=" & clsGetCombine.GetValueString("CombineResult", "GradeInfo")
        strMemINI(39) = "SanityFlag=" & clsGetCombine.GetValueString("CombineResult", "SanityFlag")

        If strDataAvg.Length = 1 Then Return strMemINI
        strMemINI(9) = "HeadSN=" & strDataAvg(4)

        '***************************Zone2 Setup1****************************
        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "[Avg_Zone2_Setup1]"
        If strDataAvg.Length > 1 Then
            For nAvg As Integer = 7 To strDataAvg.Length - 1
                ReDim Preserve strMemINI(strMemINI.Length)
                strMemINI(strMemINI.Length - 1) = strDataHead(nAvg) & "=" & strDataAvg(nAvg)
            Next nAvg
        End If
        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "ElapsedTime(sec)=" & strHead(4)

        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "[Nrm_Zone2_Setup1]"
        If strDataAvg.Length > 1 Then
            For nNrm As Integer = 7 To strDataAvg.Length - 1
                ReDim Preserve strMemINI(strMemINI.Length)
                If strDataNrm.Length = 1 Then
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataAvg(nNrm)
                ElseIf strDataNrm(nNrm) = "" Then
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataAvg(nNrm)
                Else
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataNrm(nNrm)
                End If
            Next nNrm
        End If
        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "ElapsedTime(sec)=" & strHead(4)

        '***************************Zone2 Setup2****************************
        strHead = Split(m_strCoreDataSetup2(1), ",")
        strDataHead = Split(m_strCoreDataSetup2(2), ",")
        strDataAvg = Split(m_strCoreDataSetup2(3), ",")
        strDataNrm = Split(m_strCoreDataSetup2(4), ",")
        strGrade = Split(m_strCoreDataSetup2(6), ",")

        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "[Avg_Zone2_Setup2]"
        If strDataAvg.Length > 1 Then
            For nAvg As Integer = 7 To strDataAvg.Length - 1
                ReDim Preserve strMemINI(strMemINI.Length)
                strMemINI(strMemINI.Length - 1) = strDataHead(nAvg) & "=" & strDataAvg(nAvg)
            Next nAvg
        End If
        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "ElapsedTime(sec)=" & strHead(4)

        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "[Nrm_Zone2_Setup2]"
        If strDataAvg.Length > 1 Then
            For nNrm As Integer = 7 To strDataAvg.Length - 1
                ReDim Preserve strMemINI(strMemINI.Length)
                If strDataNrm.Length = 1 Then
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataAvg(nNrm)
                ElseIf strDataNrm(nNrm) = "" Then
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataAvg(nNrm)
                Else
                    strMemINI(strMemINI.Length - 1) = strDataHead(nNrm) & "=" & strDataNrm(nNrm)
                End If
            Next nNrm
        End If
        ReDim Preserve strMemINI(strMemINI.Length)
        strMemINI(strMemINI.Length - 1) = "ElapsedTime(sec)=" & strHead(4)

        Return strMemINI
    End Function

End Class


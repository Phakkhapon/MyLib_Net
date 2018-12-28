Imports System.IO

Public Class CWEXConverter
    Private Const m_cstQuotation = """"
    Private m_dtbParaByProduct As DataTable
    Private m_strData() As String
    Private m_clsWexData As CGetMemSetting

    Public Sub New(ByVal strWexData() As String, ByVal dtbParaByProduct As DataTable)
        m_dtbParaByProduct = dtbParaByProduct
        m_strData = strWexData
        m_clsWexData = New CGetMemSetting(strWexData, "")
    End Sub

    Public Function GetWexData() As String()
        Return m_strData
    End Function

    Public Function ConvertToDTZValue() As CGetMemSetting
        Dim dtbHeader As DataTable = GetHeaderTable(m_strData)
        Dim dtsValue As DataSet = GetValueData(m_strData, m_dtbParaByProduct)

        ConvertToDTZValue = CreateINIFile(dtbHeader, dtsValue.Tables(0), dtsValue.Tables(1), dtsValue.Tables(2), dtsValue.Tables(3), dtsValue.Tables(4))
    End Function

    Private Function CreateINIFile(ByVal dtbHeader As DataTable, ByVal dtbValue As DataTable, ByVal dtbCFADD As DataTable, ByVal dtbCFMul As DataTable, ByVal dtbCFMedia As DataTable, ByVal dtbDeltaGOS As DataTable)
        Dim clsWexData As New CGetMemSetting(m_strData, Nothing)
        Dim clsINIData As New CGetMemSetting(Nothing, "")  'For create new INI data
        clsINIData.WriteValueString("Header", "StartTime", Format(CDate(dtbHeader.Rows(0).Item("Start Time")), "yyyy-MM-dd HH:mm:ss"))
        clsINIData.WriteValueString("Header", "Station", dtbHeader.Rows(0).Item("Station"))
        clsINIData.WriteValueString("Header", "Operator", dtbHeader.Rows(0).Item("Operator"))
        clsINIData.WriteValueString("Header", "Product", clsWexData.GetValueString("CombineResult", "Product"))
        clsINIData.WriteValueString("Header", "Alias", clsWexData.GetValueString("CombineResult", "Alias"))
        'clsINIData.WriteValueString("Header", "HeadSN", clsWexData.GetValueString("CombineResult", "HeadSN"))
        clsINIData.WriteValueString("Header", "HeadSN", dtbValue.Rows(0).Item("HeadSN"))
        clsINIData.WriteValueString("Header", "TrayID", dtbHeader.Rows(0).Item("Head Stack S/N").ToString.Split("/").GetValue(0).ToString) 'clsGetCombine.GetValueString("CombineResult", "HostTray"))
        clsINIData.WriteValueString("Header", "ProductPath", "")
        clsINIData.WriteValueString("Header", "TestMode", clsWexData.GetValueString("CombineResult", "TestMode"))
        Dim strModeTmp As String = clsINIData.GetValueString("Header", "TestMode")
        If strModeTmp = "3" Then
            clsINIData.WriteValueString("Header", "TestMode", "0")
        End If
        clsINIData.WriteValueString("Header", "Machine", dtbHeader.Rows(0).Item("Station"))
        clsINIData.WriteValueString("Header", "PartID", dtbHeader.Rows(0).Item("Part"))
        clsINIData.WriteValueString("Header", "WorkType", "")
        clsINIData.WriteValueString("Header", "MachineType", "")
        clsINIData.WriteValueString("Header", "MonSpec", "")
        clsINIData.WriteValueString("Header", "WTrayVersion", clsWexData.GetValueString("CombineResult", "WTrayVersion"))
        clsINIData.WriteValueString("Header", "GradeRev", "")
        clsINIData.WriteValueString("Header", "Gradeinfo", clsWexData.GetValueString("CombineResult", "Gradeinfo"))
        clsINIData.WriteValueString("Header", "DISK Pack S/N", dtbHeader.Rows(0).Item("Disk Pack S/N"))
        clsINIData.WriteValueString("Header", "GradeName", clsWexData.GetValueString("CombineResult", "GradeName"))
        Dim strCartID As String = dtbHeader.Rows(0).Item("Disk Pack S/N") & "/"
        clsINIData.WriteValueString("Header", "CartID", Split(Split(strCartID, "/")(1), "_")(0))
        clsINIData.WriteValueString("Header", "ShoeNo", clsWexData.GetValueString("CombineResult", "Shoe", "S1"))

        clsINIData.WriteValueString("Header", "rHtrOhm", clsWexData.GetValueString("CombineResult", "rHtrOhm"))
        clsINIData.WriteValueString("Header", "rgLFTAA", clsWexData.GetValueString("CombineResult", "rgLFTAA"))
        clsINIData.WriteValueString("Header", "rTdType", clsWexData.GetValueString("CombineResult", "rTdType"))
        clsINIData.WriteValueString("Header", "rMWW", clsWexData.GetValueString("CombineResult", "rMWW"))
        clsINIData.WriteValueString("Header", "rTuMR_MRR", clsWexData.GetValueString("CombineResult", "rTuMR_MRR"))
        clsINIData.WriteValueString("Header", "rTdV", clsWexData.GetValueString("CombineResult", "rTdV"))
        clsINIData.WriteValueString("Header", "rMRRCheck", clsWexData.GetValueString("CombineResult", "rMRRCheck"))
        clsINIData.WriteValueString("Header", "rTdFreqH", clsWexData.GetValueString("CombineResult", "rTdFreqH"))
        clsINIData.WriteValueString("Header", "rTdAmpH", clsWexData.GetValueString("CombineResult", "rTdAmpH"))
        clsINIData.WriteValueString("Header", "rTrkPAmp", clsWexData.GetValueString("CombineResult", "rTrkPAmp"))
        clsINIData.WriteValueString("Header", "rPESAbort", clsWexData.GetValueString("CombineResult", "rPESAbort"))
        clsINIData.WriteValueString("Header", "rTC", clsWexData.GetValueString("CombineResult", "rTC"))
        clsINIData.WriteValueString("Header", "rWriterImpCheck", clsWexData.GetValueString("CombineResult", "rWriterImpCheck"))
        clsINIData.WriteValueString("Header", "rFaultCheck", clsWexData.GetValueString("CombineResult", "rFaultCheck"))
        clsINIData.WriteValueString("Header", "MEW_Abort", clsWexData.GetValueString("CombineResult", "MEW_Abort"))
        clsINIData.WriteValueString("Header", "GradeInfo", clsWexData.GetValueString("CombineResult", "GradeInfo"))
        clsINIData.WriteValueString("Header", "SanityFlag", clsWexData.GetValueString("CombineResult", "SanityFlag"))
        clsINIData.WriteValueString("Header", "AbortToGood", clsWexData.GetValueString("CombineResult", "AbortToGood"))
        clsINIData.WriteValueString("Header", "GoodAfterAbort", clsWexData.GetValueString("CombineResult", "GoodAfterAbort"))

        If dtbValue.Columns.Count > 0 And dtbValue.Rows.Count > 0 Then
            For nParam As Integer = 1 To dtbValue.Columns.Count - 1
                Dim strParam As String = dtbValue.Columns(nParam).ColumnName
                Dim strValue As String = dtbValue.Rows(0).Item(strParam)
                clsINIData.WriteValueString("DTZValue", strParam, strValue)
            Next nParam
        End If
        Dim strTemp() As String = Split(dtbHeader.Rows(0).Item("Elapsed"), ":")
        Dim strTestTime As String = ""
        If strTemp.Length = 3 Then
            strTestTime = 3600 * CInt(strTemp(0)) + 60 * CInt(strTemp(1)) + CInt(strTemp(2))
        Else
            strTestTime = "0"
        End If
        clsINIData.WriteValueString("Header", "ElapsedTime(sec)", strTestTime)
        clsINIData.WriteValueString("DTZValue", "CycleTime", strTestTime)

        'Dim strCFAdd() As String = clsINIData.GetValueSection("CFADD")
        'Dim strCFMul() As String = clsINIData.GetValueSection("CFMUL")
        If dtbCFADD.Rows.Count = 1 Then
            For nCol As Integer = 0 To dtbCFADD.Columns.Count - 1
                clsINIData.WriteValueString("CFADD", dtbCFADD.Columns(nCol).ColumnName, dtbCFADD.Rows(0).Item(nCol).ToString)
            Next
        End If

        If dtbCFMul.Rows.Count = 1 Then
            For nCol As Integer = 0 To dtbCFMul.Columns.Count - 1
                clsINIData.WriteValueString("CFMUL", dtbCFMul.Columns(nCol).ColumnName, dtbCFMul.Rows(0).Item(nCol).ToString)
            Next
        End If

        If dtbCFMedia.Rows.Count = 1 Then
            For nCol As Integer = 0 To dtbCFMedia.Columns.Count - 1
                clsINIData.WriteValueString("CF_Media", dtbCFMedia.Columns(nCol).ColumnName, dtbCFMedia.Rows(0).Item(nCol).ToString)
            Next
        End If

        If dtbDeltaGOS.Rows.Count = 1 Then
            For nCol As Integer = 0 To dtbDeltaGOS.Columns.Count - 1
                clsINIData.WriteValueString("Delta_GOS", dtbDeltaGOS.Columns(nCol).ColumnName, dtbDeltaGOS.Rows(0).Item(nCol).ToString)
            Next
            Dim strGOSTime As String = m_clsWexData.GetValueString("Delta_GOS", "GOSTime")
            If strGOSTime <> "" Then
                clsINIData.WriteValueString("Delta_GOS", "GOSTime", strGOSTime)
            End If
            Dim strGOSLot As String = m_clsWexData.GetValueString("Delta_GOS", "Lot")
            If strGOSLot <> "" Then
                clsINIData.WriteValueString("Delta_GOS", "Lot", strGOSLot)
            End If
            Dim strGOSSpec As String = m_clsWexData.GetValueString("Delta_GOS", "Spec")
            If strGOSSpec <> "" Then
                clsINIData.WriteValueString("Delta_GOS", "Spec", strGOSSpec)
            End If
        End If
        CreateINIFile = clsINIData
    End Function

    Private Function GetHeaderTable(ByVal strWexData() As String) As DataTable
        Dim dtbHeader As New DataTable("Header")
        For nLine As Integer = 0 To strWexData.Length - 1
            Dim strData() As String = Split(Replace(strWexData(nLine), m_cstQuotation, ""), ",")
            If strData(0).ToUpper = "START TIME" Then
                Dim strrDetail() As String = Split(Replace(strWexData(nLine + 1), m_cstQuotation, ""), ",")
                For nCol As Integer = 0 To strData.Length - 1
                    If dtbHeader.Columns(strData(nCol)) Is Nothing Then
                        dtbHeader.Columns.Add(strData(nCol))
                    End If
                Next nCol
                dtbHeader.Rows.Add(strrDetail)
                Exit For
            End If
        Next nLine
        GetHeaderTable = dtbHeader
    End Function

    Private Function GetValueData(ByVal strWexData() As String, ByVal dtbParaByProduct As DataTable) As DataSet
        Dim dtbAllData As New DataTable()

        Dim dtbValue As New DataTable("Value")
        Dim dtbCFAdd As New DataTable("CFADD")
        Dim dtbCFMul As New DataTable("CFMUL")
        Dim dtbCFMedia As New DataTable("CF_Media")
        Dim dtbDeltaGOS As New DataTable("Delta_GOS")

        For nLine As Integer = 0 To strWexData.Length - 1
            Dim strData() As String = Split(Replace(strWexData(nLine), m_cstQuotation, ""), ",")
            If strData.Length > 3 Then
                If strData(0).ToUpper = "ZNAME" And strData(5).ToUpper = "STATISTIC TYPE" Then
                    Dim strrDetail() As String = Split(Replace(strWexData(nLine + 1), m_cstQuotation, ""), ",")
                    For nCol As Integer = 0 To strData.Length - 1
                        If dtbAllData.Columns(strData(nCol)) Is Nothing Then
                            dtbAllData.Columns.Add(strData(nCol))
                        End If
                    Next nCol
                ElseIf dtbAllData.Columns.Count = strData.Length Then 'Value
                    dtbAllData.Rows.Add(strData)
                End If
            End If
        Next nLine
        Dim dtbDataConf0 As DataTable = dtbAllData.Clone
        Dim drConf0() As DataRow = dtbAllData.Select("[Conf]='0.0' OR [Conf]='0'")
        For nConf As Integer = 0 To drConf0.Length - 1
            dtbDataConf0.Rows.Add(drConf0(nConf).ItemArray)
        Next nConf
        If dtbDataConf0.Rows.Count > 0 Then
            dtbValue.Columns.Add("HeadSN")
            dtbValue.Rows.Add()
            dtbValue.Rows(0).Item("HeadSN") = dtbDataConf0.Rows(0).Item("Head S/N")
            For nParam As Integer = 0 To dtbParaByProduct.Rows.Count - 1
                Dim dtrParam As DataRow = dtbParaByProduct.Rows(nParam)
                Dim strParam As String = dtrParam.Item("param_rttc")
                Dim strParamWex As String = dtrParam.Item("paramMachine")
                Dim strMachineCF As String = dtrParam.Item("MachineCF").ToString
                Dim strParamMDB As String = dtrParam.Item("ParamMDB").ToString

                If Not dtbDataConf0.Columns(strParamWex) Is Nothing Then
                    Dim strZone As String = dtrParam.Item("Zone")
                    If Len(strZone) = 1 Then strZone = "Zone" & strZone
                    Dim strSetup As String = dtrParam.Item("Setup")
                    If Len(strSetup) = 1 Then strSetup = "Setup" & strSetup
                    Dim strFilterAvg As String = "[ZName]='" & strZone & "' AND [Zone Setup]='" & strSetup & "' AND [Statistic Type]='Avg'"
                    Dim strFilterNrm As String = "[ZName]='" & strZone & "' AND [Zone Setup]='" & strSetup & "' AND [Statistic Type]='Nrm'"

                    Dim bIsNrm As Boolean = dtrParam.Item("Param_add") Or dtrParam.Item("Param_Mul")
                    Dim dtrSelect() As DataRow = Nothing
                    Dim dtrSelectAvg() As DataRow = Nothing
                    If bIsNrm = True Then
                        dtrSelect = dtbDataConf0.Select(strFilterNrm)
                        dtrSelectAvg = dtbDataConf0.Select(strFilterAvg)
                    Else
                        dtrSelect = dtbDataConf0.Select(strFilterAvg)
                    End If
                    If dtrSelect.Length > 0 Then
                        Dim strValue As String = dtrSelect(0).Item(strParamWex)
                        If strValue <> "" Then
                            dtbValue.Columns.Add(strParam)
                            dtbValue.Rows(0).Item(strParam) = strValue
                        End If
                        Dim strCFMedia As String = m_clsWexData.GetValueString("CF_Media", strParamWex & "_" & Right(strZone, 1))
                        If strCFMedia <> "" Then
                            dtbCFMedia.Columns.Add(strParam)
                            If dtbCFMedia.Rows.Count = 0 Then dtbCFMedia.Rows.Add()
                            dtbCFMedia.Rows(0).Item(strParam) = strCFMedia
                        End If

                        Dim strDeltaGOS As String = m_clsWexData.GetValueString("Delta_GOS", strParamWex)
                        If strDeltaGOS <> "" Then
                            dtbDeltaGOS.Columns.Add(strParam)
                            If dtbDeltaGOS.Rows.Count = 0 Then dtbDeltaGOS.Rows.Add()
                            dtbDeltaGOS.Rows(0).Item(strParam) = strDeltaGOS
                        End If

                        If dtrParam.Item("Param_add") = True Or strMachineCF <> "" Then   'Cal CF add
                            Dim strCFAdd As String = ""
                            If strMachineCF <> "" Then
                                strCFAdd = m_clsWexData.GetValueString("MCF_ADD", strMachineCF)
                            Else
                                strCFAdd = m_clsWexData.GetValueString("CFADD", strParamWex)
                            End If
                            If strCFAdd <> "" Then
                                dtbCFAdd.Columns.Add(strParam)
                                If dtbCFAdd.Rows.Count = 0 Then dtbCFAdd.Rows.Add()
                                dtbCFAdd.Rows(0).Item(strParam) = strCFAdd
                            Else
                                strCFAdd = m_clsWexData.GetValueString("CFADD", strParamMDB)
                                If strCFAdd <> "" Then
                                    dtbCFAdd.Columns.Add(strParam)
                                    If dtbCFAdd.Rows.Count = 0 Then dtbCFAdd.Rows.Add()
                                    dtbCFAdd.Rows(0).Item(strParam) = strCFAdd
                                End If
                            End If
                        End If
                        If dtrParam.Item("Param_Mul") = True Or strMachineCF <> "" Then   'Cal CF Mul
                            Dim strCFMul As String = ""
                            If strMachineCF <> "" Then
                                strCFMul = m_clsWexData.GetValueString("MCF_MUL", strMachineCF)
                            Else
                                strCFMul = m_clsWexData.GetValueString("CFMUL", strParamWex)
                            End If
                            If strCFMul <> "" Then
                                dtbCFMul.Columns.Add(strParam)
                                If dtbCFMul.Rows.Count = 0 Then dtbCFMul.Rows.Add()
                                dtbCFMul.Rows(0).Item(strParam) = strCFMul
                            Else
                                strCFMul = m_clsWexData.GetValueString("CFMUL", strParamMDB)
                                If strCFMul <> "" Then
                                    dtbCFMul.Columns.Add(strParam)
                                    If dtbCFMul.Rows.Count = 0 Then dtbCFMul.Rows.Add()
                                    dtbCFMul.Rows(0).Item(strParam) = strCFMul
                                End If
                            End If
                        End If
                    End If
                End If
            Next nParam
        End If
        Dim dtsValue As New DataSet
        dtsValue.Tables.Add(dtbValue)
        dtsValue.Tables.Add(dtbCFAdd)
        dtsValue.Tables.Add(dtbCFMul)
        dtsValue.Tables.Add(dtbCFMedia)
        dtsValue.Tables.Add(dtbDeltaGOS)
        GetValueData = dtsValue
    End Function

End Class


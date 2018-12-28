
Imports Lib_Net

Public Class CDTZConverter
    Private m_clsDataDCT As CGetMemSetting
    Private m_dtbParameter As DataTable
    Private m_clsDataDTZ As CGetMemSetting
    'Private m_dtbMCDefectMapping As DataTable

    Public Sub New(ByVal clsDCT As CGetMemSetting, ByVal dtbParameter As DataTable) ', ByVal dtbMCMapping As DataTable)
        m_clsDataDCT = clsDCT
        m_dtbParameter = dtbParameter
        'm_dtbMCDefectMapping = dtbMCMapping
        ' m_clsDataDCT.ChangeSection("Avg_1", "Avg_1_1")
        ' m_clsDataDCT.ChangeSection("Nrm_1", "Nrm_1_1")
        ' m_clsDataDCT.ChangeSection("Avg_2", "Avg_2_1")
        ' m_clsDataDCT.ChangeSection("Nrm_2", "Nrm_2_1")
    End Sub

    Public Function GetDTZData() As CGetMemSetting
        ConvertToDTZ()
        GetDTZData = m_clsDataDTZ
    End Function

    Private Sub ConvertToDTZ()
        Dim strCycleTime As String = m_clsDataDCT.GetValueString("Header", "CycleTime", "")
        Dim strElapsedTime As String = m_clsDataDCT.GetValueString("Header", "ElapsedTime(sec)")

        CreateDTZFile()            'Create initial DTZ file

        'm_clsDataDTZ.WriteValueString("Avg", "ElapsedTime(sec)", strElapsedTime)
        'm_clsDataDTZ.WriteValueString("DTZValue", "CycleTime", strCycleTime)

        For nData As Integer = 0 To m_dtbParameter.Rows.Count - 1
            Dim strParaName As String = m_dtbParameter.Rows(nData).Item("Param_rttc")
            Dim strParamWex As String = m_dtbParameter.Rows(nData).Item("parammachine")
            Dim strMachineCF As String = m_dtbParameter.Rows(nData).Item("MachineCF").ToString
            Dim strZoneSetup As String = Right(m_dtbParameter.Rows(nData).Item("Zone"), 1)

            Dim strValueAvg As String = m_clsDataDCT.GetValueString("Avg_" & strZoneSetup, strParamWex)       'Get Value Avg
            If strValueAvg = "" Then strValueAvg = m_clsDataDCT.GetValueString("Header", strParamWex) 'If can not get data, Get from header

            Dim strValueNrm As String = m_clsDataDCT.GetValueString("Nrm_" & strZoneSetup, strParamWex)       'Get Value Nrm
            If strValueNrm = "" Then strValueNrm = m_clsDataDCT.GetValueString("Header", strParamWex) 'If can not get data, Get from header

            Dim bCFAdd As Boolean = m_dtbParameter.Rows(nData).Item("param_add")
            Dim bCFMul As Boolean = m_dtbParameter.Rows(nData).Item("param_mul")

            If bCFAdd = False And bCFMul = False Then
                If strValueAvg <> "" Then m_clsDataDTZ.WriteValueString("DTZValue", strParaName, strValueAvg) 'Write value to DTZ file
            ElseIf strParamWex <> "" Or strMachineCF <> "" Then
                If strValueNrm <> "" Then m_clsDataDTZ.WriteValueString("DTZValue", strParaName, strValueNrm) 'Write value to DTZ file
                Dim strCFMedia As String = m_clsDataDCT.GetValueString("CF_Media", strParamWex & "_" & strZoneSetup)
                If strCFMedia <> "" Then
                    m_clsDataDTZ.WriteValueString("CF_Media", strParaName, strCFMedia)
                End If

                Dim strDeltaGOS As String = m_clsDataDCT.GetValueString("Delta_GOS", strParamWex)
                If strDeltaGOS <> "" Then
                    m_clsDataDTZ.WriteValueString("Delta_GOS", strParaName, strDeltaGOS)
                End If

                If bCFAdd Then
                    Dim strCF As String
                    If strMachineCF <> "" Then
                        strCF = m_clsDataDCT.GetValueString("MCF_ADD", strMachineCF, "")
                    Else
                        strCF = m_clsDataDCT.GetValueString("CFADD", strParamWex)
                    End If
                    If strCF <> "" Then m_clsDataDTZ.WriteValueString("CFADD", strParaName, strCF) 'Write value to DTZ file
                End If
                If bCFMul Then
                    Dim strCF As String
                    If strMachineCF <> "" Then
                        strCF = m_clsDataDCT.GetValueString("MCF_MUL", strMachineCF, "")
                    Else
                        strCF = m_clsDataDCT.GetValueString("CFMUL", strParamWex)
                    End If
                    If strCF <> "" Then m_clsDataDTZ.WriteValueString("CFMUL", strParaName, strCF) 'Write value to DTZ file
                End If
            End If
        Next nData
        Dim strGOSTime As String = m_clsDataDCT.GetValueString("Delta_GOS", "GOSTime")
        If strGOSTime <> "" Then
            m_clsDataDTZ.WriteValueString("Delta_GOS", "GOSTime", strGOSTime)
        End If
        Dim strGOSLot As String = m_clsDataDCT.GetValueString("Delta_GOS", "Lot")
        If strGOSLot <> "" Then
            m_clsDataDTZ.WriteValueString("Delta_GOS", "Lot", strGOSLot)
        End If
        Dim strGOSSpec As String = m_clsDataDCT.GetValueString("Delta_GOS", "Spec")
        If strGOSSpec <> "" Then
            m_clsDataDTZ.WriteValueString("Delta_GOS", "Spec", strGOSSpec)
        End If
    End Sub

    Private Sub CreateDTZFile()  'Create initial DTZ files
        Dim strTemp() As String = Nothing
        m_clsDataDTZ = New CGetMemSetting(strTemp, "")
        Dim strHead() As String = m_clsDataDCT.GetValueSection("Header")
        m_clsDataDTZ.WriteValueSection(strHead, False)
        Dim strOutTrayID As String = m_clsDataDTZ.GetValueString("Header", "OutTrayID")
        Dim strTrayIDOut As String = m_clsDataDTZ.GetValueString("Header", "TrayIDOut")
        If strTrayIDOut = "" And strOutTrayID <> "" Then
            m_clsDataDTZ.WriteValueString("Header", "TrayIDOut", strOutTrayID)
        End If
    End Sub

End Class

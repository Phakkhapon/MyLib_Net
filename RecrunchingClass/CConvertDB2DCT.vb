Imports System.IO

Public Class CConvertDB2DCT
    Private m_strProduct As String
    Private m_dtbRawData As DataTable
    Private m_dtbHeader As DataTable
    Private m_dtbParam As DataTable

    Public Sub New(ByVal strProduct As String, ByVal dtbRawData As DataTable, ByVal dtbHeader As DataTable, ByVal dtbParamByProduct As DataTable)
        m_dtbRawData = dtbRawData
        m_dtbHeader = dtbHeader
        m_dtbParam = dtbParamByProduct
        m_strProduct = strProduct
    End Sub

    Public Function GetAllConversionFile() As CGetMemSetting()
        Dim clsDCT(m_dtbRawData.Rows.Count - 1) As CGetMemSetting
        For nData As Integer = 0 To m_dtbRawData.Rows.Count - 1
            clsDCT(nData) = ConvertDB2DCT(m_dtbRawData.Rows(nData))
        Next nData
        GetAllConversionFile = clsDCT
    End Function

    Public Function ConvertDB2DCT(ByVal drRawData As DataRow) As CGetMemSetting
        Dim clsDCTINI As New CGetMemSetting
        Dim dtbRawData As DataTable = drRawData.Table
        For nHeader As Integer = 0 To m_dtbHeader.Rows.Count - 1
            Dim strHeaderWex As String = m_dtbHeader.Rows(nHeader).Item("Key")
            Dim strHeaderDB As String = m_dtbHeader.Rows(nHeader).Item("HeaderName")
            Dim strSection As String = m_dtbHeader.Rows(nHeader).Item("Section")
            If dtbRawData.Columns(strHeaderDB) IsNot Nothing Then
                Dim strHeaderValue As String = ""
                If dtbRawData.Columns(strHeaderDB).DataType Is GetType(DateTime) Then
                    strHeaderValue = Format(drRawData.Item(strHeaderDB), "yyyy-MM-dd HH:mm:ss")
                Else
                    strHeaderValue = drRawData.Item(strHeaderDB).ToString
                End If
                clsDCTINI.WriteValueString(strSection, strHeaderWex, strHeaderValue)
            End If
        Next nHeader
        clsDCTINI.WriteValueString("Header", "Product", m_strProduct.Split("_")(0))
        clsDCTINI.WriteValueString("Header", "Alias", m_strProduct)
        clsDCTINI.WriteValueString("Header", "Machine", clsDCTINI.GetValueString("Header", "Station"))
        Dim strSpec As String = clsDCTINI.GetValueString("Header", "Spec")
        Dim strLot As String = clsDCTINI.GetValueString("Header", "Lot")
        Dim strAssy As String = clsDCTINI.GetValueString("Header", "Assy")
        Dim strWorkID As String = clsDCTINI.GetValueString("Header", "WorkID")
        Dim strPartNumber As String = clsDCTINI.GetValueString("Header", "PartNumber")
        Dim strSliderSite As String = clsDCTINI.GetValueString("Header", "SliderSite")

        Dim strPathID As String = strSpec & "/" & strLot & "/" & strAssy & "/" & strWorkID & "/" & strPartNumber & "/" & strSliderSite & "/"
        clsDCTINI.WriteValueString("Header", "PartID", strPathID)
        Dim strShoe As String = clsDCTINI.GetValueString("Header", "ShoeNum")
        If Left(strShoe, 1) <> "S" Then strShoe = "S" & strShoe
        clsDCTINI.WriteValueString("Header", "ShoeNo", strShoe)

        Dim strMediaSN As String = clsDCTINI.GetValueString("Header", "MediaSN")
        Dim strMediaSurface As String = clsDCTINI.GetValueString("Header", "MediaSurface")
        Dim strMediaCount As String = clsDCTINI.GetValueString("Header", "MediaCount")
        Dim strDiskPack As String = strMediaSN & "_" & strMediaSurface & "_" & strMediaCount & "/99991_" & strShoe
        clsDCTINI.WriteValueString("Header", "DISK Pack S/N", strDiskPack)
        Dim strTrayIDOut As String = clsDCTINI.GetValueString("Header", "TrayIDOut")
        clsDCTINI.WriteValueString("Header", "OutTrayID", strTrayIDOut)
        clsDCTINI.WriteValueString("Header", "HGAPosition", clsDCTINI.GetValueString("Header", "SliderPosition"))

        Dim strZone As String = ""
        For nParam As Integer = 0 To m_dtbParam.Rows.Count - 1
            Dim drParam As DataRow = m_dtbParam.Rows(nParam)
            Dim strParamRTTC As String = drParam.Item("param_rttc")
            Dim strParamWex As String = drParam.Item("paramMachine")
            Dim strParamDisplay As String = drParam.Item("param_display")
            strZone = drParam.Item("Zone")
            strZone = Right(strZone, 1)
            If dtbRawData.Columns(strParamDisplay) IsNot Nothing Then
                Dim strValue As String = drRawData.Item(strParamDisplay).ToString
                If IsNumeric(strValue) Then
                    Dim dblCFAdd As Double = 0
                    Dim dblCFMul As Double = 1
                    If dtbRawData.Columns(strParamDisplay & ".CFAdd") IsNot Nothing Then
                        If drRawData.Item(strParamDisplay & ".CFAdd") IsNot DBNull.Value Then
                            dblCFAdd = drRawData.Item(strParamDisplay & ".CFAdd")
                            clsDCTINI.WriteValueString("CFAdd", strParamWex, dblCFAdd)
                        End If
                    End If
                    If dtbRawData.Columns(strParamDisplay & ".CFMul") IsNot Nothing Then
                        If drRawData.Item(strParamDisplay & ".CFMul") IsNot DBNull.Value Then
                            dblCFMul = drRawData.Item(strParamDisplay & ".CFMul")
                            clsDCTINI.WriteValueString("CFMul", strParamWex, dblCFMul)
                        End If
                    End If
                    Dim dblValue As Double = strValue
                    Dim dblAvg As Double = 0
                    Dim dblNrm As Double = 0
                    dblNrm = dblValue
                    dblAvg = (dblNrm - dblCFAdd) / dblCFMul
                    clsDCTINI.WriteValueString("Avg_" & strZone, strParamWex, dblAvg)
                    clsDCTINI.WriteValueString("Nrm_" & strZone, strParamWex, dblNrm)
                ElseIf strValue <> "" Then
                    clsDCTINI.WriteValueString("Avg_" & strZone, strParamWex, strValue)
                End If
            End If
        Next nParam
        clsDCTINI.WriteValueString("Avg_" & strZone, "ElapsedTime(sec)", clsDCTINI.GetValueString("Header", "ElapsedTime(sec)"))
        clsDCTINI.WriteValueString("Nrm_" & strZone, "ElapsedTime(sec)", clsDCTINI.GetValueString("Header", "ElapsedTime(sec)"))
        ConvertDB2DCT = clsDCTINI
    End Function

    Public Function ConvertDB2DCT(ByVal strProduct As String, ByVal drRawData As DataRow, ByVal dtbHeader As DataTable, ByVal dtbParam As DataTable) As CGetMemSetting
        Dim clsDCTINI As New CGetMemSetting()
        Dim dtbRawData As DataTable = drRawData.Table
        For nHeader As Integer = 0 To dtbHeader.Rows.Count - 1
            Dim strHeaderWex As String = dtbHeader.Rows(nHeader).Item("Key")
            Dim strHeaderDB As String = dtbHeader.Rows(nHeader).Item("HeaderName")
            Dim strSection As String = dtbHeader.Rows(nHeader).Item("Section")
            If dtbRawData.Columns(strHeaderDB) IsNot Nothing Then
                Dim strHeaderValue As String = ""
                If dtbRawData.Columns(strHeaderDB).DataType Is GetType(DateTime) Then
                    strHeaderValue = Format(drRawData.Item(strHeaderDB), "yyyy-MM-dd HH:mm:ss")
                Else
                    strHeaderValue = drRawData.Item(strHeaderDB).ToString
                End If
                clsDCTINI.WriteValueString(strSection, strHeaderWex, strHeaderValue)
            End If
        Next nHeader
        clsDCTINI.WriteValueString("Header", "Product", strProduct.Split("_")(0))
        clsDCTINI.WriteValueString("Header", "Alias", strProduct)
        clsDCTINI.WriteValueString("Header", "Machine", clsDCTINI.GetValueString("Header", "Station"))
        Dim strSpec As String = clsDCTINI.GetValueString("Header", "Spec")
        Dim strLot As String = clsDCTINI.GetValueString("Header", "Lot")
        Dim strAssy As String = clsDCTINI.GetValueString("Header", "Assy")
        Dim strWorkID As String = clsDCTINI.GetValueString("Header", "WorkID")
        Dim strPartNumber As String = clsDCTINI.GetValueString("Header", "PartNumber")
        Dim strSliderSite As String = clsDCTINI.GetValueString("Header", "SliderSite")

        Dim strPathID As String = strSpec & "/" & strLot & "/" & strAssy & "/" & strWorkID & "/" & strPartNumber & "/" & strSliderSite & "/"
        clsDCTINI.WriteValueString("Header", "PartID", strPathID)
        Dim strShoe As String = clsDCTINI.GetValueString("Header", "ShoeNum")
        If Left(strShoe, 1) <> "S" Then strShoe = "S" & strShoe
        clsDCTINI.WriteValueString("Header", "ShoeNo", strShoe)

        Dim strMediaSN As String = clsDCTINI.GetValueString("Header", "MediaSN")
        Dim strMediaSurface As String = clsDCTINI.GetValueString("Header", "MediaSurface")
        Dim strMediaCount As String = clsDCTINI.GetValueString("Header", "MediaCount")
        Dim strDiskPack As String = strMediaSN & "_" & strMediaSurface & "_" & strMediaCount & "/99991_" & strShoe
        clsDCTINI.WriteValueString("Header", "DISK Pack S/N", strDiskPack)
        Dim strTrayIDOut As String = clsDCTINI.GetValueString("Header", "TrayIDOut")
        clsDCTINI.WriteValueString("Header", "OutTrayID", strTrayIDOut)
        Dim strSldPosition As String = clsDCTINI.GetValueString("Header", "SliderPosition")
        clsDCTINI.WriteValueString("Header", "HGAPosition", strSldPosition)
        clsDCTINI.WriteValueString("Header", "OutSliderPosition", strSldPosition)

        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim drParam As DataRow = dtbParam.Rows(nParam)
            Dim strParamRTTC As String = drParam.Item("param_rttc")
            Dim strParamWex As String = drParam.Item("paramMachine")
            Dim strParamDisplay As String = drParam.Item("param_display")
            Dim strZone As String = drParam.Item("Zone")
            strZone = Right(strZone, 1)
            If dtbRawData.Columns(strParamDisplay) IsNot Nothing Then
                Dim strValue As String = drRawData.Item(strParamDisplay).ToString
                If IsNumeric(strValue) Then
                    Dim dblCFAdd As Double = 0
                    Dim dblCFMul As Double = 1
                    If dtbRawData.Columns(strParamDisplay & ".CFAdd") IsNot Nothing Then
                        If drRawData.Item(strParamDisplay & ".CFAdd") IsNot DBNull.Value Then
                            dblCFAdd = drRawData.Item(strParamDisplay & ".CFAdd")
                            clsDCTINI.WriteValueString("CFAdd", strParamWex, dblCFAdd)
                        End If
                    End If
                    If dtbRawData.Columns(strParamDisplay & ".CFMul") IsNot Nothing Then
                        If drRawData.Item(strParamDisplay & ".CFMul") IsNot DBNull.Value Then
                            dblCFMul = drRawData.Item(strParamDisplay & ".CFMul")
                            clsDCTINI.WriteValueString("CFMul", strParamWex, dblCFMul)
                        End If
                    End If
                    Dim dblValue As Double = strValue
                    Dim dblAvg As Double = 0
                    Dim dblNrm As Double = 0
                    dblNrm = dblValue
                    dblAvg = (dblNrm - dblCFAdd) / dblCFMul
                    clsDCTINI.WriteValueString("Avg_" & strZone, strParamWex, dblAvg)
                    clsDCTINI.WriteValueString("Nrm_" & strZone, strParamWex, dblNrm)
                ElseIf strValue <> "" Then
                    clsDCTINI.WriteValueString("Avg_" & strZone, strParamWex, strValue)
                End If
            End If
        Next nParam
        ConvertDB2DCT = clsDCTINI
    End Function



End Class

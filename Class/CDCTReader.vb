Option Explicit On
Imports System.IO

Public Structure SShowAfterAbort
    Dim strGoodAF_Check As String
    Dim strGoodAF_rMWW As String
    Dim strGoodAF_rTuMRR As String
    Dim strGoodAF_rLF As String
    Dim strGoodAF_rvPtd As String
    Dim strGoodAF_rHtrOhm As String
    Dim strGoodAF_rTdType As String

    Dim strGoodAF_rMRRCheck As String
    Dim strGoodAF_rTdFreqH As String
    Dim strGoodAF_rTdAmpH As String
    Dim strGoodAF_rTrkPAmp As String
    Dim strGoodAF_rPESAbort As String
    Dim strGoodAF_rTC As String
    Dim strGoodAF_rWriterImpCheck As String
    Dim strGoodAF_rFaultCheck As String
    Dim strGoodAF_MEW_Abort As String
    Dim strGoodAF_SanityFlag As String
    Dim strGrRev As String
End Structure

Public Structure SMeasureData
    Dim dMRR_Ohm As Single
    Dim dIBias_mA As Single
    Dim dWRO_uIn As Single
    Dim ATTn As Single
    Dim dTdV As Single
    Dim dTdP As Single
    Dim HtrOhm As Single
    Dim dSampl As Single
    Dim dSFreq As Single
    Dim dHampl As Single
    Dim dHFreq As Single
    Dim dLampl As Single
    Dim dLFreq As Single
    Dim dTdType As Single
    Dim PTD_vDfhRd As Single
    Dim PTD_vDfhWr As Single
    Dim PTD_ReloadFlag As Single
    Dim dWROffset_uIn As Single
    Dim dMWW_uIn As Single
    Dim dMRW_uIn As Single
    Dim dDeltaWidth_uIn As Single
    Dim gfTAA_mV As Single
    Dim R_Slope As Single
    Dim L_Slope As Single
    Dim ServoBW2 As Single
    Dim LeftToCenter As Single
    Dim RightToCenter As Single
    Dim CenterCor As Single
    Dim WSO As Single
    Dim RSO As Single
    Dim dmewWRO_uIn As Single
    Dim dMEW_uIn As Single
    Dim MEWRetest As Single
    Dim dLF_TAA_mV As Single
    Dim dEMF_LFTAA_mV As Single
    Dim dEMF_LFTAA_in_mV As Single
    Dim dHF_TAA_mV As Single
    Dim dEMF_HFTAA_mV As Single
    Dim dEMF_HFTAA_in_mV As Single
    Dim dResolution As Single
    Public dTaaAsym As Single
    Dim dOW2_dB As Single
    Dim Retest As Single
    Dim ElapsedTime_sec As Single
    Dim dMRR2_Ohm As Single
    Dim dDeltaMRR_Ohm As Single
End Structure

Public Structure SHeaderData
    Dim dtStartTime As Date '=2011-07-18 07:57:55
    Dim strMachine As String ' = 5121
    Dim strProduct As String '= SHASTA
    Dim strAlias_Product As String '=SHASTA_XLOT_DCT
    Dim strSpecRev As String 'reference spec
    Dim strLotSN As String 'Lot serial number
    Dim strAssySN As String 'Assy serial number
    Dim strMediaSN As String 'media serial number
    Dim strgr_config As String
    Dim strGradeRev As String '= AA
    Dim strShoeNo As String 'S1 or Shoe2
    Dim strOperatorID As String '=41083
    Dim strTrayID As String ' = ER1304566Q
    Dim strProductPath As String '=C:\WD_DCT400\SHASTA_DCT_XLOT
    Dim strTestMode As String '= 0
    Dim strHeadSN As String ' = PF11S07P
    Dim strSpec As String ''CMZA / JUN11D11
    Dim strWorkType As String ' = PRIME
    Dim strMachineType As String '= A
    Dim strWTrayVersion As String 'V 1.0.75.7
    Dim strGradeName As String '=PASS_BIN1(PRIME_DCT400)_REV.AA
    Dim strCartID As String '001462C13F
    Dim strSpinStand As String 'spinstand
    Dim strElapsedTime As String 'ElapsedTime
End Structure

Public Structure SRawData
    Dim sHeader As SHeaderData
    Dim sAVGData As SMeasureData
    Dim sNormData As SMeasureData
    Dim sAbort As SShowAfterAbort
    Dim bIsRead As Boolean
End Structure

Public Class CDCTReader
    Private m_strDCTFile As String
    Private m_CINIReader As CGetMemSetting
    Public Sub New(ByVal strDCTFile As String)
        m_strDCTFile = strDCTFile
    End Sub
    Public Function GetDCTData() As SRawData
        Dim data As SRawData = Nothing
        'If (File.Exists(m_strDCTFile) = False) Then
        '    data.bIsRead = False
        '    GetDCTData = data 'if file does't exist return data.bIsRead = false to function
        'End If

        'm_CINIReader = New CGetMemSetting(m_strDCTFile)

        ''read header data
        'data.sHeader.dtStartTime = m_CINIReader.GetValueString(Header, StartTime)
        'data.sHeader.strMachine = m_CINIReader.GetValueString(Header, Station)
        'data.sHeader.strOperatorID = m_CINIReader.GetValueString(Header, OperatorID)
        'data.sHeader.strProduct = m_CINIReader.GetValueString(Header, Product)
        'data.sHeader.strAlias_Product = m_CINIReader.GetValueString(Header, Alias_Product)
        'data.sHeader.strTrayID = m_CINIReader.GetValueString(Header, TrayID)
        'data.sHeader.strProductPath = m_CINIReader.GetValueString(Header, ProductPath)
        'data.sHeader.strTestMode = m_CINIReader.GetValueString(Header, TestMode)
        'data.sHeader.strHeadSN = m_CINIReader.GetValueString(Header, HeadSN)
        'data.sHeader.strWorkType = m_CINIReader.GetValueString(Header, WorkType)
        'data.sHeader.strMachineType = m_CINIReader.GetValueString(Header, MachineType)
        'data.sHeader.strWTrayVersion = m_CINIReader.GetValueString(Header, WTrayVersion)
        'data.sHeader.strGradeRev = m_CINIReader.GetValueString(Header, GradeRev)
        'data.sHeader.strGradeName = m_CINIReader.GetValueString(Header, GradeName)
        'data.sHeader.strCartID = m_CINIReader.GetValueString(Header, CartID)
        'data.sHeader.strShoeNo = m_CINIReader.GetValueString(Header, ShoeNo)
        'data.sHeader.strElapsedTime = m_CINIReader.GetValueString(AVG2, ElapsedTime)
        'Dim strPartID() As String = Split(m_CINIReader.GetValueString(Header, PartID), "/")
        'data.sHeader.strSpec = strPartID(0)
        'data.sHeader.strLotSN = strPartID(1)
        'data.sHeader.strAssySN = strPartID(2)
        'data.sHeader.strgr_config = strPartID(3)
        'data.sHeader.strSpec = Split(m_CINIReader.GetValueString(Header, MonSpec), "/")(0)
        'data.sHeader.strMediaSN = Split(Split(m_CINIReader.GetValueString(Header, DISK_Pack_SN), "/")(0), "_")(0)

        ''read avg data
        'data.sAVGData.ATTn = m_CINIReader.GetValueString(AVG2, ATTn, "0")
        'data.sAVGData.CenterCor = m_CINIReader.GetValueString(AVG2, CenterCor, "0")
        'data.sAVGData.dDeltaMRR_Ohm = m_CINIReader.GetValueString(AVG2, dDeltaMRR_Ohm, "0")
        'data.sAVGData.dDeltaWidth_uIn = m_CINIReader.GetValueString(AVG2, dDeltaWidth_uIn, "0")
        'data.sAVGData.dEMF_HFTAA_in_mV = m_CINIReader.GetValueString(AVG2, dEMF_HFTAA_in_mV, "0")
        'data.sAVGData.dEMF_HFTAA_mV = m_CINIReader.GetValueString(AVG2, dEMF_HFTAA_mV, "0")
        'data.sAVGData.dEMF_LFTAA_in_mV = m_CINIReader.GetValueString(AVG2, dEMF_LFTAA_in_mV, "0")
        'data.sAVGData.dEMF_LFTAA_mV = m_CINIReader.GetValueString(AVG2, dEMF_LFTAA_mV, "0")
        'data.sAVGData.dHampl = m_CINIReader.GetValueString(AVG2, dHampl, "0")
        'data.sAVGData.dHF_TAA_mV = m_CINIReader.GetValueString(AVG2, dHF_TAA_mV, "0")
        'data.sAVGData.dHFreq = m_CINIReader.GetValueString(AVG2, dHFreq, "0")
        'data.sAVGData.dIBias_mA = m_CINIReader.GetValueString(AVG2, dIBias_mA, "0")
        'data.sAVGData.dLampl = m_CINIReader.GetValueString(AVG2, dLampl, "0")
        'data.sAVGData.dLF_TAA_mV = m_CINIReader.GetValueString(AVG2, dLF_TAA_mV, "0")
        'data.sAVGData.dLFreq = m_CINIReader.GetValueString(AVG2, dLFreq, "0")
        'data.sAVGData.dMEW_uIn = m_CINIReader.GetValueString(AVG2, dMEW_uIn, "0")
        'data.sAVGData.dmewWRO_uIn = m_CINIReader.GetValueString(AVG2, dmewWRO_uIn, "0")
        'data.sAVGData.dMRR_Ohm = m_CINIReader.GetValueString(AVG2, dMRR_Ohm, "0")
        'data.sAVGData.dMRR2_Ohm = m_CINIReader.GetValueString(AVG2, dMRR2_Ohm, "0")
        'data.sAVGData.dMRW_uIn = m_CINIReader.GetValueString(AVG2, dMRW_uIn, "0")
        'data.sAVGData.dMWW_uIn = m_CINIReader.GetValueString(AVG2, dMWW_uIn, "0")
        'data.sAVGData.dOW2_dB = m_CINIReader.GetValueString(AVG2, dOW2_dB, "0")
        'data.sAVGData.dResolution = m_CINIReader.GetValueString(AVG2, dResolution, "0")
        'data.sAVGData.dSampl = m_CINIReader.GetValueString(AVG2, dSampl, "0")
        'data.sAVGData.dSFreq = m_CINIReader.GetValueString(AVG2, dSFreq, "0")
        'data.sAVGData.dTdP = m_CINIReader.GetValueString(AVG2, dTdP, "0")
        'data.sAVGData.dTdType = m_CINIReader.GetValueString(AVG2, dTdType, "0")
        'data.sAVGData.dTdV = m_CINIReader.GetValueString(AVG2, dTdV, "0")
        'data.sAVGData.dWRO_uIn = m_CINIReader.GetValueString(AVG2, dWRO_uIn, "0")
        'data.sAVGData.dWROffset_uIn = m_CINIReader.GetValueString(AVG2, dWROffset_uIn, "0")
        'data.sAVGData.ElapsedTime_sec = m_CINIReader.GetValueString(AVG2, ElapsedTime_sec, "0")
        'data.sAVGData.gfTAA_mV = m_CINIReader.GetValueString(AVG2, gfTAA_mV, "0")
        'data.sAVGData.HtrOhm = m_CINIReader.GetValueString(AVG2, HtrOhm, "0")
        'data.sAVGData.L_Slope = m_CINIReader.GetValueString(AVG2, L_Slope, "0")
        'data.sAVGData.LeftToCenter = m_CINIReader.GetValueString(AVG2, LeftToCenter, "0")
        'data.sAVGData.MEWRetest = m_CINIReader.GetValueString(AVG2, MEWRetest, "0")
        'data.sAVGData.PTD_ReloadFlag = m_CINIReader.GetValueString(AVG2, PTD_ReloadFlag, "0")
        'data.sAVGData.PTD_vDfhRd = m_CINIReader.GetValueString(AVG2, PTD_vDfhRd, "0")
        'data.sAVGData.PTD_vDfhWr = m_CINIReader.GetValueString(AVG2, PTD_vDfhWr, "0")
        'data.sAVGData.R_Slope = m_CINIReader.GetValueString(AVG2, R_Slope, "0")
        'data.sAVGData.Retest = m_CINIReader.GetValueString(AVG2, Retest, "0")
        'data.sAVGData.RightToCenter = m_CINIReader.GetValueString(AVG2, RightToCenter, "0")
        'data.sAVGData.RSO = m_CINIReader.GetValueString(AVG2, RSO, "0")
        'data.sAVGData.ServoBW2 = m_CINIReader.GetValueString(AVG2, ServoBW2, "0")
        'data.sAVGData.WSO = m_CINIReader.GetValueString(AVG2, WSO, "0")

        ''read normalize data
        'data.sNormData.ATTn = m_CINIReader.GetValueString(Norm2, ATTn, "0")
        'data.sNormData.CenterCor = m_CINIReader.GetValueString(Norm2, CenterCor, "0")
        'data.sNormData.dDeltaMRR_Ohm = m_CINIReader.GetValueString(Norm2, dDeltaMRR_Ohm, "0")
        'data.sNormData.dDeltaWidth_uIn = m_CINIReader.GetValueString(Norm2, dDeltaWidth_uIn, "0")
        'data.sNormData.dEMF_HFTAA_in_mV = m_CINIReader.GetValueString(Norm2, dEMF_HFTAA_in_mV, "0")
        'data.sNormData.dEMF_HFTAA_mV = m_CINIReader.GetValueString(Norm2, dEMF_HFTAA_mV, "0")
        'data.sNormData.dEMF_LFTAA_in_mV = m_CINIReader.GetValueString(Norm2, dEMF_LFTAA_in_mV, "0")
        'data.sNormData.dEMF_LFTAA_mV = m_CINIReader.GetValueString(Norm2, dEMF_LFTAA_mV, "0")
        'data.sNormData.dHampl = m_CINIReader.GetValueString(Norm2, dHampl, "0")
        'data.sNormData.dHF_TAA_mV = m_CINIReader.GetValueString(Norm2, dHF_TAA_mV, "0")
        'data.sNormData.dHFreq = m_CINIReader.GetValueString(Norm2, dHFreq, "0")
        'data.sNormData.dIBias_mA = m_CINIReader.GetValueString(Norm2, dIBias_mA, "0")
        'data.sNormData.dLampl = m_CINIReader.GetValueString(Norm2, dLampl, "0")
        'data.sNormData.dLF_TAA_mV = m_CINIReader.GetValueString(Norm2, dLF_TAA_mV)
        'data.sNormData.dLFreq = m_CINIReader.GetValueString(Norm2, dLFreq, "0")
        'data.sNormData.dMEW_uIn = m_CINIReader.GetValueString(Norm2, dMEW_uIn, "0")
        'data.sNormData.dmewWRO_uIn = m_CINIReader.GetValueString(Norm2, dmewWRO_uIn, "0")
        'data.sNormData.dMRR_Ohm = m_CINIReader.GetValueString(Norm2, dMRR_Ohm, "0")
        'data.sNormData.dMRR2_Ohm = m_CINIReader.GetValueString(Norm2, dMRR2_Ohm, "0")
        'data.sNormData.dMRW_uIn = m_CINIReader.GetValueString(Norm2, dMRW_uIn, "0")
        'data.sNormData.dMWW_uIn = m_CINIReader.GetValueString(Norm2, dMWW_uIn, "0")
        'data.sNormData.dOW2_dB = m_CINIReader.GetValueString(Norm2, dOW2_dB, "0")
        'data.sNormData.dResolution = m_CINIReader.GetValueString(Norm2, dResolution, "0")
        'data.sNormData.dSampl = m_CINIReader.GetValueString(Norm2, dSampl, "0")
        'data.sNormData.dSFreq = m_CINIReader.GetValueString(Norm2, dSFreq, "0")
        'data.sNormData.dTdP = m_CINIReader.GetValueString(Norm2, dTdP, "0")
        'data.sNormData.dTdType = m_CINIReader.GetValueString(Norm2, dTdType, "0")
        'data.sNormData.dTdV = m_CINIReader.GetValueString(Norm2, dTdV, "0")
        'data.sNormData.dWRO_uIn = m_CINIReader.GetValueString(Norm2, dWRO_uIn, "0")
        'data.sNormData.dWROffset_uIn = m_CINIReader.GetValueString(Norm2, dWROffset_uIn, "0")
        'data.sNormData.ElapsedTime_sec = m_CINIReader.GetValueString(Norm2, ElapsedTime_sec, "0")
        'data.sNormData.gfTAA_mV = m_CINIReader.GetValueString(Norm2, gfTAA_mV, "0")
        'data.sNormData.HtrOhm = m_CINIReader.GetValueString(Norm2, HtrOhm, "0")
        'data.sNormData.L_Slope = m_CINIReader.GetValueString(Norm2, L_Slope, "0")
        'data.sNormData.LeftToCenter = m_CINIReader.GetValueString(Norm2, LeftToCenter, "0")
        'data.sNormData.MEWRetest = m_CINIReader.GetValueString(Norm2, MEWRetest, "0")
        'data.sNormData.PTD_ReloadFlag = m_CINIReader.GetValueString(Norm2, PTD_ReloadFlag, "0")
        'data.sNormData.PTD_vDfhRd = m_CINIReader.GetValueString(Norm2, PTD_vDfhRd, "0")
        'data.sNormData.PTD_vDfhWr = m_CINIReader.GetValueString(Norm2, PTD_vDfhWr, "0")
        'data.sNormData.R_Slope = m_CINIReader.GetValueString(Norm2, R_Slope, "0")
        'data.sNormData.Retest = m_CINIReader.GetValueString(Norm2, Retest, "0")
        'data.sNormData.RightToCenter = m_CINIReader.GetValueString(Norm2, RightToCenter, "0")
        'data.sNormData.RSO = m_CINIReader.GetValueString(Norm2, RSO, "0")
        'data.sNormData.ServoBW2 = m_CINIReader.GetValueString(Norm2, ServoBW2, "0")
        'data.sNormData.WSO = m_CINIReader.GetValueString(Norm2, WSO, "0")

        'data.bIsRead = True
        'data.sAbort = GetDataAfterAbort()
        GetDCTData = Data

    End Function

    Private Function GetDataAfterAbort() As SShowAfterAbort

        Dim sData As SShowAfterAbort = Nothing

        sData.strGoodAF_Check = m_CINIReader.GetValueString("Header", "GoodAfterAbort", "0")
        sData.strGoodAF_rHtrOhm = m_CINIReader.GetValueString("Header", "rHtrOhm", "NULL")
        sData.strGoodAF_rLF = m_CINIReader.GetValueString("Header", "rgLFTAA", "NULL")
        sData.strGoodAF_rTdType = m_CINIReader.GetValueString("Header", "rTdType", "NULL")
        sData.strGoodAF_rMWW = m_CINIReader.GetValueString("Header", "rMWW", "NULL")
        sData.strGoodAF_rTuMRR = m_CINIReader.GetValueString("Header", "rTuMR_MRR", "NULL")
        sData.strGoodAF_rvPtd = m_CINIReader.GetValueString("Header", "rTdV", "NULL")
        sData.strGoodAF_rMRRCheck = m_CINIReader.GetValueString("Header", "rMRRCheck", "NULL")
        sData.strGoodAF_rTdFreqH = m_CINIReader.GetValueString("Header", "rTdFreqH", "NULL")
        sData.strGoodAF_rTdAmpH = m_CINIReader.GetValueString("Header", "rTdAmpH", "NULL")
        sData.strGoodAF_rTrkPAmp = m_CINIReader.GetValueString("Header", "rTrkPAmp", "NULL")
        sData.strGoodAF_rPESAbort = m_CINIReader.GetValueString("Header", "rPESAbort", "NULL")
        sData.strGoodAF_rTC = m_CINIReader.GetValueString("Header", "rTC", "NULL")
        sData.strGoodAF_rWriterImpCheck = m_CINIReader.GetValueString("Header", "rWriterImpCheck", "NULL")
        sData.strGoodAF_rFaultCheck = m_CINIReader.GetValueString("Header", "rFaultCheck", "NULL")
        sData.strGoodAF_MEW_Abort = m_CINIReader.GetValueString("Header", "MEW_Abort", "NULL")
        sData.strGrRev = m_CINIReader.GetValueString("Header", "GradeInfo", "NULL")

        GetDataAfterAbort = sData
    End Function
End Class

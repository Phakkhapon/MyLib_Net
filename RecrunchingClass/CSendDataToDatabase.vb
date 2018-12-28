
Imports MySql.Data.MySqlClient

Public Class CSendDataToDatabase
    Private m_bISReplace As Boolean = True
    Private m_myConn As MySqlConnection
    Private m_clsMySQL As New CMySQL
    Private m_strAliasProduct As String
    Private m_clsDataDTZ As CGetMemSetting
    Private m_dtbMCDefect As DataTable
    Private m_dtbHeader As DataTable

    Public Sub New(ByVal clsDataDTZ As CGetMemSetting, ByVal dtbMCDefect As DataTable, ByVal dtbHeader As DataTable, ByVal mySqlRawDataConn As MySqlConnection)
        m_clsDataDTZ = clsDataDTZ
        m_myConn = mySqlRawDataConn
        m_strAliasProduct = m_clsDataDTZ.GetValueString("Header", "Alias")
        m_dtbMCDefect = dtbMCDefect
        m_dtbHeader = dtbHeader
    End Sub

    Public Function SendToDatabase() As Boolean
        SendToDatabase = False
        'InsertToDatabase()
        InsertToDatabaseV2(m_dtbHeader)
        SendToDatabase = True
    End Function

    Public Function CreateDatabaseAndTable(ByVal strProduct As String) As Boolean
        CreateDatabaseAndTable = False

        If m_strAliasProduct = "" Then
            CreateDatabaseAndTable = False
        Else
            Dim clsNewDB As New CDatabaseManage(m_myConn)
            clsNewDB.AddNewProduct(strProduct)
            CreateDatabaseAndTable = True
        End If
    End Function

    'Private Sub InsertToDatabase()
    '    Dim strSQL As String
    '    Dim strROIOffsetLeft As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetLeft", "")
    '    Dim strROIOffsetTop As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetTop", "")
    '    Dim strROIOffsetRight As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetRight", "")
    '    Dim strROIOffsetBottom As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetBottom", "")
    '    Dim strROIRotationAngle As String = m_clsDataDTZ.GetValueString("Header", "ROIRotationAngle", "")

    '    Dim strPartID As String = m_clsDataDTZ.GetValueString("Header", "PartID")
    '    strPartID = strPartID & "//////" 'To avoid error from split function
    '    'strPartID = Replace(strPartID, "\", "/")
    '    Dim strSliderSite As String = Split(strPartID, "/")(5)
    '    Dim strAssy As String = Split(strPartID, "/")(2)
    '    Dim strGradeRev As String = m_clsDataDTZ.GetValueString("Header", "Gradeinfo")
    '    Dim strLot As String = Split(strPartID, "/")(1)
    '    strLot = Replace(strLot, "&", "M")
    '    Dim strMedia As String = m_clsDataDTZ.GetValueString("Header", "DISK Pack S/N")
    '    Dim strMediaTemp() As String = Split(strMedia, "/")
    '    strMedia = Split(strMediaTemp(0), "_")(0)
    '    Dim strMediaCount As String = ""
    '    If Split(strMediaTemp(0), "_").Length = 3 Then
    '        strMediaCount = Split(strMediaTemp(0), "_")(2)
    '    End If
    '    Dim strShoeSN As String = m_clsDataDTZ.GetValueString("Header", "CartID")
    '    Dim strSpec As String = Split(strPartID, "/")(0)
    '    Dim strTrack As String = strMediaTemp(0)
    '    Dim strWorkID As String = Split(strPartID, "/")(3)
    '    Dim strWTrayVersion As String = m_clsDataDTZ.GetValueString("Header", "WTrayVersion")
    '    Dim strDCT400Version As String = m_clsDataDTZ.GetValueString("Header", "DCT400Version")
    '    Dim strStartTime As String = m_clsDataDTZ.GetValueString("Header", "StartTime")
    '    Dim strTester As String = m_clsDataDTZ.GetValueString("Header", "Station")
    '    Dim strTestMode As String = m_clsDataDTZ.GetValueString("Header", "TestMode", "")
    '    Dim strGradeName As String = m_clsDataDTZ.GetValueString("Header", "GradeName")
    '    Dim strStepCode As String = m_clsDataDTZ.GetValueString("Header", "StepCode")
    '    Dim strTrayID As String = m_clsDataDTZ.GetValueString("Header", "TrayID", "")
    '    Dim strHGAPos As String = m_clsDataDTZ.GetValueString("Header", "HGAPosition", "")
    '    Dim strSliderPos As String = m_clsDataDTZ.GetValueString("Header", "SliderPosition", "")
    '    Dim strCGALot As String = m_clsDataDTZ.GetValueString("Header", "CGALot", "")
    '    Dim strCGANo As String = m_clsDataDTZ.GetValueString("Header", "CGANo", "")
    '    Dim strCGACount As String = m_clsDataDTZ.GetValueString("Header", "CGACount", "")
    '    Dim strCycleTime As String = m_clsDataDTZ.GetValueString("Header", "CycleTime", "")
    '    Dim strOCR_SN As String = m_clsDataDTZ.GetValueString("Header", "OCR_SN", "")
    '    Dim strROIOffsetPosition As String = ""
    '    If strROIOffsetLeft <> "" OrElse strROIOffsetTop <> "" OrElse strROIOffsetRight <> "" OrElse strROIOffsetBottom <> "" OrElse strROIRotationAngle <> "" Then
    '        strROIOffsetPosition = strROIOffsetLeft & "/" & strROIOffsetTop & "/" & strROIOffsetRight & "/" & strROIOffsetBottom & "/" & strROIRotationAngle
    '    End If
    '    Dim strChom As String = m_clsDataDTZ.GetValueString("Header", "CHOM(UV)")
    '    Dim strTotalTime As String = m_clsDataDTZ.GetValueString("Header", "Total_ElapsedTime(ms)", "")
    '    Dim strElapsedTime As String = m_clsDataDTZ.GetValueString("Header", "ElapsedTime(sec)")
    '    Dim strSwapTime As String = m_clsDataDTZ.GetValueString("Header", "SwapTime", "")
    '    Dim strShoe As String = Replace(m_clsDataDTZ.GetValueString("Header", "ShoeNo", "S1"), "S", "")
    '    Dim strDoverHMI As String = m_clsDataDTZ.GetValueString("Header", "DoverHMI", "")
    '    Dim strDoverScript As String = m_clsDataDTZ.GetValueString("Header", "DoverScript", "")
    '    Dim strTipNo As String = m_clsDataDTZ.GetValueString("Header", "TipNo", "1")
    '    Dim strBarNo As String = m_clsDataDTZ.GetValueString("Header", "BarNo", "")
    '    If strShoe = "" Or Len(strShoe) <> 1 Then Exit Sub

    '    If m_bInsertOption = True Then
    '        strSQL = "INSERT"
    '    Else
    '        strSQL = "REPLACE"
    '    End If
    '    strSQL = strSQL & " INTO db_" & m_strAliasProduct & ".tabdetail_header "
    '    strSQL = strSQL & "(tag_id,"
    '    strSQL = strSQL & "test_time_bigint,"
    '    strSQL = strSQL & "test_time," '` DATETIME NOT NULL DEFAULT '0000-00-00 00:00:00',
    '    strSQL = strSQL & "tester," ' INT(11) NOT NULL DEFAULT '0',
    '    strSQL = strSQL & "spec," ' INT(11) DEFAULT NULL,
    '    strSQL = strSQL & "lot," '` INT(11) DEFAULT NULL,
    '    strSQL = strSQL & "TestMode,"
    '    strSQL = strSQL & "assy," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "mediaSN," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "MediaCount,"
    '    strSQL = strSQL & "grade_rev," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "shoe," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "SliderSite,"
    '    strSQL = strSQL & "TipNo,"
    '    strSQL = strSQL & "BarNo,"
    '    strSQL = strSQL & "gradeName," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "StepCode," ' ` TINYINT(1) DEFAULT '0',"
    '    strSQL = strSQL & "shoeSN," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "hga_SN," ' ` VARCHAR(11) DEFAULT NULL,"
    '    strSQL = strSQL & "OCR_SN,"
    '    strSQL = strSQL & "ROIOffsetPosition,"
    '    strSQL = strSQL & "Chom,"
    '    strSQL = strSQL & "SliderAngle,"
    '    strSQL = strSQL & "SldAngleCGA,"
    '    strSQL = strSQL & "SldWidth,"
    '    strSQL = strSQL & "SldLength,"
    '    strSQL = strSQL & "ABSMatchScroll,"
    '    strSQL = strSQL & "RangeSTD,"
    '    strSQL = strSQL & "oprID," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "WorkID," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "TestTime," ' ` DOUBLE DEFAULT '0',"
    '    strSQL = strSQL & "TotalTime,"
    '    strSQL = strSQL & "DoverHMI,"
    '    strSQL = strSQL & "DoverScript,"
    '    strSQL = strSQL & "RobotSwapTime,"
    '    strSQL = strSQL & "InTrayID,"
    '    strSQL = strSQL & "InSliderPosition,"
    '    strSQL = strSQL & "TrayID," '` VARCHAR(16) DEFAULT NULL,"
    '    strSQL = strSQL & "TrayIDOut,"
    '    strSQL = strSQL & "OutSliderPosition,"
    '    strSQL = strSQL & "HGAPos," '` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "SliderPos," '` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "CGALot,"
    '    strSQL = strSQL & "CGANo," '` VARCHAR(16) DEFAULT NULL,"
    '    strSQL = strSQL & "CGACount," '` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "CycleTime," '` DOUBLE DEFAULT NULL,"
    '    strSQL = strSQL & "TrackID," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "WTrayVersion," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "DCT400Version) " ' ` INT(11) DEFAULT NULL,"

    '    Dim strTestTime_int As String = strStartTime
    '    strTestTime_int = Replace(strTestTime_int, " ", "")
    '    strTestTime_int = Replace(strTestTime_int, "-", "")
    '    strTestTime_int = Replace(strTestTime_int, ":", "")
    '    strTestTime_int = Replace(strTestTime_int, "/", "")

    '    strSQL = strSQL & "SELECT "
    '    strSQL = strSQL & "'" & strTestTime_int & strTester & "',"
    '    strSQL = strSQL & "'" & strTestTime_int & "',"
    '    strSQL = strSQL & "'" & strStartTime & "',"
    '    strSQL = strSQL & "'" & strTester & "',"
    '    strSQL = strSQL & "'" & strSpec & "',"
    '    strSQL = strSQL & "'" & strLot & "',"
    '    strSQL = strSQL & "'" & strTestMode & "',"
    '    strSQL = strSQL & "'" & strAssy & "',"
    '    strSQL = strSQL & "'" & strMedia & "',"
    '    strSQL = strSQL & "'" & strMediaCount & "',"
    '    strSQL = strSQL & "'" & strGradeRev & "',"
    '    strSQL = strSQL & "'" & strShoe & "',"
    '    strSQL = strSQL & "'" & strSliderSite & "',"
    '    strSQL = strSQL & "'" & strTipNo & "',"
    '    strSQL = strSQL & "'" & strBarNo & "',"
    '    strSQL = strSQL & "'" & strGradeName & "',"
    '    strSQL = strSQL & "'" & strStepCode & "',"
    '    strSQL = strSQL & "'" & strShoeSN & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "HeadSN") & "',"
    '    strSQL = strSQL & "'" & strOCR_SN & "',"
    '    strSQL = strSQL & "'" & strROIOffsetPosition & "',"
    '    strSQL = strSQL & "'" & strChom & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "SldAngle") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "SldAngleCGA") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "SldWidth") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "SldLength") & "',"
    '    Dim strMathScroll As String = m_clsDataDTZ.GetValueString("Header", "ABSMatchScroll")
    '    If strMathScroll = "" Then strMathScroll = m_clsDataDTZ.GetValueString("Header", "MatchScore")
    '    strSQL = strSQL & "'" & strMathScroll & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "RangeSTD") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "Operator") & "',"
    '    strSQL = strSQL & "'" & strWorkID & "',"
    '    strSQL = strSQL & "'" & strElapsedTime & "',"
    '    strSQL = strSQL & "'" & strTotalTime & "',"
    '    strSQL = strSQL & "'" & strDoverHMI & "',"
    '    strSQL = strSQL & "'" & strDoverScript & "',"
    '    strSQL = strSQL & "'" & strSwapTime & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "InTrayID") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "InSliderPosition") & "',"
    '    strSQL = strSQL & "'" & strTrayID & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "TrayIDOut") & "',"
    '    strSQL = strSQL & "'" & m_clsDataDTZ.GetValueString("Header", "OutSliderPosition") & "',"
    '    strSQL = strSQL & "'" & strHGAPos & "',"
    '    strSQL = strSQL & "'" & strSliderPos & "',"
    '    strSQL = strSQL & "'" & strCGALot & "',"
    '    strSQL = strSQL & "'" & strCGANo & "',"
    '    strSQL = strSQL & "'" & strCGACount & "',"
    '    strSQL = strSQL & "'" & strCycleTime & "',"
    '    strSQL = strSQL & "'" & strTrack & "'," ' ` INT(11) DEFAULT NULL,"
    '    strSQL = strSQL & "'" & strWTrayVersion & "',"
    '    strSQL = strSQL & "'" & strDCT400Version & "';"

    '    Dim strHeadCF_Add As String
    '    Dim strHeadCF_Mul As String
    '    Dim strHeadValue As String
    '    Dim strHead_DeltaGOS As String

    '    Dim strInsertOption As String

    '    If m_bInsertOption = True Then
    '        strInsertOption = "INSERT"
    '    Else
    '        strInsertOption = "REPLACE"
    '    End If

    '    strHeadCF_Add = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_cfadd(tag_id,test_time_bigint,test_time,tester,"
    '    strHeadCF_Mul = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_cfmul(tag_id,test_time_bigint,test_time,tester,"
    '    strHeadValue = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_value(tag_id,test_time_bigint,test_time,tester,"
    '    strHead_DeltaGOS = "REPLACE INTO db_" & m_strAliasProduct & ".tabfactor_deltagos(tag_id,test_time_bigint,test_time,tester,"


    '    Dim strValue As String = ""
    '    Dim strCF_Add As String = ""
    '    Dim strCF_Mul As String = ""
    '    Dim strDelta_GOS As String = ""

    '    Dim strDTZValue() As String = m_clsDataDTZ.GetValueSection("DTZValue")
    '    Dim strCFAdd() As String = m_clsDataDTZ.GetValueSection("CFAdd")
    '    Dim strCFMul() As String = m_clsDataDTZ.GetValueSection("CFMul")
    '    Dim strCFMedia() As String = m_clsDataDTZ.GetValueSection("CF_MEDIA")
    '    Dim strDeltaGOS() As String = m_clsDataDTZ.GetValueSection("Delta_GOS")

    '    Dim strGOSLot As String = ""
    '    Dim strGOSSpec As String = ""

    '    For nDeltaGOS As Integer = 1 To strDeltaGOS.Length - 1
    '        Dim strParaTemp() As String = strDeltaGOS(nDeltaGOS).Split("=")
    '        Dim strParaName As String = strParaTemp(0)
    '        If strParaName.ToUpper <> "GOSTIME" Then
    '            Dim strDelta As String = ""
    '            If strParaName.ToUpper = "LOT" Or strParaName.ToUpper = "SPEC" Then
    '                strDelta = strParaTemp(1)
    '            Else
    '                strDelta = Split(strParaTemp(1), ",")(2)
    '            End If
    '            strHead_DeltaGOS = strHead_DeltaGOS & strParaName & ","
    '            strDelta_GOS = strDelta_GOS & "'" & strDelta & "',"
    '        End If
    '    Next nDeltaGOS

    '    For nAdd As Integer = 1 To strCFAdd.Length - 1
    '        Dim strParaTemp() As String = strCFAdd(nAdd).Split("=")
    '        Dim strParaName As String = strParaTemp(0)
    '        Dim strValueAdd As String = strParaTemp(1)
    '        strHeadCF_Add = strHeadCF_Add & strParaName & ","
    '        strCF_Add = strCF_Add & "'" & strValueAdd & "',"
    '    Next nAdd

    '    For nMul As Integer = 1 To strCFMul.Length - 1
    '        Dim strParaTemp() As String = strCFMul(nMul).Split("=")
    '        Dim strParaName As String = strParaTemp(0)
    '        Dim strValueMul As String = strParaTemp(1)
    '        strHeadCF_Mul = strHeadCF_Mul & strParaName & ","
    '        strCF_Mul = strCF_Mul & "'" & strValueMul & "',"
    '    Next nMul

    '    For nData As Integer = 1 To strDTZValue.Length - 1
    '        Dim strParaTemp() As String = strDTZValue(nData).Split("=")
    '        Dim strParaName As String = strParaTemp(0)
    '        Dim strValueDTZ As String = strParaTemp(1)
    '        strHeadValue = strHeadValue & strParaName & ","
    '        strValue = strValue & "'" & strValueDTZ & "',"
    '    Next nData

    '    strHeadCF_Add = Left(strHeadCF_Add, Len(strHeadCF_Add) - 1)
    '    strHeadCF_Mul = Left(strHeadCF_Mul, Len(strHeadCF_Mul) - 1)
    '    strHead_DeltaGOS = Left(strHead_DeltaGOS, Len(strHead_DeltaGOS) - 1)
    '    If strCF_Add <> "" Then strCF_Add = Left(strCF_Add, Len(strCF_Add) - 1)
    '    If strCF_Mul <> "" Then strCF_Mul = Left(strCF_Mul, Len(strCF_Mul) - 1)
    '    If strDelta_GOS <> "" Then strDelta_GOS = Left(strDelta_GOS, Len(strDelta_GOS) - 1)

    '    If Right(strHeadValue, 1) = "," Then
    '        strHeadValue = Left(strHeadValue, Len(strHeadValue) - 1)
    '        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
    '    End If

    '    Dim strGOSTime As String = m_clsDataDTZ.GetValueString("Delta_GOS", "GOSTime", "")
    '    Dim strGOSTime_int As String = strGOSTime
    '    strGOSTime_int = Replace(strGOSTime_int, " ", "")
    '    strGOSTime_int = Replace(strGOSTime_int, "-", "")
    '    strGOSTime_int = Replace(strGOSTime_int, ":", "")
    '    strGOSTime_int = Replace(strGOSTime_int, "/", "")

    '    If strCF_Add <> "" Then strCF_Add = strHeadCF_Add & ") SELECT '" & strTestTime_int & strTester & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strCF_Add & ";"
    '    If strCF_Mul <> "" Then strCF_Mul = strHeadCF_Mul & ") SELECT '" & strTestTime_int & strTester & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strCF_Mul & ";"
    '    If strDelta_GOS <> "" Then strDelta_GOS = strHead_DeltaGOS & ") SELECT '" & strGOSTime_int & strTester & "','" & strGOSTime_int & "','" & strGOSTime & "','" & strTester & "'," & strDelta_GOS & ";"
    '    If strValue <> "" Then strValue = strHeadValue & ") SELECT '" & strTestTime_int & strTester & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strValue & ";"

    '    strSQL = strSQL & strValue & strCF_Add & strCF_Mul & strDelta_GOS

    '    Dim strDateTmp() As String = strStartTime.Split(":")
    '    Dim strDateByHour As String = ""
    '    If strDateTmp.Length > 1 Then
    '        If CInt(strDateTmp(1)) < 30 Then
    '            strDateByHour = strDateTmp(0) & ":00:00"
    '        Else
    '            strDateByHour = strDateTmp(0) & ":30:00"
    '        End If
    '    End If
    '    Dim strTesterSQL As String = "REPLACE INTO db_" & m_strAliasProduct & ".tabtester "
    '    strTesterSQL = strTesterSQL & "SELECT '" & strDateByHour & "','" & strTester & "','" & strLot & "','" & strSpec & "';"

    '    Dim bIsGradePass As Boolean = False
    '    If InStr(strGradeName, "PASS", CompareMethod.Text) Then
    '        bIsGradePass = True
    '    End If

    '    Dim strSQLDefect As String = GetSqlUpdateSumHgaDefect(m_clsDataDTZ, strStartTime, bIsGradePass, strGradeName, m_strAliasProduct, strTester, strLot, strSpec, strShoe, m_dtbMCDefect)

    '    strSQL = strSQL & strTesterSQL & strSQLDefect

    '    m_clsMySQL.CommandNonQuery(strSQL, m_myConn)        'Insert all data

    'End Sub

    Private Sub InsertToDatabaseV2(ByVal dtbHeader As DataTable)

        Dim strSQL As String = ""
        Dim strROIOffsetLeft As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetLeft", "")
        Dim strROIOffsetTop As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetTop", "")
        Dim strROIOffsetRight As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetRight", "")
        Dim strROIOffsetBottom As String = m_clsDataDTZ.GetValueString("Header", "ROIOffsetBottom", "")
        Dim strROIRotationAngle As String = m_clsDataDTZ.GetValueString("Header", "ROIRotationAngle", "")
        Dim strROIOffsetPosition As String = ""
        If strROIOffsetLeft <> "" OrElse strROIOffsetTop <> "" OrElse strROIOffsetRight <> "" OrElse strROIOffsetBottom <> "" OrElse strROIRotationAngle <> "" Then
            strROIOffsetPosition = strROIOffsetLeft & "/" & strROIOffsetTop & "/" & strROIOffsetRight & "/" & strROIOffsetBottom & "/" & strROIRotationAngle
            m_clsDataDTZ.WriteValueString("Header", "ROIOffsetPosition", strROIOffsetPosition)
        End If

        Dim strPartID As String = m_clsDataDTZ.GetValueString("Header", "PartID")
        strPartID = strPartID & "//////" 'To avoid error from split function
        Dim strSliderSite As String = Split(strPartID, "/")(5)
        If strSliderSite <> "" Then m_clsDataDTZ.WriteValueString("Header", "SliderSite", strSliderSite)

        Dim strAssy As String = Split(strPartID, "/")(2)
        If strAssy <> "" Then m_clsDataDTZ.WriteValueString("Header", "Assy", strAssy)

        'Dim strGradeRev As String = m_clsDataDTZ.GetValueString("Header", "Gradeinfo")
        Dim strLot As String = Split(strPartID, "/")(1)
        'strLot = Replace(strLot, "&", "M")
        If strLot <> "" Then m_clsDataDTZ.WriteValueString("Header", "Lot", strLot)

        Dim strPartNumber As String = Split(strPartID, "/")(4)
        If strPartNumber <> "" Then m_clsDataDTZ.WriteValueString("Header", "PartNumber", strPartNumber)

        Dim strMedia As String = m_clsDataDTZ.GetValueString("Header", "DISK Pack S/N")
        Dim strMediaTemp() As String = Split(strMedia, "/")
        strMedia = Split(strMediaTemp(0), "_")(0)
        m_clsDataDTZ.WriteValueString("Header", "MediaSN", strMedia)

        Dim strMediaSurface As String = ""
        If Split(strMediaTemp(0), "_").Length = 3 Then
            strMediaSurface = Split(strMediaTemp(0), "_")(1)
        End If
        m_clsDataDTZ.WriteValueString("Header", "MediaSurface", strMediaSurface)

        Dim strMediaCount As String = ""
        If Split(strMediaTemp(0), "_").Length = 3 Then
            strMediaCount = Split(strMediaTemp(0), "_")(2)
        End If
        m_clsDataDTZ.WriteValueString("Header", "MediaCount", strMediaCount)

        Dim strSpec As String = Split(strPartID, "/")(0)
        m_clsDataDTZ.WriteValueString("Header", "Spec", strSpec)
        Dim strTrack As String = strMediaTemp(0)
        m_clsDataDTZ.WriteValueString("Header", "TrackID", strTrack)
        Dim strWorkID As String = Split(strPartID, "/")(3)
        m_clsDataDTZ.WriteValueString("Header", "WorkID", strWorkID)
        Dim strShoe As String = Replace(m_clsDataDTZ.GetValueString("Header", "ShoeNo", "S1"), "S", "")
        m_clsDataDTZ.WriteValueString("Header", "ShoeNum", strShoe)

        Dim strMathScroll As String = m_clsDataDTZ.GetValueString("Header", "ABSMatchScroll")
        If strMathScroll = "" Then strMathScroll = m_clsDataDTZ.GetValueString("Header", "MatchScore")
        m_clsDataDTZ.WriteValueString("Header", "ABSMatchScroll", strMathScroll)

        Dim strStartTime As String = m_clsDataDTZ.GetValueString("Header", "StartTime")
        Dim strTester As String = m_clsDataDTZ.GetValueString("Header", "Station")
        Dim strGradeName As String = m_clsDataDTZ.GetValueString("Header", "GradeName")
        Dim strHeadSN As String = m_clsDataDTZ.GetValueString("Header", "HeadSN")
        Dim strTestTime_int As String = strStartTime
        strTestTime_int = Replace(strTestTime_int, " ", "")
        strTestTime_int = Replace(strTestTime_int, "-", "")
        strTestTime_int = Replace(strTestTime_int, ":", "")
        strTestTime_int = Replace(strTestTime_int, "/", "")

        Dim strPrimaryKey As String = strTestTime_int & strTester & strHeadSN

        If strShoe = "" Or Len(strShoe) <> 1 Then Exit Sub

        Dim strHeaderSQL As String = ""
        Dim strValueSQL As String = ""
        If m_bISReplace = True Then
            strHeaderSQL = "REPLACE"
        Else
            strHeaderSQL = "INSERT"
        End If

        strHeaderSQL = strHeaderSQL & " INTO db_" & m_strAliasProduct & ".tabdetail_header "
        strHeaderSQL = strHeaderSQL & "(tag_id,test_time_bigint,"

        strValueSQL = strValueSQL & "SELECT "
        strValueSQL = strValueSQL & "'" & strPrimaryKey & "','" & strTestTime_int & "',"

        For nHeader As Integer = 0 To dtbHeader.Rows.Count - 1
            Dim strHeaderName As String = dtbHeader.Rows(nHeader).Item("HeaderName")
            Dim strSection As String = dtbHeader.Rows(nHeader).Item("Section")
            Dim strKey As String = dtbHeader.Rows(nHeader).Item("Key")
            Dim strHeaderValue As String = m_clsDataDTZ.GetValueString(strSection, strKey)
            strHeaderSQL = strHeaderSQL & strHeaderName & ","
            strValueSQL = strValueSQL & "'" & strHeaderValue & "',"
        Next nHeader
        If Right(strHeaderSQL, 1) = "," Then strHeaderSQL = Left(strHeaderSQL, strHeaderSQL.Length - 1) & ") "
        If Right(strValueSQL, 1) = "," Then strValueSQL = Left(strValueSQL, strValueSQL.Length - 1) & ";"
        strSQL = strHeaderSQL & strValueSQL

        Dim strHeadCF_Add As String
        Dim strHeadCF_Mul As String
        Dim strHeadValue As String
        Dim strHead_DeltaGOS As String

        Dim strInsertOption As String

        If m_bISReplace = True Then
            strInsertOption = "REPLACE"
        Else
            strInsertOption = "INSERT"
        End If

        strHeadCF_Add = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_cfadd(tag_id,test_time_bigint,test_time,tester,"
        strHeadCF_Mul = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_cfmul(tag_id,test_time_bigint,test_time,tester,"
        strHeadValue = strInsertOption & " INTO db_" & m_strAliasProduct & ".tabfactor_value(tag_id,test_time_bigint,test_time,tester,"
        strHead_DeltaGOS = "REPLACE INTO db_" & m_strAliasProduct & ".tabfactor_deltagos(tag_id,test_time_bigint,test_time,tester,"

        Dim strValue As String = ""
        Dim strCF_Add As String = ""
        Dim strCF_Mul As String = ""
        Dim strDelta_GOS As String = ""

        Dim strDTZValue() As String = m_clsDataDTZ.GetValueSection("DTZValue")
        Dim strCFAdd() As String = m_clsDataDTZ.GetValueSection("CFAdd")
        Dim strCFMul() As String = m_clsDataDTZ.GetValueSection("CFMul")
        Dim strCFMedia() As String = m_clsDataDTZ.GetValueSection("CF_MEDIA")
        Dim strDeltaGOS() As String = m_clsDataDTZ.GetValueSection("Delta_GOS")

        Dim strGOSLot As String = ""
        Dim strGOSSpec As String = ""

        For nDeltaGOS As Integer = 1 To strDeltaGOS.Length - 1
            Dim strParaTemp() As String = strDeltaGOS(nDeltaGOS).Split("=")
            Dim strParaName As String = strParaTemp(0)
            If strParaName.ToUpper <> "GOSTIME" Then
                Dim strDelta As String = ""
                If strParaName.ToUpper = "LOT" Or strParaName.ToUpper = "SPEC" Then
                    strDelta = strParaTemp(1)
                Else
                    strDelta = Split(strParaTemp(1), ",")(2)
                End If
                strHead_DeltaGOS = strHead_DeltaGOS & strParaName & ","
                strDelta_GOS = strDelta_GOS & "'" & strDelta & "',"
            End If
        Next nDeltaGOS

        For nAdd As Integer = 1 To strCFAdd.Length - 1
            Dim strParaTemp() As String = strCFAdd(nAdd).Split("=")
            Dim strParaName As String = strParaTemp(0)
            Dim strValueAdd As String = strParaTemp(1)
            strHeadCF_Add = strHeadCF_Add & strParaName & ","
            strCF_Add = strCF_Add & "'" & strValueAdd & "',"
        Next nAdd

        For nMul As Integer = 1 To strCFMul.Length - 1
            Dim strParaTemp() As String = strCFMul(nMul).Split("=")
            Dim strParaName As String = strParaTemp(0)
            Dim strValueMul As String = strParaTemp(1)
            strHeadCF_Mul = strHeadCF_Mul & strParaName & ","
            strCF_Mul = strCF_Mul & "'" & strValueMul & "',"
        Next nMul

        For nData As Integer = 1 To strDTZValue.Length - 1
            Dim strParaTemp() As String = strDTZValue(nData).Split("=")
            Dim strParaName As String = strParaTemp(0)
            Dim strValueDTZ As String = strParaTemp(1)
            strHeadValue = strHeadValue & strParaName & ","
            strValue = strValue & "'" & strValueDTZ & "',"
        Next nData

        strHeadCF_Add = Left(strHeadCF_Add, Len(strHeadCF_Add) - 1)
        strHeadCF_Mul = Left(strHeadCF_Mul, Len(strHeadCF_Mul) - 1)
        strHead_DeltaGOS = Left(strHead_DeltaGOS, Len(strHead_DeltaGOS) - 1)
        If strCF_Add <> "" Then strCF_Add = Left(strCF_Add, Len(strCF_Add) - 1)
        If strCF_Mul <> "" Then strCF_Mul = Left(strCF_Mul, Len(strCF_Mul) - 1)
        If strDelta_GOS <> "" Then strDelta_GOS = Left(strDelta_GOS, Len(strDelta_GOS) - 1)

        If Right(strHeadValue, 1) = "," Then
            strHeadValue = Left(strHeadValue, Len(strHeadValue) - 1)
            If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        End If

        Dim strGOSTime As String = m_clsDataDTZ.GetValueString("Delta_GOS", "GOSTime", "")
        Dim strGOSTime_int As String = strGOSTime
        strGOSTime_int = Replace(strGOSTime_int, " ", "")
        strGOSTime_int = Replace(strGOSTime_int, "-", "")
        strGOSTime_int = Replace(strGOSTime_int, ":", "")
        strGOSTime_int = Replace(strGOSTime_int, "/", "")

        If strCF_Add <> "" Then strCF_Add = strHeadCF_Add & ") SELECT '" & strPrimaryKey & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strCF_Add & ";"
        If strCF_Mul <> "" Then strCF_Mul = strHeadCF_Mul & ") SELECT '" & strPrimaryKey & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strCF_Mul & ";"
        If strDelta_GOS <> "" Then strDelta_GOS = strHead_DeltaGOS & ") SELECT '" & strGOSTime_int & strTester & "','" & strGOSTime_int & "','" & strGOSTime & "','" & strTester & "'," & strDelta_GOS & ";"
        If strValue <> "" Then strValue = strHeadValue & ") SELECT '" & strPrimaryKey & "','" & strTestTime_int & "','" & strStartTime & "','" & strTester & "'," & strValue & ";"

        strSQL = strSQL & strValue & strCF_Add & strCF_Mul & strDelta_GOS

        Dim strDateTmp() As String = strStartTime.Split(":")
        Dim strDateByHour As String = ""
        If strDateTmp.Length > 1 Then
            If CInt(strDateTmp(1)) < 30 Then
                strDateByHour = strDateTmp(0) & ":00:00"
            Else
                strDateByHour = strDateTmp(0) & ":30:00"
            End If
        End If
        Dim strTesterSQL As String = "REPLACE INTO db_" & m_strAliasProduct & ".tabtester "
        strTesterSQL = strTesterSQL & "SELECT '" & strDateByHour & "','" & strTester & "','" & strLot & "','" & strSpec & "';"

        Dim bIsGradePass As Boolean = False
        If InStr(strGradeName, "PASS", CompareMethod.Text) Then
            bIsGradePass = True
        End If

        Dim strSQLDefect As String = GetSqlUpdateSumHgaDefect(m_clsDataDTZ, strStartTime, bIsGradePass, strGradeName, m_strAliasProduct, strTester, strLot, strSpec, strShoe, m_dtbMCDefect)

        strSQL = strSQL & strTesterSQL & strSQLDefect

        m_clsMySQL.CommandNoQuery(strSQL, m_myConn)        'Insert all data

    End Sub

    Private Function GetSqlUpdateSumHgaDefect(ByVal clsFileDCT As CGetMemSetting, ByVal strTestTime As String, ByVal bGradePass As Boolean, ByVal strGradeName As String, ByVal strAliasProduct As String, ByVal strTester As String, ByVal strLot As String, ByVal strSpec As String, ByVal strShoe As String, ByVal dtbGradeMapping As DataTable) As String

        Dim strDefectHeader As String = ""
        Dim strDefectValue As String = " VALUES("
        Dim strDefectUpdate As String = "ON DUPLICATE KEY UPDATE "

        strDefectHeader = "INSERT INTO db_" & strAliasProduct & ".tabsummary_hgadefect("
        strDefectHeader = strDefectHeader & "update_time,"
        strDefectHeader = strDefectHeader & "tester,"
        strDefectHeader = strDefectHeader & "Lot,"
        strDefectHeader = strDefectHeader & "Spec,"
        strDefectHeader = strDefectHeader & "Shoe,"
        strDefectHeader = strDefectHeader & "TotalHGA,"
        strDefectHeader = strDefectHeader & "TotalPass,"

        strDefectValue = strDefectValue & "'" & strTestTime & "',"
        strDefectValue = strDefectValue & "'" & strTester & "',"
        strDefectValue = strDefectValue & "'" & strLot & "',"
        strDefectValue = strDefectValue & "'" & strSpec & "',"
        strDefectValue = strDefectValue & "'" & strShoe & "',"
        strDefectValue = strDefectValue & "'" & 1 & "',"
        If bGradePass = True Then
            strDefectValue = strDefectValue & "'1',"
        Else
            strDefectValue = strDefectValue & "'0',"
        End If

        strDefectUpdate = strDefectUpdate & "Update_time='" & strTestTime & "',"

        If InStr(strGradeName, "REJECT LOW_BNOHGA") = 0 And InStr(strGradeName, "REJECT LOW_BNOSLIDER") = 0 And strGradeName <> "" Then
            strDefectUpdate = strDefectUpdate & "TotalHGA=IFNULL(TotalHGA,0)+1,"
        Else
            strDefectUpdate = strDefectUpdate & "TotalHGA=IFNULL(TotalHGA,0),"
        End If

        If bGradePass = True Then
            strDefectUpdate = strDefectUpdate & "TotalPass=IFNULL(TotalPass,0)+1,"
        End If

        For nGrade As Integer = 0 To dtbGradeMapping.Rows.Count - 1
            Dim strHeader As String = dtbGradeMapping.Rows(nGrade).Item("HeaderName").ToString
            Dim strKey As String = dtbGradeMapping.Rows(nGrade).Item("SectionName").ToString
            Dim strResult As String = clsFileDCT.GetValueString(strHeader, strKey).ToUpper
            Dim nDefectID As Integer = dtbGradeMapping.Rows(nGrade).Item("MCCodeID")
            If strResult <> "" Then
                Dim strResultCount() As String = Split(dtbGradeMapping.Rows(nGrade).Item("ResultCount").ToString, ";")
                Dim bDetectGrade As Boolean = False
                Dim strGradeMapping As String = dtbGradeMapping.Rows(nGrade).Item("MCDefectName").ToString.ToUpper
                Select Case strGradeMapping
                    Case "TD_AUTO_RETEST"
                        Dim dblTdTyp As Double = CDbl(strResult) * 1000
                        Dim nMod As Integer = dblTdTyp Mod 1000
                        If nMod > 0 Then
                            bDetectGrade = True
                        End If
                    Case "ABORTNOTD"
                        For nCondition As Integer = 0 To strResultCount.Length - 1
                            If strResult = strResultCount(nCondition).ToUpper And strGradeName = "FAIL_NO_READING" Then
                                bDetectGrade = True
                            End If
                        Next nCondition
                    Case Else
                        For nCondition As Integer = 0 To strResultCount.Length - 1
                            If InStr(strGradeMapping, "PASS", CompareMethod.Text) Then
                                If InStr(strResult, strResultCount(nCondition).ToUpper, CompareMethod.Text) Then
                                    bDetectGrade = True
                                End If
                            Else
                                If strResult = strResultCount(nCondition).ToUpper Then
                                    bDetectGrade = True
                                End If
                            End If
                        Next nCondition
                End Select
                If bDetectGrade = True Then    'Detect Defect = True.
                    Dim strDefectID As String = "Defect" & nDefectID
                    strDefectHeader = strDefectHeader & strDefectID & ","
                    strDefectValue = strDefectValue & "1,"
                    strDefectUpdate = strDefectUpdate & strDefectID & "=IFNULL(" & strDefectID & ",0)+1,"
                End If
            End If
        Next nGrade
        If Right(strDefectHeader, 1) = "," Then strDefectHeader = Left(strDefectHeader, Len(strDefectHeader) - 1) & ") "
        If Right(strDefectValue, 1) = "," Then strDefectValue = Left(strDefectValue, Len(strDefectValue) - 1) & ") "
        If Right(strDefectUpdate, 1) = "," Then strDefectUpdate = Left(strDefectUpdate, Len(strDefectUpdate) - 1) & ";"
        GetSqlUpdateSumHgaDefect = strDefectHeader & strDefectValue & strDefectUpdate
    End Function

End Class

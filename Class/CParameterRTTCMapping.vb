Imports MySql.Data.MySqlClient

Public Class CParameterRTTCMapping
    Public Enum enuOrderParameterBy
        eByParamID = 0
        eByOrderID
        eByName
    End Enum

    Public Enum enuMCGradeType
        eGradeAll = 0
        eGradeNone
        eGradeU
        eGradeV
    End Enum

    Private m_MysqlConn As MySqlConnection

    Public Sub New(ByVal MysqlConn As MySqlConnection)
        m_MysqlConn = MysqlConn
    End Sub

    Public Function GetTemplateParamTable() As DataTable
        Dim strSQL As String
        strSQL = "SHOW COLUMNS FROM db_master.tabfactor_value "
        Dim clsSQL As New CMySQL
        Dim dtbParam As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbParam.Rows.RemoveAt(0)   'remove tag_id
        dtbParam.Rows.RemoveAt(0)   'remove test_time
        dtbParam.Rows.RemoveAt(0)   'remove tester
        dtbParam.Rows.RemoveAt(0)   'remove tester
        Dim dtbTemplate As New DataTable
        For nParam As Integer = 0 To dtbParam.Rows.Count - 1
            Dim strParam As String = dtbParam.Rows(nParam).Item("field")
            dtbTemplate.Columns.Add(strParam & "_n", GetType(Double))
            dtbTemplate.Columns.Add(strParam & "_a", GetType(Double))
        Next nParam
        GetTemplateParamTable = dtbTemplate
    End Function

    Public Function GetMCDefectMapping(ByVal nGrade As enuMCGradeType) As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM db_parameter_mapping.tabmcdefect A "
        strSQL = strSQL & "WHERE A.IsEnable =1 "
        Select Case nGrade
            Case enuMCGradeType.eGradeNone
                strSQL = strSQL & "AND ValueType='GradeNone' "
            Case enuMCGradeType.eGradeU
                strSQL = strSQL & "AND ValueType='GradeU' "
            Case enuMCGradeType.eGradeV
                strSQL = strSQL & "AND ValueType='GradeV' "
            Case Else

        End Select
        strSQL = strSQL & "ORDER BY MCCodeID;"
        Dim clsSQL As New CMySQL
        GetMCDefectMapping = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)

    End Function
    Public Function GetParameterMapping() As DataTable
        Dim strSQL As String
        strSQL = "SELECT paramID,"
        strSQL = strSQL & "param_rttc,"
        strSQL = strSQL & "param_dct,"
        strSQL = strSQL & "param_eh,"
        strSQL = strSQL & "param_v2002,"
        strSQL = strSQL & "param_add,"
        strSQL = strSQL & "param_mul,"
        strSQL = strSQL & "IsNomalize,"
        strSQL = strSQL & "IsEnable "
        strSQL = strSQL & "FROM db_parameter_mapping.parameter_mapping "
        'strSQL = strSQL & "WHERE IsEnable=True "
        strSQL = strSQL & "ORDER BY paramID;"
        Dim clsSQL As New CMySQL
        Dim dtbData As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)

        GetParameterMapping = dtbData
    End Function

    Public Function GetAllParameter() As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM db_parameter_mapping.parameter_mapping ORDER BY paramID;"
        Dim clsSQL As New CMySQL
        Dim dtbData As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)

        GetAllParameter = dtbData
    End Function

    Public Function GetParamFromDatabaseStructure() As DataTable
        Dim strSQL As String
        strSQL = "SHOW COLUMNS FROM db_master.tabfactor_cfadd;"
        Dim clsSQL As New CMySQL
        Dim dtbParam As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbParam.Rows.RemoveAt(0)   'remove tag_id
        dtbParam.Rows.RemoveAt(0)   'remove test_time_bigint
        dtbParam.Rows.RemoveAt(0)   'remove test_time
        dtbParam.Rows.RemoveAt(0)   'remove tester
        GetParamFromDatabaseStructure = dtbParam
        dtbParam.Columns(0).ColumnName = "param_rttc"
    End Function

    Private Function GetHeaderFromDatabaseStructure() As DataTable
        Dim strSQL As String
        strSQL = "SHOW COLUMNS FROM db_master.tabdetail_header;" 'ORDER BY field;"
        Dim clsSQL As New CMySQL
        Dim dtbParam As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbParam.Rows.RemoveAt(0)   'remove tag_id
        dtbParam.Rows.RemoveAt(0)   'remove test_time_bigint
        GetHeaderFromDatabaseStructure = dtbParam
    End Function

    Public Function GetHeaderDetail() As DataTable
        Dim dtbHeadStruct As DataTable = GetHeaderFromDatabaseStructure()
        Dim strSQL As String
        strSQL = "SELECT * FROM db_parameter_mapping.tabheader_detail;" 'ORDER BY field;"
        Dim clsSQL As New CMySQL
        Dim dtbTemp As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)

        Dim dtbHeader As DataTable = dtbTemp.Clone
        For nHead As Integer = 0 To dtbHeadStruct.Rows.Count - 1      'Select only exist header name
            Dim strHeadName As String = dtbHeadStruct.Rows(nHead).Item("Field")
            Dim drHead() As DataRow = dtbTemp.Select("HeaderName='" & strHeadName & "'")
            If drHead.Length > 0 Then dtbHeader.Rows.Add(drHead(0).ItemArray)
        Next
        GetHeaderDetail = dtbHeader
    End Function

    Public Function GetSliderSite() As DataTable
        Dim strSQL As String = "SELECT * FROM db_parameter_mapping.tabslidersite;"
        Dim clsMySQL As New CMySQL
        GetSliderSite = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetParamByProduct(ByVal strProduct As String, Optional ByVal bGetAll As Boolean = False, Optional ByVal eOrderBy As enuOrderParameterBy = enuOrderParameterBy.eByOrderID, Optional ByVal bGetOnlyParamCF As Boolean = False) As DataTable
        Dim strSQL As String = ""
        Dim clsMySQL As New CMySQL
        Dim dtbParam As DataTable = Nothing

        strSQL = strSQL & "SELECT *,"
        strSQL = strSQL & "IF(param_rttc LIKE 'para%' AND parammachine<>'' AND parammachine IS NOT NULL,REPLACE(REPLACE(CONCAT(parammachine,'.',param_rttc),'[','('),']',')'),param_rttc) param_display "
        strSQL = strSQL & " FROM db_" & strProduct & ".tabparameterbyproduct A "
        If bGetAll Then

        Else
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "A.IsEnable =True "
            strSQL = strSQL & "AND A.parammachine <> '' "
        End If
        If bGetOnlyParamCF Then
            strSQL = strSQL & "AND (param_add=True OR param_mul=True) "
        End If
        If eOrderBy = enuOrderParameterBy.eByOrderID Then
            strSQL = strSQL & "ORDER BY paraOrder;"
        ElseIf eOrderBy = enuOrderParameterBy.eByParamID Then
            strSQL = strSQL & "ORDER BY paramID;"
        Else
            strSQL = strSQL & "ORDER BY param_rttc;"
        End If
        dtbParam = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbParam.TableName = strProduct
        GetParamByProduct = dtbParam
    End Function

    Public Function GetCalculatorByTester(ByVal strProduct As String) As DataTable
        Dim strSQL As String = "SELECT * FROM db_" & strProduct & ".tabparameter_calculatorbytester "
        strSQL = strSQL & "WHERE Tester IN "
        strSQL = strSQL & "(SELECT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "ORDER BY CalculatorName,Tester,Shoe;"
        Dim clsMySQL As New CMySQL
        GetCalculatorByTester = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetCalculatorByProductSetting(ByVal strProduct As String) As DataTable
        Dim strSQL As String = "SELECT * FROM db_" & strProduct & ".tabparameter_calculatorbyProduct "
        strSQL = strSQL & "WHERE CalculatorName IN "
        strSQL = strSQL & "(SELECT DISTINCT CalculatorName FROM db_" & strProduct & ".tabparameter_calculatorbytester "
        strSQL = strSQL & "WHERE Tester IN "
        strSQL = strSQL & "(SELECT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "')) "
        strSQL = strSQL & "ORDER BY CalculatorName;"
        Dim clsMySQL As New CMySQL
        GetCalculatorByProductSetting = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetCalculationParameter(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT * FROM db_" & strProduct & ".tabparameter_calculator;"
        Dim clsMySQL As New CMySQL
        GetCalculationParameter = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetProductList(ByVal nProductType As enuProductType) As DataTable

        Dim strWordFind As String = ""
        Select Case nProductType
            Case enuProductType.enuProductAll
                strWordFind = ""
            Case enuProductType.enuProductFastTrack
                strWordFind = "fasttrack"
            Case enuProductType.enuProductXLot
                strWordFind = "xlot"
            Case enuProductType.enuProductSDET
                strWordFind = "_sdet"
            Case enuProductType.enuProductNPL
                strWordFind = "_npl"
            Case enuProductType.enuProductSTD
                strWordFind = "_std_"
            Case enuProductType.enuProductDOE
                strWordFind = "_doe_"
        End Select

        Dim clsProduct As New CDatabaseManage(m_MysqlConn)
        Dim dtbTemp As DataTable = clsProduct.GetDatabaseList()
        Dim strProduct As String
        Dim dtbProduct As New DataTable("ProductName")
        dtbProduct.Columns.Add("Product")
        For nProduct As Integer = 0 To dtbTemp.Rows.Count - 1
            strProduct = dtbTemp.Rows(nProduct).Item(0)
            If Split(strProduct, "_")(0) = "db" And strProduct.ToLower <> "db_master" And strProduct.ToLower <> "db_parameter_mapping" _
            And InStr(strProduct, strWordFind) Then
                dtbProduct.Rows.Add(UCase(Replace(strProduct, "db_", "")))
            End If
        Next nProduct
        GetProductList = dtbProduct

    End Function

    Public Function GetAllParamByProduct(ByVal dtbProduct As DataTable) As DataSet

        Dim strSQL As String = ""
        Dim clsMySQL As New CMySQL
        Dim dtsParam As DataSet = Nothing
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item("Product")
            strSQL = strSQL & "SELECT *,"
            strSQL = strSQL & "IF(param_rttc LIKE 'para%' AND parammachine<>'' AND parammachine IS NOT NULL,REPLACE(REPLACE(CONCAT(parammachine,'.',param_rttc),'[','('),']',')'),param_rttc) param_display "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabparameterbyproduct A "
            strSQL = strSQL & "ORDER BY paraOrder;"
            'strSQL = strSQL & " WHERE A.parammachine <> '';"
        Next nProduct
        If strSQL <> "" Then
            dtsParam = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)
            For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
                Dim dtbParam As DataTable = dtsParam.Tables(nProduct)
                dtbParam.TableName = dtbProduct.Rows(nProduct).Item("Product")
            Next nProduct
        End If
        GetAllParamByProduct = dtsParam
    End Function

    Public Function GetParamOrderbyStructure(ByVal strProduct As String) As DataTable
        Dim dtbParamStructure As DataTable = GetParamFromDatabaseStructure(strProduct)
        Dim dtbParamByProduct As DataTable = GetParamByProduct(strProduct, True, enuOrderParameterBy.eByName)

        Dim dtbParam As New DataTable(strProduct)
        dtbParam.Merge(dtbParamByProduct)
        dtbParam.Rows.Clear()

        For nParam As Integer = 0 To dtbParamStructure.Rows.Count - 1
            Dim strParamStructure As String = dtbParamStructure.Rows(nParam).Item("Field")
            Dim dtrData() As DataRow = dtbParamByProduct.Select("param_rttc='" & strParamStructure & "'")
            If dtrData.Length > 0 Then
                dtbParam.Rows.Add(dtrData(0).ItemArray)
            End If
        Next nParam
        GetParamOrderbyStructure = dtbParam

    End Function

    Public Function GetParamFromDatabaseStructure(ByVal strProduct As String) As DataTable
        Dim strSQL As String
        strSQL = "SHOW COLUMNS FROM db_" & strProduct & ".tabfactor_cfadd "
        Dim clsSQL As New CMySQL
        Dim dtbParam As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        'dtbParam.Rows.RemoveAt(0)   'remove tag_id
        dtbParam.Rows.RemoveAt(0)   'remove test_time
        dtbParam.Rows.RemoveAt(0)   'remove tester
        dtbParam.Rows.RemoveAt(0)   'remove tester
        dtbParam.Rows.RemoveAt(0)   'remove tester
        GetParamFromDatabaseStructure = dtbParam
    End Function

    Public Function GetSearchListByDate(ByVal strProduct As String, ByVal dtStart As DateTime, ByVal dtEnd As DateTime, ByVal eSearchOption As enumSearchOption, Optional ByVal strKeySearch As String = "", Optional ByVal strMachineType As String = "") As DataTable

        Dim strSQL As String = ""
        Dim clsMySQL As New CMySQL
        Dim dtbSearch As DataTable = Nothing
        Select Case eSearchOption
            Case enumSearchOption.eSearchByTester
                strSQL = "SELECT Tester,Spec FROM db_" & strProduct & ".tabtester "
                strSQL = strSQL & "WHERE test_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL = strSQL & "AND tester<>'' "
                If strKeySearch <> "" Then strSQL = strSQL & "AND Tester LIKE '%" & strKeySearch & "%' "
                If strMachineType <> "" Then strSQL = strSQL & "AND RIGHT(Spec,1)='" & strMachineType & "' "
                strSQL = strSQL & "GROUP BY Tester "
                strSQL = strSQL & "ORDER BY Tester;"
                dtbSearch = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
                dtbSearch.TableName = "Tester"
            Case enumSearchOption.eSearchByLot
                strSQL = "SELECT Lot,Spec FROM db_" & strProduct & ".tabtester "
                strSQL = strSQL & "WHERE test_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL = strSQL & "AND Lot<>'' "
                If strKeySearch <> "" Then strSQL = strSQL & "AND Lot LIKE '%" & strKeySearch & "%' "
                If strMachineType <> "" Then strSQL = strSQL & "AND RIGHT(Spec,1)='" & strMachineType & "' "
                strSQL = strSQL & "GROUP BY Lot "
                strSQL = strSQL & " ORDER BY Lot;"
                dtbSearch = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
                dtbSearch.TableName = "Lot"
            Case enumSearchOption.eSearchBySpec
                strSQL = "SELECT Spec FROM db_" & strProduct & ".tabtester "
                strSQL = strSQL & " WHERE test_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL = strSQL & "AND Spec<>'' "
                If strKeySearch <> "" Then strSQL = strSQL & "AND Spec LIKE '%" & strKeySearch & "%' "
                If strMachineType <> "" Then strSQL = strSQL & "AND RIGHT(Spec,1)='" & strMachineType & "' "
                strSQL = strSQL & "GROUP BY Spec "
                strSQL = strSQL & " ORDER BY Spec;"
                dtbSearch = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
                dtbSearch.TableName = "Spec"
            Case enumSearchOption.eSearchByMachineType
                dtbSearch = New DataTable("MachineType")
                dtbSearch.Columns.Add("OptionIndex", GetType(Int16))
                dtbSearch.Columns.Add("MachineType")
                dtbSearch.Columns.Add("Spec")
                dtbSearch.Rows.Add(enumMachineType.eTypeUp, "Type A", "XXXA")
                dtbSearch.Rows.Add(enumMachineType.eTypeDown, "Type B", "XXXB")
            Case enumSearchOption.eSearchByWafer
                strSQL = "SELECT LEFT(Lot,4) Wafer,Spec FROM db_" & strProduct & ".tabtester "
                strSQL = strSQL & " WHERE test_time BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "' AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "'"
                strSQL = strSQL & "AND Lot<>'' "
                If strKeySearch <> "" Then strSQL = strSQL & "AND LEFT(Lot,4) LIKE '%" & strKeySearch & "%' "
                If strMachineType <> "" Then strSQL = strSQL & "AND RIGHT(Spec,1)='" & strMachineType & "' "
                strSQL = strSQL & "GROUP BY LEFT(Lot,4) "
                strSQL = strSQL & " ORDER BY Lot;"
                dtbSearch = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
                dtbSearch.TableName = "Wafer"
        End Select
        GetSearchListByDate = dtbSearch
    End Function

    Public Function GetMainProduct() As DataTable

        Dim strWordFind As String = ""

        Dim strSQL As String = "SHOW DATABASES"
        Dim clsMyProduct As New CMySQL
        Dim dtbDatabase As DataTable = clsMyProduct.CommandMySqlDataTable(strSQL, m_MysqlConn)
        Dim dtbProduct As New DataTable("ProductName")
        dtbProduct.Columns.Add("Product")
        Dim strProduct As String
        For nProduct As Integer = 0 To dtbDatabase.Rows.Count - 1
            strProduct = dtbDatabase.Rows(nProduct).Item(0)
            If Split(strProduct, "_")(0) = "db" And strProduct.ToLower <> "db_master" And strProduct.ToLower <> "db_parameter_mapping" _
            And InStr(strProduct, strWordFind) Then
                dtbProduct.Rows.Add(UCase(Replace(strProduct, "db_", "")))
            End If
        Next nProduct

        Dim dtbMainProduct As New DataTable

        dtbMainProduct.Columns.Add("Product")
        dtbMainProduct.Columns("Product").Unique = True

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strTemp() As String = Split(dtbProduct.Rows(nProduct).Item("Product"), "_")
            Dim strMainProduct As String = strTemp(0)

            Dim dtrProduct() As DataRow = dtbMainProduct.Select("Product='" & strMainProduct & "'", "Product")
            If dtrProduct.Length = 0 Then dtbMainProduct.Rows.Add(strMainProduct)
        Next nProduct

        GetMainProduct = dtbMainProduct
    End Function

    Public Function IsProductExist(ByVal strProduct As String) As Boolean
        IsProductExist = False
        strProduct = strProduct.ToUpper
        Dim clsProduct As New CParameterRTTCMapping(m_MysqlConn)
        Dim dtbProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductAll)
        Dim drProduct() As DataRow = dtbProduct.Select("Product='" & strProduct & "'")
        If drProduct.Length > 0 Then IsProductExist = True
    End Function

    Public Function GetSTDParameter() As DataTable
        Dim strSQL As String = "SELECT A.paramID,"
        strSQL = strSQL & "B.param_rttc,"
        'strSQL = strSQL & "B.parammachine,"
        'strSQL = strSQL & "IF(B.param_rttc LIKE 'para%' AND B.parammachine<>'' AND B.parammachine IS NOT NULL,REPLACE(REPLACE(CONCAT(B.parammachine,'.',B.param_rttc),'[','('),']',')'),B.param_rttc) param_display,"
        strSQL = strSQL & "A.AdjustGOS,"
        strSQL = strSQL & "A.CFTypeID "
        strSQL = strSQL & "FROM std_standard.tabstdparameter A "
        strSQL = strSQL & "INNER JOIN db_parameter_mapping.parameter_mapping B ON A.paramID=B.paramID "
        strSQL = strSQL & "ORDER BY A.STDOrder;"
        Dim clsMySql As New CMySQL
        GetSTDParameter = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetSPGSTDParameter() As DataTable      ' Get SPG Header
        Dim strSQL As String = "SELECT paramID,"
        strSQL = strSQL & "param_rttc "
        'strSQL = strSQL & "A.AdjustGOS,"
        'strSQL = strSQL & "A.CFTypeID "
        strSQL = strSQL & "FROM std_standard.tabspgparameter  "
        strSQL = strSQL & "ORDER BY STDOrder;"
        Dim clsMySql As New CMySQL
        GetSPGSTDParameter = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetParamBySPGProduct(ByVal strProduct As String, Optional ByVal bGetAll As Boolean = False, Optional ByVal eOrderBy As enuOrderParameterBy = enuOrderParameterBy.eByOrderID, Optional ByVal bGetOnlyParamCF As Boolean = False) As DataTable
        Dim strSQL As String = ""
        Dim clsMySQL As New CMySQL
        Dim dtbParam As DataTable = Nothing

        strSQL = strSQL & "SELECT *,"
        strSQL = strSQL & "IF(param_rttc LIKE 'para%' AND parammachine<>'' AND parammachine IS NOT NULL,REPLACE(REPLACE(CONCAT(parammachine,'.',param_rttc),'[','('),']',')'),param_rttc) param_display "
        strSQL = strSQL & " FROM std_standard.tabspgparameter A " 'db_" & strProduct & ".tabparameterbyproduct A "
        If bGetAll Then

        Else
            strSQL = strSQL & " WHERE "
            '  strSQL = strSQL & "A.IsEnable =True "
            strSQL = strSQL & "A.parammachine <> '' " '"AND A.parammachine <> '' "
        End If
        'If bGetOnlyParamCF Then
        '    strSQL = strSQL & "AND (param_add=True OR param_mul=True) "
        'End If
        If eOrderBy = enuOrderParameterBy.eByOrderID Then
            strSQL = strSQL & "ORDER BY STDOrder;"
        ElseIf eOrderBy = enuOrderParameterBy.eByParamID Then
            strSQL = strSQL & "ORDER BY paramID;"
        Else
            strSQL = strSQL & "ORDER BY param_rttc;"
        End If
        dtbParam = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbParam.TableName = strProduct
        GetParamBySPGProduct = dtbParam
    End Function
End Class

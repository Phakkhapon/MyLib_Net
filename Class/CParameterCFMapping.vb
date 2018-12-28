Imports MySql.Data.MySqlClient

Public Class CParameterCFMapping

    Private m_MysqlConn As MySqlConnection

    Public Sub New(ByVal MysqlConn As MySqlConnection)
        m_MysqlConn = MysqlConn
    End Sub

    Public Function GetParameterMapping(ByVal strProduct As String) As DataTable
        Dim strSQL As String
        strSQL = "SELECT A.paramID,"
        strSQL = strSQL & "A.param_rttc,"
        strSQL = strSQL & "C.paramMDB,"
        strSQL = strSQL & "C.MachineCF,"
        strSQL = strSQL & "IFNULL(B.CFMediaType,FALSE) CFMediaType,"
        strSQL = strSQL & "IF(A.param_rttc LIKE 'para%',IF(C.paramMDB='' OR C.paramMDB IS NULL,A.param_rttc,CONCAT(C.paramMDB,'.',A.param_rttc)),A.param_rttc) param_display,"
        strSQL = strSQL & "C.Zone,"
        strSQL = strSQL & "C.param_add,"
        strSQL = strSQL & "C.param_mul "
        strSQL = strSQL & "FROM db_parameter_mapping.parameter_mapping A "
        strSQL = strSQL & "LEFT JOIN db_parameter_mapping.tabcfmediatype B USING(paramID) "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabparameterbyproduct C USING(param_rttc) "
        strSQL = strSQL & "ORDER BY A.paramID;"
        Dim clsSQL As New CMySQL
        Dim dtbData As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        GetParameterMapping = dtbData
    End Function

    Public Sub SetTesterSendLimit(ByVal strTester As String, ByVal bSetTo As Boolean)
        Dim strSQL As String = "UPDATE db_parameter_mapping.tabmachinebyproduct "
        strSQL = strSQL & "SET SendLimit=" & bSetTo.ToString & " "
        strSQL = strSQL & "WHERE Tester='" & strTester & "';"
        Dim clsSQL As New CMySQL
        clsSQL.CommandNoQuery(strSQL, m_MysqlConn)
    End Sub

    Public Function GetTesterInProduct(ByVal strProduct As String, ByVal strShoe As String) As DataTable
        Dim strSQL As String = ""

        strSQL = strSQL & "SELECT Tester,Shoe,MediaSN,SendLimit FROM db_parameter_mapping.tabmachinebyproduct "
        strSQL = strSQL & "WHERE ProductName='" & strProduct & "' "
        strSQL = strSQL & "AND Shoe='" & strShoe & "' "
        'strSQL = strSQL & "AND UpdateTime>'" & Format(Now.AddDays(-30), "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "ORDER BY Tester;"

        Dim clsMySQL As New CMySQL
        Dim dtbTester As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbTester.TableName = strProduct
        GetTesterInProduct = dtbTester
    End Function

    Private ReadOnly Property GetSQLCFLimitByTesterSetting(ByVal strProduct As String) As String
        Get
            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "SettingDate,"
            strSQL = strSQL & "ControlType,"
            strSQL = strSQL & "paramID,"
            strSQL = strSQL & "CFTypeID,"
            strSQL = strSQL & "LimitMin,"
            strSQL = strSQL & "LimitMax,"
            strSQL = strSQL & "LimitRange,"
            strSQL = strSQL & "CollectDay,"
            strSQL = strSQL & "CollectPoint,"
            strSQL = strSQL & "EmailID,"
            strSQL = strSQL & "SettingUser "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabctr_cflimitbytester A "
            strSQL = strSQL & "WHERE SettingDate=(SELECT MAX(settingdate) FROM db_" & strProduct & ".tabctr_cflimitbytester WHERE ParamID=A.ParamID AND CFTypeID=A.CFTypeID) "
            strSQL = strSQL & "GROUP BY A.ParamID,A.CFTypeID;"
            GetSQLCFLimitByTesterSetting = strSQL
        End Get
    End Property

    Public ReadOnly Property GetCFLimitByTesterSetting(ByVal strProduct As String) As DataTable
        Get
            Dim strSQL As String = GetSQLCFLimitByTesterSetting(strProduct)
            Dim clsMySQL As New CMySQL
            GetCFLimitByTesterSetting = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        End Get
    End Property

    Public Function GetAllCFLimitByTesterSetting(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & GetSQLCFLimitByTesterSetting(strProduct)
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetAllCFLimitByTesterSetting = dtsSetting
    End Function

    Private ReadOnly Property GetSqlCFLimitByProduct(ByVal strProduct As String) As String
        Get
            Dim strSQL As String = ""
            strSQL = strSQL & "SELECT Date_time,"
            strSQL = strSQL & "CFLimitTypeID,"
            strSQL = strSQL & "Tester,"
            strSQL = strSQL & "Shoe,"
            strSQL = strSQL & "MediaSN,"
            strSQL = strSQL & "paramID,"
            strSQL = strSQL & "IF(B.param_rttc LIKE 'para%',IF(B.paramMDB='' OR B.paramMDB IS NULL,B.param_rttc,CONCAT(B.paramMDB,'.',B.param_rttc)),B.param_rttc) param_display,"
            strSQL = strSQL & "CFTypeID,"
            strSQL = strSQL & "LimitValueMin,"
            strSQL = strSQL & "LimitValueMax,"
            strSQL = strSQL & "CenterRef "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabcflimit A "
            strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabparameterbyproduct B USING(paramID) "
            strSQL = strSQL & "WHERE date_time="
            strSQL = strSQL & "(SELECT MAX(date_time) FROM db_" & strProduct & ".tabcflimit "
            strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND CFTypeID=A.CFTypeID AND IsEnable=True) "
            strSQL = strSQL & "AND A.IsEnable=True "
            strSQL = strSQL & "AND Tester IN (SELECT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "AND ParamID IN (SELECT paramID FROM db_parameter_mapping.parameter_mapping) "
            strSQL = strSQL & "GROUP BY tester,shoe,paramID,CFTypeID "
            strSQL = strSQL & "ORDER BY tester,shoe,paramID;"
            GetSqlCFLimitByProduct = strSQL
        End Get
    End Property

    Public ReadOnly Property GetCFLimitByProduct(ByVal strProduct As String) As DataTable
        Get
            Dim strSQL As String = GetSqlCFLimitByProduct(strProduct)
            Dim clsMySQL As New CMySQL
            GetCFLimitByProduct = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        End Get
    End Property

    Public Function GetAllCFLimitByProduct(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & GetSqlCFLimitByProduct(strProduct)
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetAllCFLimitByProduct = dtsSetting
    End Function

    Public Function GetCFLimitNowByTester(ByVal strProduct As String, ByVal strTester As String) As DataTable
        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT Date_time,"
        strSQL = strSQL & "CFLimitTypeID,"
        strSQL = strSQL & "tester,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "paramID,"
        strSQL = strSQL & "CFTypeID,"
        strSQL = strSQL & "LimitValueMin,"
        strSQL = strSQL & "LimitValueMax,"
        strSQL = strSQL & "CenterRef "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabcflimit A "
        strSQL = strSQL & "WHERE date_time="
        strSQL = strSQL & "(SELECT MAX(date_time) FROM db_" & strProduct & ".tabcflimit "
        strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND CFTypeID=A.CFTypeID AND IsEnable=True) "
        strSQL = strSQL & "AND IsEnable=True "
        strSQL = strSQL & "AND Tester='" & strTester & "' "
        strSQL = strSQL & "AND Tester IN (SELECT DISTINCT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "AND ParamID IN (SELECT DISTINCT paramID FROM db_parameter_mapping.parameter_mapping) "
        strSQL = strSQL & "GROUP BY tester,shoe,paramID,CFTypeID "
        strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        Dim clsMySQL As New CMySQL
        Dim dtbLimit As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        GetCFLimitNowByTester = dtbLimit
    End Function


    Public Function GetCFNow(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM cf_" & strProduct & ".tabcfnow A "
            strSQL = strSQL & "WHERE Tester IN (SELECT DISTINCT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetCFNow = dtsSetting
    End Function

    Public Function GetCFNowByProduct(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabcfnow A "
        strSQL = strSQL & "WHERE Tester IN (SELECT Tester FROM db_parameter_mapping.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Dim clsMySQL As New CMySQL
        Dim dtbSetting As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbSetting.TableName = strProduct
        GetCFNowByProduct = dtbSetting
    End Function

    Public Function IsProductExist(ByVal strProduct As String) As Boolean
        IsProductExist = False
        strProduct = strProduct.ToUpper
        Dim clsProduct As New CParameterRTTCMapping(m_MysqlConn)
        Dim dtbProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductAll)
        Dim drProduct() As DataRow = dtbProduct.Select("Product='" & strProduct & "'")
        If drProduct.Length > 0 Then IsProductExist = True
    End Function

    Public Function GetAllTesterEvent(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        Dim dtbProductReturn As DataTable = dtbProduct.Clone
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM db_" & strProduct & ".tabcftester_event "
            strSQL = strSQL & "WHERE IsActionDone=False "
            strSQL = strSQL & "ORDER BY Tester,Shoe,Date_time;"
            'strSQL = strSQL & "UPDATE db_" & strProduct & ".tabcftester_event SET IsActionDone=True;"
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsTemp As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        'Dim dtsTesterEvent As New DataSet
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item("Product")
            dtsTemp.Tables(nProduct).TableName = strProduct
            If dtsTemp.Tables(nProduct).Rows.Count > 0 Then
                dtbProductReturn.Rows.Add(strProduct)
            End If
        Next
        dtsTemp.Tables.Add(dtbProductReturn)
        GetAllTesterEvent = dtsTemp
    End Function

    Public Sub UpdateActionDone(ByVal strProduct As String, ByVal drEvent As DataRow)
        Dim strSQL As String = "UPDATE db_" & strProduct & ".tabcftester_event SET IsActionDone=True "
        strSQL = strSQL & "WHERE Date_time='" & Format(drEvent.Item("Date_time"), "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "AND Tester='" & drEvent.Item("Tester") & "' "
        strSQL = strSQL & "AND Shoe='" & drEvent.Item("Shoe") & "' "
        strSQL = strSQL & "AND CFTypeID='" & drEvent.Item("CFTypeID") & "' "
        strSQL = strSQL & "AND ParameterID='" & drEvent.Item("ParameterID") & "';"
        Dim clsMySql As New CMySQL
        clsMySql.CommandNoQuery(strSQL, m_MysqlConn)
    End Sub

    Public Function GetMachineDetail() As DataTable
        Dim strSQL As String = "SELECT * FROM db_parameter_mapping.tabmachinebyproduct ORDER BY Tester;"
        Dim clsMySql As New CMySQL
        GetMachineDetail = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public ReadOnly Property GetTesterAndIPAdr() As DataTable
        Get
            Dim strSQL As String = "SELECT * FROM db_parameter_mapping.tabmachinebyproduct ORDER BY Tester;"
            Dim clsMySql As New CMySQL
            GetTesterAndIPAdr = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
        End Get
    End Property

    Public ReadOnly Property GetIPAdrByTester(ByVal strTester As String) As String
        Get
            GetIPAdrByTester = ""
            Dim strSQL As String = "SELECT * FROM db_parameter_mapping.tabmachinebyproduct WHERE Tester='" & strTester & "' ORDER BY Tester;"
            Dim clsMySql As New CMySQL
            Dim dtbIP As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
            If dtbIP.Rows.Count > 0 Then
                GetIPAdrByTester = dtbIP.Rows(0).Item("IPAdr")
            End If
        End Get
    End Property

End Class

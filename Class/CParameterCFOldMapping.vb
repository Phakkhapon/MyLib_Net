Imports MySql.Data.MySqlClient

Public Class CParameterCFOldMapping

    Private m_MysqlConn As MySqlConnection

    Public Sub New(ByVal MysqlConn As MySqlConnection)
        m_MysqlConn = MysqlConn
    End Sub

    Public Function GetParameterMapping() As DataTable
        Dim strSQL As String
        strSQL = "SELECT paramID,"
        strSQL = strSQL & "param_rttc "
        strSQL = strSQL & "FROM db_parameter.tabparamapping "
        strSQL = strSQL & "ORDER BY paramID;"
        Dim clsSQL As New CMySQL
        Dim dtbData As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        GetParameterMapping = dtbData
    End Function

    Public Function GetTesterInProduct(ByVal strProduct As String, ByVal strShoe As String) As DataTable
        Dim strSQL As String = ""

        strSQL = strSQL & "SELECT Tester,Shoe,MediaSN,SendLimit FROM db_machine.tabmachinebyproduct "
        strSQL = strSQL & "WHERE ProductName='" & strProduct & "' "
        strSQL = strSQL & "AND Shoe='" & strShoe & "' "
        strSQL = strSQL & "AND UpdateTime>'" & Format(Now.AddDays(-30), "yyyy-MM-dd HH:mm:ss") & "' "
        strSQL = strSQL & "ORDER BY Tester;"

        Dim clsMySQL As New CMySQL
        Dim dtbTester As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbTester.TableName = strProduct
        GetTesterInProduct = dtbTester
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
        End Select

        Dim strSQL As String = "SHOW DATABASES"
        Dim clsMyProduct As New CMySQL
        Dim dtbDatabase As DataTable = clsMyProduct.CommandMySqlDataTable(strSQL, m_MysqlConn)
        Dim dtbProduct As New DataTable("ProductName")
        dtbProduct.Columns.Add("Product")
        Dim strProduct As String
        For nProduct As Integer = 0 To dtbDatabase.Rows.Count - 1
            strProduct = dtbDatabase.Rows(nProduct).Item(0)
            If Split(strProduct, "_")(0) = "cf" And InStr(strProduct, strWordFind) Then
                dtbProduct.Rows.Add(UCase(Replace(strProduct, "cf_", "")))
            End If
        Next nProduct
        GetProductList = dtbProduct
    End Function

    Public Function GetAllCFLimitByTesterSetting(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT * FROM cf_" & strProduct & ".tabctr_cflimitbytester A "
            strSQL = strSQL & "LEFT JOIN db_Parameter.tabparamapping B USING(ParamID) "
            strSQL = strSQL & "WHERE SettingDate=(SELECT MAX(settingdate) FROM cf_" & strProduct & ".tabctr_cflimitbytester WHERE ParamID=A.ParamID) "
            strSQL = strSQL & "GROUP BY ParamID,CFType;"
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetAllCFLimitByTesterSetting = dtsSetting
    End Function

    Public Function GetAllCFLimitByProductAdd(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT Date_time,"
            strSQL = strSQL & "CFLimitTypeID,"
            strSQL = strSQL & "tester,"
            strSQL = strSQL & "Shoe,"
            strSQL = strSQL & "MediaSN,"
            strSQL = strSQL & "paramID,"
            strSQL = strSQL & "LimitValueMin,"
            strSQL = strSQL & "LimitValueMax,"
            strSQL = strSQL & "CenterRef "
            strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_add A "
            strSQL = strSQL & "WHERE date_time="
            strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_add "
            strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID) "
            strSQL = strSQL & "AND Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "GROUP BY tester,shoe,paramID "
            strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetAllCFLimitByProductAdd = dtsSetting

    End Function

    Public Function GetCFLimitNowByTester(ByVal strProduct As String, ByVal strTester As String) As DataSet
        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT Date_time,"
        strSQL = strSQL & "CFLimitTypeID,"
        strSQL = strSQL & "tester,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "paramID,"
        strSQL = strSQL & "Param_rttc,"
        strSQL = strSQL & "LimitValueMin,"
        strSQL = strSQL & "LimitValueMax,"
        strSQL = strSQL & "CenterRef "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_add A "
        strSQL = strSQL & "LEFT JOIN db_parameter.tabparamapping B USING(paramID) "
        strSQL = strSQL & "WHERE date_time="
        strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_add "
        strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND IsEnable=True) "
        strSQL = strSQL & "AND IsEnable=True "
        strSQL = strSQL & "AND Tester='" & strTester & "' "
        strSQL = strSQL & "GROUP BY tester,shoe,paramID "
        strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        strSQL = strSQL & "SELECT Date_time,"
        strSQL = strSQL & "CFLimitTypeID,"
        strSQL = strSQL & "tester,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "paramID,"
        strSQL = strSQL & "Param_rttc,"
        strSQL = strSQL & "LimitValueMin,"
        strSQL = strSQL & "LimitValueMax,"
        strSQL = strSQL & "CenterRef "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_mul A "
        strSQL = strSQL & "LEFT JOIN db_parameter.tabparamapping B USING(paramID) "
        strSQL = strSQL & "WHERE date_time="
        strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_mul "
        strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND IsEnable=True) "
        strSQL = strSQL & "AND IsEnable=True "
        strSQL = strSQL & "AND Tester='" & strTester & "' "
        strSQL = strSQL & "GROUP BY tester,shoe,paramID "
        strSQL = strSQL & "ORDER BY tester,shoe,paramID;"
        Dim clsMySQL As New CMySQL
        Dim dtsLimit As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)
        dtsLimit.Tables(0).TableName = "Add"
        dtsLimit.Tables(1).TableName = "Mul"
        GetCFLimitNowByTester = dtsLimit
    End Function

    Public Function GetAllCFLimitByProductMul(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT Date_time,"
            strSQL = strSQL & "CFLimitTypeID,"
            strSQL = strSQL & "tester,"
            strSQL = strSQL & "Shoe,"
            strSQL = strSQL & "MediaSN,"
            strSQL = strSQL & "paramID,"
            strSQL = strSQL & "LimitValueMin,"
            strSQL = strSQL & "LimitValueMax,"
            strSQL = strSQL & "CenterRef "
            strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_mul A "
            strSQL = strSQL & "WHERE date_time="
            strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_mul "
            strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID) "
            strSQL = strSQL & "AND A.Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "GROUP BY tester,shoe,paramID "
            strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetAllCFLimitByProductMul = dtsSetting

    End Function

    Public Function GetCFLimitByProductAdd(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT Date_time,"
        strSQL = strSQL & "CFLimitTypeID,"
        strSQL = strSQL & "tester,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "paramID,"
        strSQL = strSQL & "LimitValueMin,"
        strSQL = strSQL & "LimitValueMax,"
        strSQL = strSQL & "CenterRef "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_add A "
        strSQL = strSQL & "WHERE date_time="
        strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_add "
        strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND IsEnable=True) "
        strSQL = strSQL & "AND IsEnable=True "
        strSQL = strSQL & "AND A.Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "GROUP BY tester,shoe,paramID "
        strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        Dim clsMySQL As New CMySQL
        GetCFLimitByProductAdd = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetCFLimitByProductMul(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""
        strSQL = strSQL & "SELECT Date_time,"
        strSQL = strSQL & "CFLimitTypeID,"
        strSQL = strSQL & "tester,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "MediaSN,"
        strSQL = strSQL & "paramID,"
        strSQL = strSQL & "LimitValueMin,"
        strSQL = strSQL & "LimitValueMax,"
        strSQL = strSQL & "CenterRef "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcflimit_mul A "
        strSQL = strSQL & "WHERE date_time="
        strSQL = strSQL & "(SELECT MAX(date_time) FROM cf_" & strProduct & ".tabcflimit_mul "
        strSQL = strSQL & "WHERE tester=A.tester AND shoe=A.shoe AND paramID=A.ParamID AND IsEnable=True) "
        strSQL = strSQL & "AND IsEnable=True "
        strSQL = strSQL & "AND A.Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "GROUP BY tester,shoe,paramID "
        strSQL = strSQL & "ORDER BY tester,shoe,paramID;"

        Dim clsMySQL As New CMySQL
        GetCFLimitByProductMul = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetCFAddNow(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM cf_" & strProduct & ".tabcfnow_add A "
            strSQL = strSQL & "WHERE Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetCFAddNow = dtsSetting
    End Function

    Public Function GetCFMulNow(ByVal dtbProduct As DataTable) As DataSet
        Dim strSQL As String = ""
        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "FROM cf_" & strProduct & ".tabcfnow_mul A "
            strSQL = strSQL & "WHERE Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
            strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Next nProduct
        Dim clsMySQL As New CMySQL
        Dim dtsSetting As DataSet = clsMySQL.CommandMySqlDataset(strSQL, m_MysqlConn)

        For nProduct As Integer = 0 To dtbProduct.Rows.Count - 1
            Dim strProduct As String = dtbProduct.Rows(nProduct).Item(0)
            dtsSetting.Tables(nProduct).TableName = strProduct
        Next
        GetCFMulNow = dtsSetting
    End Function

    Public Function GetCFAddNowByProduct(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""

        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcfnow_add A "
        strSQL = strSQL & "WHERE Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Dim clsMySQL As New CMySQL
        Dim dtbSetting As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbSetting.TableName = strProduct
        GetCFAddNowByProduct = dtbSetting
    End Function

    Public Function GetCFMulNowByProduct(ByVal strProduct As String) As DataTable
        Dim strSQL As String = ""

        strSQL = strSQL & "SELECT * "
        strSQL = strSQL & "FROM cf_" & strProduct & ".tabcfnow_mul A "
        strSQL = strSQL & "WHERE Tester IN (SELECT DISTINCT Tester FROM db_machine.tabmachinebyproduct WHERE ProductName='" & strProduct & "') "
        strSQL = strSQL & "ORDER BY Tester,Shoe;"
        Dim clsMySQL As New CMySQL
        Dim dtbSetting As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_MysqlConn)
        dtbSetting.TableName = strProduct
        GetCFMulNowByProduct = dtbSetting
    End Function

    Public Function IsProductExist(ByVal strProduct As String) As Boolean
        IsProductExist = False
        strProduct = strProduct.ToUpper
        Dim clsProduct As New CParameterRTTCMapping(m_MysqlConn)
        Dim dtbProduct As DataTable = clsProduct.GetProductList(enuProductType.enuProductAll)
        Dim drProduct() As DataRow = dtbProduct.Select("Product='" & strProduct & "'")
        If drProduct.Length > 0 Then IsProductExist = True
    End Function

    Public Function GetTesterAndIPAdr() As DataTable
        Dim strSQL As String = "SELECT * FROM db_Machine.tabmachine ORDER BY Tester;"
        Dim clsMySql As New CMySQL
        GetTesterAndIPAdr = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
    End Function

    Public Function GetIPAdrByTester(ByVal strTester As String) As String
        GetIPAdrByTester = ""
        Dim strSQL As String = "SELECT * FROM db_Machine.tabmachine WHERE Tester='" & strTester & "' ORDER BY Tester;"
        Dim clsMySql As New CMySQL
        Dim dtbIP As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_MysqlConn)
        If dtbIP.Rows.Count > 0 Then
            GetIPAdrByTester = dtbIP.Rows(0).Item("IPAdr")
        End If
    End Function

End Class

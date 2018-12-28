
Imports MySql.Data.MySqlClient

Public Class CFailureCounterSetting
    Private m_cslMySQL As New CMySQL
    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

 
    Public Function GetFailureSettingByProduct(ByVal strProduct As String) As DataTable
        Dim strSQL As String
        strSQL = "SELECT * "
        strSQL = strSQL & " FROM db_" & strProduct & ".tabctr_failurecounter A;"
        'strSQL = strSQL & " LEFT JOIN ctr_controlsetting.tabproductdetail B USING(productID)"
        'strSQL = strSQL & " WHERE A.ProductName='" & strProduct & "';"
        Dim dtbFailureSetting As DataTable = m_cslMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        'dtbFailureSetting.Columns("CompareCondition").SetOrdinal(dtbFailureSetting.Columns("FailCount").Ordinal + 1)
        GetFailureSettingByProduct = dtbFailureSetting
    End Function


    Public Function LoadSkipLot(ByVal strProduct As String, ByVal nFailureID As String) As DataTable
        Dim strSQL As String
        strSQL = "SELECT "
        strSQL = strSQL & "FailureID,"
        strSQL = strSQL & "AddTime,"
        strSQL = strSQL & "Owner,"
        strSQL = strSQL & "LotSkip,"
        strSQL = strSQL & "Reason "
        strSQL = strSQL & "FROM db_" & strProduct & ".tabctr_failureskipLot "
        strSQL = strSQL & "WHERE FailureID=" & nFailureID & " OR FailureID='-1' "
        strSQL = strSQL & "ORDER BY LotSkip;"
        Dim clsMySql As New CMySQL
        LoadSkipLot = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

    Public Function GetAllGradeName() As DataTable
        Dim strSQL As String
        strSQL = "SELECT DISTINCT ResultCount FROM db_parameter_mapping.tabmcdefect "
        strSQL = strSQL & "WHERE SectionName='GradeName' "
        strSQL = strSQL & "ORDER BY ResultCount ASC;"
        Dim clsMySql As New CMySQL
        GetAllGradeName = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
    End Function

End Class

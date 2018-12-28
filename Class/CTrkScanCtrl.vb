
Imports Lib_Net
Imports MySql.Data.MySqlClient
Public Class CTrkScanCtrl
    Private m_mySqlConn As MySqlConnection
    Private m_clsMySQL As New CMySQL

    Public Sub New(ByVal myConn As MySqlConnection)
        m_mySqlConn = myConn
    End Sub

    Public Function getTstList(ProductName As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT Tester FROM db_parameter_mapping.tabmachinebyproduct A "
        strSQL = strSQL & "WHERE ProductName='" & ProductName & "' "
        strSQL = strSQL & "ORDER BY Tester;"
        Dim clsMySQL As New CMySQL
        Dim dtbSearch As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        getTstList = dtbSearch

    End Function
    Public Function getSetting(ProductName As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT Tester FROM db_" & ProductName & ".tabmachinebyproduct A "
        strSQL = strSQL & "WHERE ProductName='" & ProductName & "' "
        strSQL = strSQL & "ORDER BY Tester;"
        Dim clsMySQL As New CMySQL
        Dim dtbSetting As DataTable = clsMySQL.CommandMySqlDataTable(strSQL, m_mySqlConn)
        getSetting = dtbSetting

    End Function


End Class

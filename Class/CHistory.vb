
Imports MySql.Data.MySqlClient
Public Class CHistory
    Private m_mySqlConn As MySqlConnection

    Public Sub New(ByVal mySqlConn As MySqlConnection)
        m_mySqlConn = mySqlConn
    End Sub

    Public Function InsertTesterHistory(ByVal strProduct As String, ByVal strTester As String, ByVal strLot As String, ByVal strSpec As String, ByVal strShoe As String, ByVal strProblemCode As String, ByVal strActionCode As String, ByVal strStart_time As String, ByVal strTest_time As String, ByVal dtSendDate As DateTime, ByVal strSendBy As String, ByVal strCriteria As String, ByVal strDesc As String) As Integer
        InsertTesterHistory = 0
        Dim strSQL As String
        Dim strSendDate = Format(dtSendDate, "yyyy-MM-dd HH:mm:ss")
        Dim strStartDate As String = Format(dtSendDate, "yyyy-MM-dd") & " 00:00:00"
        Dim strEndDate As String = Format(dtSendDate, "yyyy-MM-dd") & " 23:59:59"
        Dim clsMySql As New CMySQL

        strSQL = "INSERT INTO db_" & strProduct & ".testerhistory("
        strSQL = strSQL & "Tester,"
        strSQL = strSQL & "Lot,"
        strSQL = strSQL & "Spec,"
        strSQL = strSQL & "Shoe,"
        strSQL = strSQL & "ProblemCode,"
        strSQL = strSQL & "ActionCode,"
        strSQL = strSQL & "Start_time,"
        strSQL = strSQL & "Test_time,"
        strSQL = strSQL & "SendDate,"
        strSQL = strSQL & "SendBy,"
        strSQL = strSQL & "Criteria,"
        strSQL = strSQL & "Description) VALUES("
        strSQL = strSQL & "'" & strTester & "',"
        strSQL = strSQL & "'" & strLot & "',"
        strSQL = strSQL & "'" & strSpec & "',"
        strSQL = strSQL & "'" & strShoe & "',"
        strSQL = strSQL & "'" & strProblemCode & "',"
        strSQL = strSQL & "'" & strActionCode & "',"
        strSQL = strSQL & "'" & strStart_time & "',"
        strSQL = strSQL & "'" & strTest_time & "',"
        strSQL = strSQL & "'" & strSendDate & "',"
        strSQL = strSQL & "'" & strSendBy & "',"
        strSQL = strSQL & "'" & strCriteria & "',"
        strSQL = strSQL & "'" & strDesc & "');"
        clsMySql.CommandNoQuery(strSQL, m_mySqlConn)
        strSQL = "SELECT count(A.Tester) CriteriaPerDay "
        strSQL = strSQL & " FROM db_" & strProduct & ".testerhistory A "
        strSQL = strSQL & " WHERE A.tester='" & strTester & "' "
        strSQL = strSQL & " AND A.SendDate BETWEEN '" & strStartDate & "' AND '" & strEndDate & "'"
        strSQL = strSQL & " AND A.Criteria='" & strCriteria & "';"

        Dim dtbHistory As DataTable
        dtbHistory = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        InsertTesterHistory = dtbHistory.Rows(0).Item("CriteriaPerDay")

    End Function

    Public Function InsertHistoryAlertLot(ByVal strProduct As String, ByVal strTester As String, ByVal strLot As String, ByVal strSpec As String, ByVal strShoe As String, ByVal strProblemCode As String, ByVal strActionCode As String, ByVal strStart_time As String, ByVal strTest_time As String, ByVal dtSendDate As DateTime, ByVal strSendBy As String, ByVal strCriteria As String, ByVal strDesc As String) As Integer

        InsertHistoryAlertLot = 0

        Dim strSQL As String
        Dim strSendDate = Format(dtSendDate, "yyyy-MM-dd HH:mm:ss")
        Dim strStartDate As String = Format(dtSendDate, "yyyy-MM-dd") & " 00:00:00"
        Dim strEndDate As String = Format(dtSendDate, "yyyy-MM-dd") & " 23:59:59"
        Dim clsMySql As New CMySQL

        strSQL = "SELECT COUNT(A.Tester) CriteriaPerDay FROM db_" & strProduct & ".testerhistory A "
        strSQL = strSQL & "WHERE A.Tester='" & strTester & "' "
        strSQL = strSQL & "AND A.Criteria='" & strCriteria & "' "
        strSQL = strSQL & "AND A.Lot='" & strLot & "' "
        strSQL = strSQL & " AND A.SendDate BETWEEN '" & strStartDate & "' AND '" & strEndDate & "';"
        Dim dtbHistory As DataTable
        dtbHistory = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        Dim nCriteriaPerDay As Integer = dtbHistory.Rows(0).Item("CriteriaPerDay")
        If nCriteriaPerDay = 0 Then
            InsertTesterHistory(strProduct, strTester, strLot, strSpec, strShoe, strProblemCode, strActionCode, strStart_time, strTest_time, dtSendDate, strSendBy, strCriteria, strDesc)
        End If
        InsertHistoryAlertLot = nCriteriaPerDay

    End Function

    Public Function GetTesterHistory(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime, _
        ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM db_" & strProduct & ".testerhistory A "
        strSQL = strSQL & "WHERE ("
        Dim strSearchBy As String = dtbSearchBy.TableName
        If strSearchBy.ToLower = "optionindex" Then
            strSearchBy = "tester"
            For nItem As Integer = 0 To dtbSearchBy.Rows.Count - 1
                Dim strMachineType As String = ""
                If dtbSearchBy.Rows(nItem).Item("OptionIndex") Then
                    strMachineType = "Down"
                Else
                    strMachineType = "Up"
                End If
                If nItem <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & " A." & strSearchBy & "='" & strMachineType & "' OR "
                Else
                    strSQL = strSQL & " A." & strSearchBy & "='" & strMachineType & "') "
                End If
            Next nItem
        Else
            For nItem As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nItem <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & " A." & strSearchBy & "='" & dtbSearchBy.Rows(nItem).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & " A." & strSearchBy & "='" & dtbSearchBy.Rows(nItem).Item(strSearchBy) & "') "
                End If
            Next nItem
        End If

        strSQL = strSQL & " AND (A.SendDate BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " ORDER BY tester,SendDate;"
        Dim clsMySql As New CMySQL
        Dim dtbHistory As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        dtbHistory.Columns("Test_time").ColumnName = "LockTime"
        GetTesterHistory = dtbHistory
    End Function

    Public Function GetTesterHistoryNewVersion(ByVal strProduct As String, ByVal dtbSearchBy As DataTable, ByVal dtbParam As DataTable, ByVal dtStart As DateTime,
        ByVal dtEnd As DateTime) As DataTable
        Dim strSQL As String
        strSQL = "SELECT A.* , B.CGALot "
        strSQL = strSQL & "FROM db_" & strProduct & ".testerhistory A "
        strSQL = strSQL & "LEFT JOIN db_" & strProduct & ".tabdetail_header B USING(Tester) "
        strSQL = strSQL & "WHERE ("
        Dim strSearchBy As String = dtbSearchBy.TableName
        If strSearchBy.ToLower = "optionindex" Then
            strSearchBy = "tester"
            For nItem As Integer = 0 To dtbSearchBy.Rows.Count - 1
                Dim strMachineType As String = ""
                If dtbSearchBy.Rows(nItem).Item("OptionIndex") Then
                    strMachineType = "Down"
                Else
                    strMachineType = "Up"
                End If
                If nItem <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & " A." & strSearchBy & "='" & strMachineType & "' OR "
                Else
                    strSQL = strSQL & " A." & strSearchBy & "='" & strMachineType & "') "
                End If
            Next nItem
        Else
            For nItem As Integer = 0 To dtbSearchBy.Rows.Count - 1
                If nItem <> dtbSearchBy.Rows.Count - 1 Then
                    strSQL = strSQL & " A." & strSearchBy & "='" & dtbSearchBy.Rows(nItem).Item(strSearchBy) & "' OR "
                Else
                    strSQL = strSQL & " A." & strSearchBy & "='" & dtbSearchBy.Rows(nItem).Item(strSearchBy) & "') "
                End If
            Next nItem
        End If
        strSQL = strSQL & " AND (A.SendDate BETWEEN '" & Format(dtStart, "yyyy-MM-dd HH:mm:ss") & "'  AND '" & Format(dtEnd, "yyyy-MM-dd HH:mm:ss") & "') "
        strSQL = strSQL & " AND (B.Test_time = A.Test_time) "
        strSQL = strSQL & " ORDER BY tester,SendDate;"
        Dim clsMySql As New CMySQL
        Dim dtbHistory As DataTable = clsMySql.CommandMySqlDataTable(strSQL, m_mySqlConn)
        dtbHistory.Columns("Test_time").ColumnName = "LockTime"
        GetTesterHistoryNewVersion = dtbHistory

    End Function

End Class

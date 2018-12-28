Imports MySql.Data.MySqlClient
Imports System.Windows.Forms

Public Class CUserRTTC
    Public Enum enumUserGroup
        eSoftwareEngineer = 0
        eSoftwareTechnician
        eTesterControlEngineer
        eHVMTechnician
        eNPLEngineer
        eProductEngineer
        eJeneralUSER
        eEngineer
    End Enum

    Private m_MyUserConn As MySqlConnection

    Public Sub New(ByVal MySqlConn As MySqlConnection)
        m_MyUserConn = MySqlConn
    End Sub

    Public Function GetUserTable(ByVal strUserName As String) As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM ctr_user_rttc.User_Table "
        strSQL = strSQL & "LEFT JOIN ctr_user_rttc.user_level USING (userlevel) "
        strSQL = strSQL & "WHERE UserAliasName ='" & strUserName & "';"
        Dim cMyUser As New CMySQL
        GetUserTable = cMyUser.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Function GetAllUser() As DataTable
        Dim strSQL As String
        strSQL = "SELECT * FROM ctr_user_rttc.User_Table A "
        strSQL = strSQL & "ORDER BY A.UserAliasName;"
        Dim cAllUser As New CMySQL
        GetAllUser = cAllUser.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Function GetUserAuthorize(ByVal strProduct As String, ByVal nPageID As Integer) As DataTable
        Dim strSQL As String
        strSQL = "SELECT UserAliasName,"
        strSQL = strSQL & "IF(UserAliasName IN (SELECT UserAliasName "
        strSQL = strSQL & "FROM ctr_user_rttc.tabuser_authorize WHERE PageID='" & nPageID & "'),1,0) IsAuthorize "
        strSQL = strSQL & "FROM ctr_user_rttc.User_Table "
        strSQL = strSQL & "WHERE UserLevel<3 "
        strSQL = strSQL & "ORDER BY IsAuthorize DESC,UserAliasName;"
        Dim clsUserLevel As New CMySQL
        GetUserAuthorize = clsUserLevel.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Sub AddAuthorize(ByVal strProduct As String, ByVal nPageID As Integer, ByVal strUserName As String)
        Dim strSQL As String
        strSQL = "INSERT INTO ctr_user_rttc.tabuser_authorize(PageID,UserAliasName) VALUES("
        strSQL = strSQL & "'" & nPageID & "',"
        strSQL = strSQL & "'" & strUserName & "');"

        Dim clsUserLevel As New CMySQL
        clsUserLevel.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Sub

    Public Sub RemoveAuthorize(ByVal strProduct As String, ByVal nPageID As Integer, ByVal strUserName As String)
        Dim strSQL As String
        strSQL = "DELETE FROM ctr_user_rttc.tabuser_authorize WHERE "
        strSQL = strSQL & "PageID='" & nPageID & "' AND "
        strSQL = strSQL & "UserAliasName='" & strUserName & "';"

        Dim clsUserLevel As New CMySQL
        clsUserLevel.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Sub

    Public Function GetUserLevelTable() As DataTable
        Dim strSQL As String = "SELECT DISTINCT A.UserLevel,A.UserLevelName FROM ctr_user_rttc.User_Level A ORDER BY A.UserLevel;"
        Dim cUserLevel As New CMySQL
        GetUserLevelTable = cUserLevel.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Function GetUserGroup() As DataTable
        Dim strSQL As String = "SELECT * FROM ctr_user_rttc.user_group ORDER BY UserGroupName;"
        Dim clsMySql As New CMySQL
        GetUserGroup = clsMySql.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Function GetUserByGroupType(ByVal eUserGroup As enumUserGroup) As DataTable
        Dim strSQL As String = "SELECT * FROM ctr_user_rttc.user_table "
        strSQL = strSQL & "LEFT JOIN ctr_user_rttc.user_group USING(UserGroupID) "
        strSQL = strSQL & "WHERE UserGroupID='" & eUserGroup & "';"
        Dim clsMySql As New CMySQL
        GetUserByGroupType = clsMySql.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function

    Public Function DeleteUser(ByVal strAliasName As String) As Boolean
        DeleteUser = False
        Dim strSQL As String = "DELETE FROM ctr_user_rttc.User_Table  WHERE UserAliasName ='" & strAliasName & "';"
        Dim clsDeleteUser As New CMySQL
        Try
            clsDeleteUser.CommandNonQuery(strSQL, m_MyUserConn, strAliasName, Me.ToString)
            DeleteUser = True
        Catch ex As Exception
            DeleteUser = False
        End Try
    End Function

    Public Sub UpdateUser(ByVal rowUser As DataGridViewRow, ByVal columnSet As DataGridViewColumnCollection)
        Dim strSQL As String = "UPDATE ctr_user_rttc.User_Table SET "
        For nCell As Integer = 1 To 6
            strSQL = strSQL & columnSet(nCell).Name & "='" & rowUser.Cells(nCell).Value & "',"
        Next nCell
        strSQL = strSQL & "UserLevel='" & rowUser.Cells(rowUser.Cells.Count - 1).Value & "' "
        strSQL = strSQL & " WHERE UserAliasName ='" & rowUser.Cells("UserAliasName").Value & "';"
        Dim cUpdateUser As New CMySQL
        cUpdateUser.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Sub

    Public Function InsertUser(ByVal rowUser As DataGridViewRow, ByVal columnSet As DataGridViewColumnCollection) As Boolean
        Try
            Dim strSQL As String = "INSERT INTO ctr_user_rttc.User_Table ("
            For nCell As Integer = 0 To 6
                strSQL = strSQL & columnSet(nCell).Name & ","
            Next nCell
            strSQL = strSQL & "UserLevel) "
            strSQL = strSQL & "SELECT (SELECT IFNULL(Max(UserID)+1,1) FROM ctr_user_rttc.User_Table),"
            For nCell As Integer = 1 To 6
                strSQL = strSQL & "'" & rowUser.Cells(nCell).Value & "',"
            Next nCell
            strSQL = strSQL & "'" & rowUser.Cells(rowUser.Cells.Count - 1).Value & "';"
            Dim cUpdateUser As New CMySQL
            cUpdateUser.CommandNonQuery(strSQL, m_MyUserConn, "ctr_user_rttc.User_Table", Me.ToString)
            InsertUser = True
        Catch ex As Exception
            InsertUser = False
        End Try
    End Function

    Public Function GetUserEmail(ByVal strMailID As String) As String
        Dim strSQL As String
        Dim clsSQL As New CMySQL
        GetUserEmail = ""
        If strMailID <> "" Then
            Dim strSplitID() As String = Split(strMailID, ",")

            strSQL = "SELECT UserEmail FROM ctr_user_rttc.User_Table "
            strSQL = strSQL & " WHERE "
            For nID As Integer = 0 To strSplitID.Length - 2
                strSQL = strSQL & "user_table.userID = " & strSplitID(nID) & " OR "
            Next nID
            strSQL = strSQL & "user_table.userID=" & strSplitID(strSplitID.Length - 1)
            Dim dtbEmail As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
            Dim strMailName As String = ""
            For nRow As Integer = 0 To dtbEmail.Rows.Count - 1
                strMailName = strMailName & dtbEmail.Rows(nRow).Item("UserEmail").ToString & ";"
            Next nRow
            If strMailName <> "" Then
                strMailName = Left(strMailName, Len(strMailName) - 1)
            End If
            GetUserEmail = strMailName
        End If
    End Function

    Public Function GetUserEmailList(ByVal strMailID As String) As DataTable
        Dim strSQL As String
        Dim clsSQL As New CMySQL
        GetUserEmailList = Nothing
        'If strMailID <> "" Then
        Dim strSplitID() As String = Split(strMailID, ",")

        strSQL = "SELECT UserID,UserEmail,"
        strSQL = strSQL & "IF(UserEmail IN (SELECT A.UserEmail "
        strSQL = strSQL & "FROM ctr_user_rttc.User_Table A "
        strSQL = strSQL & " WHERE "
        For nID As Integer = 0 To strSplitID.Length - 2
            strSQL = strSQL & "A.userID = " & strSplitID(nID) & " OR "
        Next nID
        strSQL = strSQL & "A.userID='" & strSplitID(strSplitID.Length - 1) & "'"
        strSQL = strSQL & " ),1,0) SendMailTo"
        strSQL = strSQL & " FROM ctr_user_rttc.User_Table"
        strSQL = strSQL & " ORDER BY SendMailTo DESC,UserEmail;"
        Dim dtbEmail As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
        GetUserEmailList = dtbEmail
        'End If

    End Function

    Public Function GetUserIDByEmail(ByVal chkEmailList As CheckedListBox) As String
        GetUserIDByEmail = ""
        Dim strSQL As String
        Dim bHasSelect As Boolean = False
        strSQL = "SELECT UserID FROM ctr_user_rttc.user_table "
        strSQL = strSQL & " WHERE "
        For nSelect As Integer = 0 To chkEmailList.CheckedItems.Count - 1
            bHasSelect = True
            strSQL = strSQL & " UserEMail ='" & chkEmailList.CheckedItems.Item(nSelect) & "' OR"
        Next nSelect

        If Right(strSQL, 2).ToUpper = "OR" Then strSQL = Left(strSQL, Len(strSQL) - 2)
        strSQL = strSQL & " ORDER BY UserID;"
        If bHasSelect Then
            Dim clsSQL As New CMySQL
            Dim dtbUserID As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
            If dtbUserID.Rows.Count > 0 Then
                For nUser As Integer = 0 To dtbUserID.Rows.Count - 2
                    GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(nUser).Item("UserID").ToString & ","
                Next nUser
                GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(dtbUserID.Rows.Count - 1).Item("UserID").ToString
            End If
        End If
    End Function

    Public Function GetUserIDByEmail(ByVal strEmailAddr() As String) As String
        GetUserIDByEmail = ""
        If strEmailAddr IsNot Nothing Then
            Dim strSQL As String
            Dim bHasSelect As Boolean = False
            strSQL = "SELECT UserID FROM ctr_user_rttc.user_table "
            strSQL = strSQL & " WHERE "
            For nSelect As Integer = 0 To strEmailAddr.Length - 1
                bHasSelect = True
                strSQL = strSQL & " UserEMail ='" & strEmailAddr(nSelect) & "' OR"
            Next nSelect

            If Right(strSQL, 2).ToUpper = "OR" Then strSQL = Left(strSQL, Len(strSQL) - 2)
            strSQL = strSQL & " ORDER BY UserID;"
            If bHasSelect Then
                Dim clsSQL As New CMySQL
                Dim dtbUserID As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
                If dtbUserID.Rows.Count > 0 Then
                    For nUser As Integer = 0 To dtbUserID.Rows.Count - 2
                        GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(nUser).Item("UserID").ToString & ","
                    Next nUser
                    GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(dtbUserID.Rows.Count - 1).Item("UserID").ToString
                End If
            End If
        End If
    End Function


    Public Function GetUserIDByEmail(ByVal dtbEmailName As DataTable) As String
        GetUserIDByEmail = ""
        Dim strSQL As String
        Dim bHasSelect As Boolean = False
        strSQL = "SELECT UserID FROM ctr_user_rttc.user_table "
        strSQL = strSQL & " WHERE "
        For nSelect As Integer = 0 To dtbEmailName.Rows.Count - 1
            bHasSelect = True
            strSQL = strSQL & " UserEMail ='" & dtbEmailName.Rows(nSelect).Item(0) & "' OR"
        Next nSelect

        If Right(strSQL, 2).ToUpper = "OR" Then strSQL = Left(strSQL, Len(strSQL) - 2)
        strSQL = strSQL & " ORDER BY UserID;"
        If bHasSelect Then
            Dim clsSQL As New CMySQL
            Dim dtbUserID As DataTable = clsSQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
            If dtbUserID.Rows.Count > 0 Then
                For nUser As Integer = 0 To dtbUserID.Rows.Count - 2
                    GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(nUser).Item("UserID").ToString & ","
                Next nUser
                GetUserIDByEmail = GetUserIDByEmail & dtbUserID.Rows(dtbUserID.Rows.Count - 1).Item("UserID").ToString
            End If
        End If
    End Function

    Public Sub LoginRTTC(ByVal sUserDetail As SCurrentUser)
        LoginLogOut(sUserDetail, True)
    End Sub

    Public Sub LogOutRTTC(ByVal sUserDetail As SCurrentUser)
        LoginLogOut(sUserDetail, False)
    End Sub

    Private Sub LoginLogOut(ByVal sUserDetail As SCurrentUser, ByVal bLogInOrOut As Boolean)
        Dim strSQL As String = "INSERT INTO ctr_user_rttc.tabcurrentuser("
        strSQL = strSQL & "UserAliasName,"
        strSQL = strSQL & "LoginDate,"
        strSQL = strSQL & "IPAdr,"
        strSQL = strSQL & "OnlineStatus) VALUE("
        strSQL = strSQL & "'" & sUserDetail.strUserName & "',"
        strSQL = strSQL & "'" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "',"
        strSQL = strSQL & "'" & sUserDetail.strIPAdr & "',"
        strSQL = strSQL & bLogInOrOut.ToString & ") "
        strSQL = strSQL & "ON DUPLICATE KEY UPDATE "
        strSQL = strSQL & "LoginDate='" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "',"
        strSQL = strSQL & "IPAdr='" & sUserDetail.strIPAdr & "',"
        strSQL = strSQL & "OnlineStatus=" & bLogInOrOut.ToString & ";"
        Dim clsMySql As New CMySQL
        clsMySql.CommandNoQuery(strSQL, m_MyUserConn)

    End Sub

    Public Function GetUserDetail(ByVal strUserName As String) As SCurrentUser
        GetUserDetail.strUserName = strUserName
        GetUserDetail.strPassword = ""
        GetUserDetail.eUserLevel = enuUserLevel.enuUnAutorize
        GetUserDetail.strLevelText = ""
        GetUserDetail.strIPAdr = ""

        Dim clsUser As New CUserRTTC(m_MyUserConn)
        Dim dtbUser As DataTable = clsUser.GetUserTable(GetUserDetail.strUserName)
        If dtbUser.Rows.Count > 0 Then
            GetUserDetail.strUserName = dtbUser.Rows(0).Item("UserAliasName")
            GetUserDetail.strPassword = dtbUser.Rows(0).Item("UserPasswd")
            GetUserDetail.eUserLevel = dtbUser.Rows(0).Item("UserLevel")
            GetUserDetail.strLevelText = dtbUser.Rows(0).Item("UserLevelName")
            GetUserDetail.strIPAdr = GetIPAddress()
        End If
    End Function

    Public Function GetUserOnline() As DataTable
        Dim strSQL As String = "SELECT * FROM ctr_user_rttc.tabcurrentuser "
        strSQL = strSQL & "WHERE OnlineStatus=TRUE "
        strSQL = strSQL & "ORDER BY UserAliasName;"
        Dim clsMySQL As New CMySQL
        GetUserOnline = clsMySQL.CommandMySqlDataTable(strSQL, m_MyUserConn)
    End Function
End Class

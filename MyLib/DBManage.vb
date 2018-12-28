Option Strict On
Imports MySql.Data.MySqlClient
Imports System.Text

Public Class DBManage
#Region "Fields"

    Private _appName As String
    Private _stampTime As String
    Private _owner As String
    Private _retErr As String  ' Return any error during process 


    Protected _conn As MySqlConnection
    Protected _sql As String    '  SQL Query
    Private _database As String
    Private _tabName As String
    Private _paras() As String
    Private _values() As String
#End Region

#Region "Constructure"
    Public Sub New(ByVal conn As MySqlConnection)
        If conn IsNot Nothing Then _conn = conn
    End Sub
#End Region

#Region "Properties"

    Public WriteOnly Property AppName() As String
        Set(ByVal value As String)
            _appName = value
        End Set
    End Property
    Public WriteOnly Property StampTime() As String
        Set(ByVal value As String)
            _stampTime = value
        End Set
    End Property
    Public WriteOnly Property Values() As String()
        Set(ByVal value As String())
            _values = value
        End Set
    End Property
    Public WriteOnly Property Owner() As String
        Set(ByVal value As String)
            _owner = value
        End Set
    End Property
    Public WriteOnly Property TabName() As String
        Set(ByVal value As String)
            _tabName = value
        End Set
    End Property

    Public WriteOnly Property Parameters() As String()
        Set(ByVal value As String())
            _paras = value
        End Set
    End Property

    Public WriteOnly Property Database() As String
        Set(ByVal value As String)
            _database = value
        End Set
    End Property

    Public ReadOnly Property RetErr() As String
        Get
            Return _retErr
        End Get
    End Property

#End Region

#Region "Methode"
    ''' <summary>
    ''' Function SelectData is used to select data from a database.
    ''' Default IsWHERE = False. True = The WHERE clause is used to extract only those records that fulfill a specified condition. 
    ''' Condition is using for operators("AND"/"OR").The AND and OR operators are used to filter records based on more than one condition:
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="IsWHERE"></param>
    ''' <param name="condition"></param>
    ''' <returns></returns>
    Public Function SelectData(Optional ByRef dt As DataTable = Nothing, Optional ByVal IsWHERE As Boolean = False, Optional ByVal condition As String = Nothing) As Integer
        Dim Ret As Integer
        Try
            _sql = "SELECT * FROM " & _tabName & " A"
            If IsWHERE = False Then

                Ret = excuteQuery(_sql, dt)

            Else
                If Not _values Is Nothing Then
                    If _paras.Length = _values.Length And String.IsNullOrEmpty(condition) = False Then
                        _sql = _sql & " WHERE "
                        For i As Integer = 0 To _paras.Length - 1
                            _sql = _sql & _paras(i) & " = " & "'" & _values(i) & "' " & condition
                        Next
                        If Right(_sql, 4) = "AND " Then _sql = (Left(_sql, Len(_sql) - 4) & " ").Trim
                        If Right(_sql, 3) = "OR " Then _sql = (Left(_sql, Len(_sql) - 3) & " ").Trim
                        If Right(_sql, 1) <> ";" Then _sql = _sql & ";"
                        Ret = excuteQuery(_sql, dt)
                    Else
                        _sql = _sql & " WHERE "
                        _sql = _sql & _paras(0) & " = " & "'" & _values(0) & "'" & condition
                        If Right(_sql, 1) <> ";" Then _sql = _sql & ";"
                        Ret = excuteQuery(_sql, dt)
                    End If
                Else
                    Ret = -1
                    _retErr = "value is nothing"
                End If

            End If

        Catch ex As Exception
            Ret = -1
            _retErr = "CheckData " & ex.Message
        End Try

        Return Ret

    End Function
    ''' <summary>
    ''' Insert function.The INSERT INTO statement is used to insert new records in a table.
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    Public Function Insert(Optional ByRef dt As DataTable = Nothing) As Integer

        Dim Ret As Integer
        Try
            _sql = "INSERT INTO " & _tabName & "("
            Dim para As String = ""
            Dim val As String = ""
            If _paras.Length = _values.Length Then
                'insert multiple parameters 
                For i As Integer = 0 To _paras.Length - 1
                    para = para & "," & _paras(i)
                    If _paras(i).Contains("_id") = False Then
                        val = val & "," & "'" & _values(i) & "'"  ' none integer type needs "'"
                    Else
                        val = val & "," & _values(i)
                    End If
                Next
                '------set string format parameter ------------
                If Left(para, 1) = "," Then para = (Right(para, Len(para) - 1) & " ").Trim
                If Right(para, 1) = "," Then para = (Left(para, Len(para) - 1) & " ").Trim
                If Right(para, 1) <> ")" Then para = para & ")"
                '------set string format values ------------
                If Left(val, 1) = "," Then val = (Right(val, Len(val) - 1) & " ").Trim
                If Right(val, 1) = "," Then val = (Left(para, Len(val) - 1) & " ").Trim
                If Left(val, 1) <> "(" Then val = "(" & val
                If Right(val, 1) <> ")" Then val = val & ");"
                _sql = _sql & para
                _sql = _sql & " VALUES " & val

            Else 'insert one data
                para = _paras(0) & ")"
                If Left(val, 1) <> "(" Then val = val & "("
                val = _values(0) & ")"

                _sql = _sql & para
                _sql = _sql & " VALUES " & val
            End If
            Ret = excuteQuery(_sql, dt)  'Return excuted Query
        Catch ex As Exception
            Ret = -1
            _retErr = "CheckData" & ex.Message
        End Try
        Return Ret

    End Function
    ''' <summary>
    ''' Update database.The UPDATE statement is used to modify the existing records in a table.
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    Public Function Update(Optional ByRef dt As DataTable = Nothing) As Integer
        Dim Ret As Integer
        Try
            _sql = "UPDATE " & _tabName
            _sql = _sql & " SET " & _paras(1) & "= '" & _values(1) & "' ," 'stamp_time_bigint
            _sql = _sql & _paras(2) & "= '" & _values(2) & "' ,"  ' stamp_time
            _sql = _sql & _paras(4) & "= '" & _values(4) & "'" 'STATUS
            _sql = _sql & " WHERE " & _paras(3) & "= '" & _values(3) & "';" 'STATUS
            'Update tabstatus
            'SET stamp_time_bigint = '20170623172329' , stamp_time = '2017-06-23 05:23:29', STATUS = '1'
            'WHERE tag_id = 1
            Ret = excuteQuery(_sql, dt)
        Catch ex As Exception
            Ret = -1
            _retErr = "CheckData" & ex.Message
        End Try

        Return Ret
    End Function
    ''' <summary>
    ''' GetIdbyMultiTables is used to get  server id ,system is and owner id to provide the condition for update tabapplication 
    ''' Parameter >> (0) server_name , (1) system_name , (3) owner_name
    ''' /*****SQL Command******
    '''Select Case a.server_id, b.system_id, c.owner_id
    '''From tabserverinfo a, tabsysteminfo b , tabownerinfo c
    '''Where a.server_name = 'wdtbtss02' AND b.system_name = 'CentralLock' AND c.owner_name = "phakkhapon_w"; */
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    Public Function GetIdbyMultiTables(Optional ByRef dt As DataTable = Nothing) As Integer

        Dim Ret As Integer
        Try
            If _values.Length > 0 Then
                _sql = "SELECT"
                _sql = _sql & " a.server_id, b.system_id, c.owner_id"
                _sql = _sql & " FROM tabserverinfo a, tabsysteminfo b , tabownerinfo c"
                _sql = _sql & " WHERE a.server_name = '" & _values(0).ToString() & "' AND b.system_name = '" & _values(1).ToString() & "'"
                _sql = _sql & " AND c.owner_name = '" & _values(2).ToString() & "'"
                _sql = _sql & " LIMIT 0,100;"
                Ret = excuteQuery(_sql, dt)  'Return excuted Query
            Else
                Ret = -1
                _retErr = "parmaeter _values is not valid "
            End If

        Catch ex As Exception
            Ret = -1
            _retErr = "CheckData" & ex.Message
        End Try
        Return Ret

    End Function

    Public Function excuteQuery(sqlcmd As String, Optional ByRef dt As DataTable = Nothing) As Integer
        Dim Ret As Integer
        Try
            If _conn.State = ConnectionState.Closed Then _conn.Open()
            If Not String.IsNullOrEmpty(sqlcmd) Then
                ' sqlcmd = sqlcmd.ToUpper
                Dim ds As New DataSet
                Dim da As New MySqlDataAdapter(sqlcmd, _conn)
                da.Fill(ds)
                If dt IsNot Nothing Then
                    dt.Reset()
                    dt = ds.Tables(0)
                End If
            Else
                Dim cmd As New MySqlCommand
                With cmd
                    .Connection = _conn
                    .CommandText = sqlcmd
                    Ret = .ExecuteNonQuery()
                End With
            End If
        Catch ex As Exception

            Ret = -1
            _retErr = "excuteQuery" & ex.Message
        End Try
        Return Ret
    End Function

#End Region
End Class



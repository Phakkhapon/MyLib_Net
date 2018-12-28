Option Strict On
Imports MySql.Data.MySqlClient

Public Class ConnString

#Region "Properties"
    Private _server As String
    Public WriteOnly Property Server() As String
        Set(ByVal value As String)
            _server = value
        End Set
    End Property

    Private _user As String
    Public WriteOnly Property User() As String
        Set(ByVal value As String)
            _user = value
        End Set
    End Property

    Private _pw As String
    Public WriteOnly Property Password() As String
        Set(ByVal value As String)
            _pw = value
        End Set
    End Property

    Private _port As String
    Public WriteOnly Property Port() As String
        Set(ByVal value As String)
            _port = value
        End Set
    End Property

    Private _database As String
    Public WriteOnly Property Database() As String
        Set(ByVal value As String)
            _port = value
        End Set
    End Property

    Private _msqlConn As MySqlConnection
    Public Property GetConnectionString() As MySqlConnection
        Get
            SetStringConnection
            Return _msqlConn
        End Get
        Set(ByVal value As MySqlConnection)
            _msqlConn = value
        End Set
    End Property
#End Region

#Region "Methode"
    Private Sub SetStringConnection()
        If _msqlConn Is Nothing Then
            _msqlConn = New MySqlConnection
        End If
        If _msqlConn.State = ConnectionState.Open Then _msqlConn.Close()
        _msqlConn.ConnectionString = GetConnString()
    End Sub

    Private Function GetConnString() As String
        Dim retSQLConn As String
        retSQLConn = "server=" & _server & ";"
        retSQLConn = retSQLConn & "uid=" & _user & ";"
        retSQLConn = retSQLConn & "pwd=" & _pw & ";"
        If _database <> "" Then retSQLConn = retSQLConn & "database=" & _database & ";"
        retSQLConn = retSQLConn & "port=" & _port & ";"
        retSQLConn = retSQLConn & "Connection Lifetime=15;"
        Return retSQLConn
    End Function

#End Region
    Declare Function SetProcessWorkingSetSize Lib "kernel32.dll" (ByVal process As IntPtr, ByVal minimumWorkingSetSize As Integer, ByVal maximumWorkingSetSize As Integer) As Integer
    Public Sub FlushMemory()
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If (Environment.OSVersion.Platform = PlatformID.Win32NT) Then
                SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1)
                Dim myProcesses As Process() = Process.GetProcessesByName("ApplicationName")
                Dim myProcess As Process
                'Dim ProcessInfo As Process
                For Each myProcess In myProcesses
                    SetProcessWorkingSetSize(myProcess.Handle, -1, -1)
                Next myProcess
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub

End Class

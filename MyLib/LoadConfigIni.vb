Option Strict On
Imports System.IO

Public Class LoadConfigIni

#Region "fields"
    Private _server As String
    Private _user As String
    Private _pw As String
    Private _port As String
    Private _database As String
    Private _configIniPath As String
#End Region

#Region "Properties"
    Public Property Server() As String
        Get
            Return _server
        End Get
        Set(ByVal value As String)
            _server = value
        End Set
    End Property
    Public Property User() As String
        Get
            Return _user
        End Get
        Set(ByVal value As String)
            _user = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return _pw
        End Get
        Set(ByVal value As String)
            _pw = value
        End Set
    End Property
    Public Property Port() As String
        Get
            Return _port
        End Get
        Set(ByVal value As String)
            _port = value
        End Set
    End Property
    Public Property Database() As String
        Get
            Return _database
        End Get
        Set(ByVal value As String)
            _database = value
        End Set
    End Property
    Public Property ConfigIniPath() As String
        Get
            Return _configIniPath
        End Get
        Set(ByVal value As String)
            _configIniPath = value
        End Set
    End Property
#End Region

#Region "Methode"
    ''' <summary>
    ''' server , user , pw , port and database
    ''' </summary>
    ''' <param name="configIni"></param>

    Public Overridable Sub InitialConfig(configIni As String)

        Dim ini As IniFile = New IniFile(configIni)
        ini.WriteString("ServerConfig", "server", _server)
        ini.WriteString("ServerConfig", "user", _user)
        ini.WriteString("ServerConfig", "pw", _pw)
        ini.WriteString("ServerConfig", "port", _port)
        ini.WriteString("ServerConfig", "database", _database)

    End Sub

    Public Overridable Sub LoadConfig()
        If File.Exists(_configIniPath) = False Then InitialConfig(_configIniPath)
        Dim cini As New IniFile(_configIniPath)
        _server = cini.GetString("serverConfig", "server", "")
        _user = cini.GetString("serverConfig", "user", "")
        _pw = cini.GetString("serverConfig", "pw", "")
        _port = cini.GetString("serverConfig", "port", "")
        _database = cini.GetString("serverConfig", "database", "")

    End Sub

#End Region

    ''  strSQL = String.Format("sp_cst2 '{0}','{1}','{2}' ;", strTester, strCode, strDetail)


End Class

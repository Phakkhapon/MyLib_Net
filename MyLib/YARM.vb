Option Strict On
Imports System.IO

Public Class YARM
    ''' <summary>
    ''' set properties
    ''' </summary>
#Region "Properties"

    Private _yarmpath As String
    Public WriteOnly Property YarmPath() As String
        Set(ByVal value As String)
            _yarmpath = value
        End Set
    End Property

    Private _yarmuser As String
    Public WriteOnly Property YarmUser() As String
        Set(ByVal value As String)
            _yarmuser = value
        End Set
    End Property

    Private _fileExeName As String
    Public WriteOnly Property FileExeName() As String
        Set(ByVal value As String)
            _fileExeName = value
        End Set
    End Property

    Private _sytemtype As String
    Public WriteOnly Property SystemType() As String
        Set(ByVal value As String)
            _sytemtype = value
        End Set
    End Property

    Private _stmptime As Boolean = False
    Public ReadOnly Property StampYarm() As Boolean
        Get
            _stmptime = StampTime()
            Return _stmptime
        End Get
    End Property

    Private _RetErr As String = ""
    Public ReadOnly Property RetErr() As String
        Get
            Return _RetErr
        End Get
    End Property

#End Region

#Region "Methode"

    Private Function StampTime() As Boolean
        Dim ret As Boolean = True
        Try
            Dim detector As String = ""
            Dim comName As String = System.Environment.MachineName
            Dim stamppath As String = _yarmpath
            stamppath = stamppath & _fileExeName & ".ckt"
            detector = CStr(DateAndTime.Now) & ",300," & comName & "," & _yarmuser
            Dim ini As IniFile = New IniFile(stamppath)
            ini.WriteString(_sytemtype, _fileExeName, detector)
            ret = True
        Catch ex As Exception
            ret = False
            _RetErr = ex.Message

        End Try
        Return ret

    End Function


#End Region

End Class

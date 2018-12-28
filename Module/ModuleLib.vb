
Imports MySql.Data.MySqlClient

Public Module ModuleLib
    Public Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As IntPtr

    Public Enum enuProductType
        enuProductAll = 0
        enuProductFastTrack
        enuProductXLot
        enuProductSDET
        enuProductNPL
        enuProductDOE
        enuProductSTD
    End Enum

    Public Enum enumSearchOption
        eSearchByTester = 0
        eSearchByLot
        eSearchBySpec
        eSearchByMachineType
        eSearchByWafer
        eSearchByCGALot
    End Enum

    Public Structure SMachineDetail
        Dim strMediaSN As String
        Dim strProductName As String
        Dim strMachineName As String
        Dim strMachineType As String
        Dim strIPAdr As String
        Dim strProductFolderPath As String
        Dim strCFPathShoe1 As String
        Dim strCFPathShoe2 As String
        Dim strCFFolderShoe1 As String
        Dim strCFFolderShoe2 As String
    End Structure

    Public Enum enumGradeOption
        eUnloadDefect = 0
        eGradeU
        eGradeV
        eGradeNone
        eGradeAll
    End Enum

    Public Enum enumUserType
        eAdmin = 1
        eSuperUser
        eUser
        eViewer
        eNPLUser
    End Enum

    Public Enum enuUserLevel
        enuUnAutorize = -1
        enuSystemOwner = 0
        enuAdmin = 1
        enuSuperUser = 2
        enuUser = 3
        enuViewer = 4
    End Enum

    Public Enum enumCFType
        eCFAdd = 0
        eCFMul
    End Enum

    Public Enum enumMachineType
        eTypeUp = 0
        eTypeDown
    End Enum

    Public Structure SCurrentUser
        Dim strUserName As String
        Dim strPassword As String
        Dim strLevelText As String
        Dim eUserLevel As enuUserLevel
        Dim strIPAdr As String
    End Structure

    Public g_sCurrentUserDetail As SCurrentUser

    Public Structure SRTTCSoftware
        Dim strFileName As String
        Dim strFullFileName As String
        Dim strVersion As String
    End Structure


    Public Function GetUserName() As String
        If TypeOf My.User.CurrentPrincipal Is  _
            Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            Dim username As String = parts(1)
            Return username
        Else
            ' The application is using custom authentication.
            Return My.User.Name
        End If
    End Function

    Public Function GetDomainName() As String
        If TypeOf My.User.CurrentPrincipal Is  _
               Security.Principal.WindowsPrincipal Then
            ' The application is using Windows authentication.
            ' The name format is DOMAIN\USERNAME.
            Dim parts() As String = Split(My.User.Name, "\")
            Dim strDomain As String = parts(0)
            Return strDomain
        Else
            ' The application is using custom authentication.
            Return ""
        End If
    End Function

    Public Function GetIPAddress() As String

        GetIPAddress = String.Empty
        Dim strHostName As String = System.Net.Dns.GetHostName()
        Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                GetIPAddress = ipheal.ToString()
            End If
        Next

        'Dim strHostName As String
        'Dim strIPAddress As String = ""

        'strHostName = Environment.MachineName 'System.Net.Dns.GetHostName()
        'strHostName = System.Net.Dns.GetHostName()

        'Dim IPAdr() As System.Net.IPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList
        ''   Dim IP As System.Net.IPAddress = System.Net.Dns.GetHostAddresses


        'For Each IP As System.Net.IPAddress In IPAdr
        '    If Not IP.IsIPv6LinkLocal Then
        '        strIPAddress = IP.ToString
        '    End If
        'Next
        ''strIPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList(1).ToString()

        'Return strIPAddress
    End Function

    Public Function GetMachineName() As String
        GetMachineName = Environment.MachineName
    End Function
 
End Module

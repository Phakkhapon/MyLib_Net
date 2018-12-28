Public Class CStructure
    Public Structure SCurrentUser
        Dim strUserName As String
        Dim strPassword As String
        Dim strLevelText As String
        Dim eUserLevel As enuUserLevel
        Dim strIPAdr As String
    End Structure
    Public CurrentUserDetail As SCurrentUser
    Public Enum enuUserLevel
        enuUnAutorize = -1
        enuSystemOwner = 0
        enuAdmin = 1
        enuSuperUser = 2
        enuUser = 3
        enuViewer = 4
    End Enum

End Class

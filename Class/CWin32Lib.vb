
Public Class CWin32Lib
    Private Const BM_CLICK = 245
    Private Const WM_DESTROY = 2
    Private Declare Ansi Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal strClass As String, ByVal strWindow As String) As IntPtr
    Private Declare Ansi Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hwndParent As IntPtr, ByVal hwndChildAfter As IntPtr, ByVal strClass As String, ByVal strWindow As String) As IntPtr
    Private Declare Ansi Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal UintMessage As Integer, ByVal wParam As Integer, ByVal LParam As UInt64) As IntPtr
    Private Declare Ansi Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As IntPtr, ByVal UintMessage As Integer, ByVal wParam As Integer, ByVal LParam As UInt64) As IntPtr

    'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr
    'Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr


    Public Function GetHandleWindow(ByVal strClassName As String, ByVal strWindowName As String) As IntPtr
        Return FindWindow(strClassName, strWindowName)
    End Function

    Public Function GetChildHandleWindow(ByVal hwndParent As IntPtr, ByVal hwndChildAfter As IntPtr, ByVal strClass As String, ByVal strWindow As String) As IntPtr
        Return FindWindowEx(hwndParent, hwndChildAfter, strClass, strWindow)
    End Function

    Public Function SendWindowMessage(ByVal hwndTarget As IntPtr, ByVal UintMessage As UInt16, ByVal wParam As UInt64, ByVal LParam As UInt64) As IntPtr
        Return SendMessage(hwndTarget, UintMessage, wParam, LParam)
    End Function
    Public Function SendClick(ByVal hwndTarget As IntPtr) As IntPtr
        Return SendMessage(hwndTarget, BM_CLICK, 0, 0)
    End Function


End Class


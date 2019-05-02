Imports System.Runtime.InteropServices
Imports System.Text

Module CommonLibraries
    Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As IntPtr, ByVal ncode As Integer, ByVal wParam As IntPtr, lParam As IntPtr) As IntPtr
    Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal idHook As Integer) As Boolean
    Public Declare Function GetForegroundWindow Lib "user32" () As IntPtr
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Integer) As Short
    Public Declare Function GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As IntPtr) As Integer
    Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    Public Declare Function EmptyClipboard Lib "user32" () As Boolean
    Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As IntPtr) As Boolean
    Public Declare Function SetClipboardViewer Lib "user32" (ByVal hWndNewViewer As IntPtr) As IntPtr
    Public Declare Function ChangeClipboardChain Lib "user32" (ByVal hWndRemove As IntPtr, ByVal hWndNewNext As IntPtr) As Boolean
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
End Module


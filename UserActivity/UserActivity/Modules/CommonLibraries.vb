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
End Module

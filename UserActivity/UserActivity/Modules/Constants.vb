﻿Module Constants
    'Usuario'
    Public user As String = Environment.UserName
    'Se refiere a la ruta correspondiente del escritorio, task manager... etc'
    Public explorer As String = "C:\WINDOWS\Explorer.EXE"
    'Notificaciones de entrada de ratón'
    Public Const WM_LBUTTONDOWN As Integer = &H201
    Public Const WM_LBUTTONUP As Integer = &H202
    Public Const WM_RBUTTONDOWN As Integer = &H204
    Public Const WM_MOUSEWHEEL As Integer = &H20A
    Public Const VK_LCONTROL As Integer = &HA2
    Public Const VK_RCONTROL As Integer = &HA3
    Public Const VK_LMENU As Integer = &HA4
    Public Const VK_TAB As Integer = &H9
    Public Const VK_C As Integer = &H43
    Public Const VK_X As Integer = &H58
    Public Const VK_V As Integer = &H56
    Public Const VK_F As Integer = &H46
    Public Const VK_Z As Integer = &H5A
    Public Const VK_Y As Integer = &H59
    Public Const VK_S As Integer = &H53
    Public Const VK_G As Integer = &H47
    Public Const VK_F4 As Integer = &H73
    Public Const WM_MENURBUTTONUP As Integer = &H122
    Public Const WM_INITMENUPOPUP As Integer = &H117
    Public Const WM_CONTEXTMENU As Integer = &H7B
    'Notificaciones de entrada de teclado'
    Public Const WM_KEYDOWN As Integer = &H100
    Public Const WM_SYSKEYDOWN As Integer = &H104
    'Notificaciones de Clipboard'
    Public Const WM_DRAWCLIPBOARD As Integer = 776
    Public Const WM_CHANGECBCHAIN As Integer = 781
    Public Const WM_PASTE As Integer = &H302
    Public Const WM_COPYDATA As Integer = 74
End Module

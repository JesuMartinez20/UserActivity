﻿Module Constants
    'Ficheros'
    Public pathIni As String = Application.StartupPath + "\config.ini"
    Public pathCatalogActions As String
    Public pathLogActions As String
    'Usuario'
    Public userName As String = Environment.UserName
    'Se refiere a la ruta correspondiente del escritorio, task manager... etc'
    Public explorer As String = "explorer"
    Public userActivity As String = "useractivitycompany"
    'Notificaciones de entrada de ratón'
    Public Const WM_MOUSEWHEEL As Integer = &H20A
    'Notificaciones de entrada de teclado'
    Public Const WM_KEYDOWN As Integer = &H100
    Public Const WM_SYSKEYDOWN As Integer = &H104
    Public Const VK_LCONTROL As Integer = &HA2
    Public Const VK_RCONTROL As Integer = &HA3
    Public Const VK_LMENU As Integer = &HA4
    Public Const VK_LSHIFT As Integer = &HA0
    Public Const VK_RSHIFT As Integer = &HA1
    Public Const VK_INSERT As Integer = &H2D
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
    'Notificaciones de Clipboard'
    Public Const WM_DRAWCLIPBOARD As Integer = 776
    Public Const WM_CHANGECBCHAIN As Integer = 781
End Module

﻿Imports System.Text

Module CommonMethods
    'Obtiene el path completo de una aplicación determinada'
    Public Function GetPathName()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim proc As Process
        Dim wProcID As Integer = Nothing
        Dim wFileName As String = ""
        '
        If hWnd <> IntPtr.Zero Then
            GetWindowThreadProcessId(hWnd, wProcID)
            proc = Process.GetProcessById(wProcID)
            'capturamos los procesos por si alguno no tiene permisos de lectura'
            Try
                wFileName = proc.MainModule.FileName
            Catch ex As Exception
                wFileName = ""
            End Try
        End If
        Return wFileName
    End Function
    'Obtiene el número de un proceso determinado'
    Public Function GetProcessID()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim procID As Integer
        Dim wProcID As Integer = Nothing
        procID = GetWindowThreadProcessId(hWnd, wProcID)
        Return procID
    End Function
    'Método que devuelve el foco y por lo tanto el titulo de la aplicación'
    Public Function GetForegroundInfo()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim length As Integer
        Dim wTitle As StringBuilder = New StringBuilder("", 0)
        '
        If hWnd <> IntPtr.Zero Then
            length = GetWindowTextLength(hWnd)
            wTitle = New StringBuilder("", length + 1)
            If length > 0 Then
                GetWindowText(hWnd, wTitle, wTitle.Capacity)
            End If
        End If
        Return wTitle.ToString
    End Function
    'Método que devuelve el valor de una key específica de un diccionario'
    Public Function SearchValue(ByVal dictionary As Dictionary(Of String, Integer), ByVal key As String)
        Dim action As Integer
        If dictionary.ContainsKey(key) Then
            action = dictionary.Where(Function(p) p.Key = key).FirstOrDefault.Value
        Else action = 0
        End If
        Return action
    End Function

    'Método que devuelve la key de un valor específico de un diccionario'
    Public Function SearchKey(ByVal dictionary As Dictionary(Of String, Integer), ByVal value As Integer)
        Dim actionName As String
        If dictionary.ContainsValue(value) Then
            actionName = dictionary.Where(Function(p) p.Value = value).FirstOrDefault.Key
        Else actionName = 0
        End If
        Return actionName
    End Function
End Module
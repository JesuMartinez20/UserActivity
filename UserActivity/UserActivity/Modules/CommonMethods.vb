Imports System.Text

Module CommonMethods
    'Obtiene el path completo de una aplicación determinada'
    Public Function GetPathName()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim proc As Process
        Dim wProcID As Integer = Nothing
        Dim wFileName As String = ""

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
        Dim wTitle As StringBuilder = New System.Text.StringBuilder("", 0)

        If hWnd <> IntPtr.Zero Then
            length = GetWindowTextLength(hWnd)
            wTitle = New System.Text.StringBuilder("", length + 1)
            If length > 0 Then
                GetWindowText(hWnd, wTitle, wTitle.Capacity)
            End If
        End If
        Return wTitle.ToString
    End Function
End Module

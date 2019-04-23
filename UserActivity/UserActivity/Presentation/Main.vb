Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents wHook As WindowsHook
    'Private thread As Threading.Thread
    'Private Delegate Sub AddItemCallBack(ByVal item As String)
    Private focusfinal As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        focusfinal = GetPathName()
        'Static focusMem As String
        'thread = New Threading.Thread(AddressOf window)
        'thread.Start()
        kbHook = New KeyboardHook(focusfinal)
        mHook = New MouseHook(focusfinal)
        If mHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set mouse hook")
            mHook.Dispose()
        ElseIf kbHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
            kbHook.Dispose()
        End If
        'wHook = New WindowsHook(Me)
    End Sub

    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
        'kbHook = New KeyboardHook()
    End Sub

    Private Sub kbHook_KeyDown(ByVal pathTitle As String, ByVal processID As Integer) Handles kbHook.KeyDown
        ListBox1.Items.Add(Now.ToString + "#" + "escribiendo en App: " + pathTitle + "#" + "ID Proceso: " + processID.ToString + "#" + "Usuario: " + user)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    Private Sub kbHook_CombKey(ByVal Key As Keys, ByVal vKey As Keys, ByVal wTitle As String, ByVal processID As Integer) Handles kbHook.CombKey
        If focusfinal <> wTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + "activa App[" + Key.ToString + "+" + vKey.ToString + "]" + ": " + wTitle + "#" + "ID Proceso: " + processID.ToString + "#" + "Usuario: " + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            focusfinal = wTitle
        Else
        End If
    End Sub

    Private Sub mHook_MouseLeftDown(ByVal pathTitle As String, ByVal processID As Integer) Handles mHook.MouseLeftDown
        If focusfinal <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + "activa App: " + pathTitle + "#" + "ID Proceso: " + processID.ToString + "#" + "Usuario: " + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            focusfinal = pathTitle
        Else
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal pathTitle As String, ByVal processID As Integer) Handles mHook.MouseWheel
        ListBox1.Items.Add(Now.ToString + "#" + "scroll en App: " + pathTitle + "#" + "ID Proceso: " + processID.ToString + "#" + "Usuario" + user)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    Private Sub wHook_WindowsOpened(ByVal processName As String, ByVal processID As Integer) Handles wHook.WindowsOpened
        If processName.Equals("") Then
            'Si el proceso aún no se ha registrado completamente en la lista de procesos del sistema no se imprime
        Else
            ListBox1.Items.Add("El usuario: " + user + " ha abierto la aplicación: " + processName + " cuyo nombre de proceso es: " + processID.ToString)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
        End If
    End Sub

    Private Sub wHook_WindowsClosed(ByVal processName As String, ByVal processID As Integer) Handles wHook.WindowsClosed
        ListBox1.Items.Add("El usuario: " + user + " ha cerrado la aplicación: " + processName + " cuyo nombre de proceso es: " + processID.ToString)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Or mHook IsNot Nothing Then
            kbHook.Dispose()
            mHook.Dispose()
        End If
        'thread.Abort()
        'Try
        'wHook.StopAll()
        'Catch ex As Exception
        'MessageBox.Show(ex.Message)
        'End Try
    End Sub

    'Public Sub window()
    'While True
    '       focusfinal = GetPathName()
    '      System.Threading.Thread.Sleep(1000)
    'End While
    'End Sub

    'Path Completo'
    Private Function GetPathName()
        Dim hWnd As IntPtr = GetForegroundWindow()
        Dim proc As Process
        Dim wProcID As Integer = Nothing
        Dim wFileName As String = ""

        If hWnd <> IntPtr.Zero Then
            GetWindowThreadProcessId(hWnd, wProcID)
            proc = Process.GetProcessById(wProcID)
            'por si alguno no tiene permisos de lectura'
            Try
                wFileName = proc.MainModule.FileName
            Catch ex As Exception
                wFileName = ""
            End Try
        End If
        Return wFileName
    End Function

    'Public Sub AddItemToList(ByVal item As String)
    'If Me.ListBox1.InvokeRequired Then
    'Me.Invoke(New AddItemCallBack(AddressOf AddItemToList), item)
    'Else
    '       ListBox1.Items.Add(item)
    'End If
    'End Sub
End Class

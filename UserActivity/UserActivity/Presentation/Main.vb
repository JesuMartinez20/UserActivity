Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private thread As Threading.Thread
    Private lastfocus As String
    Private Delegate Sub AddItemCallBack(ByVal item As String)
    Private focusfinal As String = ""
    Dim IntNextClip As IntPtr
    <DllImport("User32.dll")>
    Protected Shared Function SetClipboardViewer(ByVal hWndNewViewer As IntPtr) As IntPtr
    End Function

    <DllImport("User32.dll")>
    Public Shared Function ChangeClipboardChain(ByVal hWndRemove As IntPtr, ByVal hWndNewNext As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'IntNextClip = SetClipboardViewer(Me.Handle)
        'thread = New Threading.Thread(AddressOf window)
        'thread.Start()
        kbHook = New KeyboardHook(focusfinal)
        mHook = New MouseHook(focusfinal)
        'If mHook.HHookID = IntPtr.Zero Then
        'Throw New Exception("Could not set mouse hook")
        'mHook.Dispose()
        'ElseIf kbHook.HHookID = IntPtr.Zero Then
        'Throw New Exception("Could not set keyboard hook")
        'kbHook.Dispose()
        'End If
    End Sub

    'Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
    'Const WM_DRAWCLIPBOARD As Integer = 776
    'Const WM_CHANGECBCHAIN As Integer = 781
    'Select Case m.Msg
    'Case WM_DRAWCLIPBOARD
    '           DisplayClipboardData()
    '          SendMessage(IntNextClip, m.Msg, m.WParam, m.LParam)
    ' break
    'Case WM_CHANGECBCHAIN
    'If m.WParam = IntNextClip Then
    '               IntNextClip = m.LParam
    'Else
    '               SendMessage(IntNextClip, m.Msg, m.WParam, m.LParam)
    'End If
    ' break
    'Case Else
    'MyBase.WndProc(m)
    ' break
    'End Select
    'End Sub

    'Sub DisplayClipboardData()
    'Try
    'Dim iData As New DataObject
    '       iData = Clipboard.GetDataObject

    'If iData.GetDataPresent(DataFormats.Rtf) Then
    '           MsgBox(iData.GetData(DataFormats.Text, True).ToString())
    'ElseIf iData.GetDataPresent(DataFormats.Text) Then
    '           MsgBox(iData.GetData(DataFormats.Text, True).ToString())
    'Else
    '           MsgBox("Other data format")
    'End If

    'Catch ex As Exception

    'End Try
    'End Sub


    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
        MsgBox(My.Computer.Clipboard.GetText())
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

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Or mHook IsNot Nothing Then
            kbHook.Dispose()
            mHook.Dispose()
        End If
        'thread.Abort()
    End Sub

    Public Sub window()
        While True
            focusfinal = GetPathName()
            If focusfinal <> lastfocus Then
                AddItemToList(Now.ToString + "#" + "activa App: " + focusfinal + "#" + "Usuario: " + user)
                lastfocus = focusfinal
            Else

            End If
            System.Threading.Thread.Sleep(2000)
        End While
    End Sub

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

    Public Sub AddItemToList(ByVal item As String)
        If Me.ListBox1.InvokeRequired Then
            Me.Invoke(New AddItemCallBack(AddressOf AddItemToList), item)
        Else
            ListBox1.Items.Add(item)
        End If
    End Sub
End Class

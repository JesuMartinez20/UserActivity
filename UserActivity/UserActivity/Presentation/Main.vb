Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    'Private threadFocus As Threading.Thread
    'Esta varibale es la encargada de controlar que el foco sea el mismo y no se repitan mismas acciones
    Private lastFocus As String
    Private Delegate Sub AddItemCallBack(ByVal item As String)
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
        'inicializamos'
        lastFocus = GetPathName()
        'threadFocus = New Threading.Thread(AddressOf GetFocusInfo)
        'threadFocus.Start()
        'IntNextClip = SetClipboardViewer(Me.Handle)
        kbHook = New KeyboardHook()
        mHook = New MouseHook()
        If mHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set mouse hook")
            mHook.Dispose()
        ElseIf kbHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
            kbHook.Dispose()
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_DRAWCLIPBOARD As Integer = 776
        Const WM_CHANGECBCHAIN As Integer = 781
        Select Case m.Msg
            Case WM_DRAWCLIPBOARD
                DisplayClipboardData()
                SendMessage(IntNextClip, m.Msg, m.WParam, m.LParam)
            Case WM_CHANGECBCHAIN
                If m.WParam = IntNextClip Then
                    IntNextClip = m.LParam
                Else
                    SendMessage(IntNextClip, m.Msg, m.WParam, m.LParam)
                End If
            Case Else
                MyBase.WndProc(m)
                ' break
        End Select
    End Sub

    Sub DisplayClipboardData()
        Try
            Dim iData As New DataObject
            iData = Clipboard.GetDataObject

            If iData.GetDataPresent(DataFormats.Rtf) Then
                Debug.WriteLine(iData.GetData(DataFormats.Text, True).ToString())
            ElseIf iData.GetDataPresent(DataFormats.Text) Then
                Debug.WriteLine(iData.GetData(DataFormats.Text, True).ToString())
            Else
                MsgBox("Other data format")
            End If

        Catch ex As Exception
            MsgBox("Fallo")

        End Try
    End Sub

    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
    End Sub

    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastFocus = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    'Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Integer, ByVal pathTitle As String) Handles kbHook.CombKey
    '   AddItemToList(Now.ToString + "#" + typeAction.ToString + "[" + key.ToString + "+" + vKey.ToString + "]" + " en App: " + pathTitle + "#" + user)
    '  lastFocus = pathTitle
    'End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " en App: " + pathTitle + "#" + user)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Or mHook IsNot Nothing Then
            kbHook.Dispose()
            mHook.Dispose()
        End If
        'threadFocus.Abort()
    End Sub

    'Public Sub GetFocusInfo()
    'While True
    '       finalfocus = GetPathName()
    'If finalfocus <> lastfocus Then
    'Dim process As Integer = GetProcessID()
    '           AddItemToList(Now.ToString + "#" + "activa App: " + finalfocus + "#" + "ID Proceso: " + process.ToString + "#" + "Usuario: " + user)
    '          lastfocus = finalfocus
    'Else
    'do nothing'
    'End If
    '       System.Threading.Thread.Sleep(2000)
    'End While
    'End Sub
    Public Sub AddItemToList(ByVal item As String)
        If Me.ListBox1.InvokeRequired Then
            Me.Invoke(New AddItemCallBack(AddressOf AddItemToList), item)
        Else
            ListBox1.Items.Add(item)
        End If
    End Sub
End Class

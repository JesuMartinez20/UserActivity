Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    'Esta varibale es la encargada de controlar que el foco sea el mismo y no se repitan mismas acciones
    Private lastFocus As String
    'Delegado que se encarga de llamar al método de manera asíncrona'
    Private Delegate Sub AddItemCallBack(ByVal item As String)
    Private nextClipViewer As IntPtr
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
        StartHooks()
        StartClipboard()
    End Sub
    'Se inicializan los hooks'
    Private Sub StartHooks()
        'Se inicializa el foco principal de la aplicación'
        lastFocus = GetPathName()
        kbHook = New KeyboardHook()
        mHook = New MouseHook()
        fHook = New FocusHook()
        If mHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set mouse hook")
            mHook.Dispose()
        ElseIf kbHook.HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
            kbHook.Dispose()
        End If
    End Sub
    'Se inicializa el clipboard'
    Private Sub StartClipboard()
        Clipboard.Clear()
        nextClipViewer = SetClipboardViewer(Me.Handle)
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_DRAWCLIPBOARD
                GetClipboard()
                SendMessage(nextClipViewer, m.Msg, m.WParam, m.LParam)
            Case WM_CHANGECBCHAIN
                If m.WParam = nextClipViewer Then 'wParam = hWndRemove = hWnd2
                    nextClipViewer = m.LParam 'lParam = hWnd3
                Else
                    SendMessage(nextClipViewer, m.Msg, m.WParam, m.LParam)
                End If
            Case Else
                MyBase.WndProc(m)
                ' break
        End Select
    End Sub

    Private Sub GetClipboard()
        Try
            Dim iData As New DataObject
            iData = Clipboard.GetDataObject
            If iData.GetDataPresent(DataFormats.UnicodeText) Then
                MsgBox(iData.GetData(DataFormats.Text, True).ToString())
            ElseIf iData.GetDataPresent(DataFormats.Text) Then
                MsgBox(iData.GetData(DataFormats.Text, True).ToString())
            Else
                'do nothing'
            End If
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
    End Sub

    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        'De esta manera no interfiere con en el resto de eventos'
        Static focusKey As String
        If focusKey <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            focusKey = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    'Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Integer, ByVal pathTitle As String) Handles kbHook.CombKey
    '   AddItemToList(Now.ToString + "#" + typeAction.ToString + "[" + key.ToString + "+" + vKey.ToString + "]" + " en App: " + pathTitle + "#" + user)
    '  lastFocus = pathTitle
    'End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        Static focusWheel As String
        If focusWheel <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            focusWheel = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_FocusRise(ByVal typeAction As Integer, ByVal pathTitle As String) Handles fHook.FocusRise
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle Then
            AddItemToList(Now.ToString + "#" + typeAction.ToString + " en App: " + pathTitle + "#" + user)
            lastFocus = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Or mHook IsNot Nothing Then
            kbHook.Dispose()
            mHook.Dispose()
        End If
        fHook.FocusThread.Abort()
    End Sub

    'Método que se encarga de capturar el control del hilo para poder modificar la lista de otro proceso'
    Public Sub AddItemToList(ByVal item As String)
        If Me.ListBox1.InvokeRequired Then
            Me.Invoke(New AddItemCallBack(AddressOf AddItemToList), item)
        Else
            ListBox1.Items.Add(item)
        End If
    End Sub
End Class

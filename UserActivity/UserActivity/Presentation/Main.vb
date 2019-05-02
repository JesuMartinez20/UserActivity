Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    'Esta variable es la encargada de controlar que el foco sea el mismo y no se repitan mismas acciones'
    Private lastFocus As String
    'Dictionary<String,Integer>'
    Private dictionary As New Dictionary(Of String, Integer)
    'Delegado que se encarga de llamar al método de manera asíncrona'
    Private Delegate Sub AddItemCallBack(ByVal item As String)
    'Obtiene el Handle de la ventana actual para el clipboard'
    Private nextClipViewer As IntPtr
    'Evento del Clipboard'
    Public Event ClipboardData(ByVal clipboardText As String)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StartHooks()
        StartClipboard()
    End Sub

    'Se inicializan los hooks'
    Private Sub StartHooks()
        'Se inicializa el foco principal de la aplicación'
        lastFocus = GetPathName()
        dictionary = ReadIni()
        MsgBox(dictionary.Where(Function(p) p.Key = "CounterFocus").FirstOrDefault.Value.ToString)
        kbHook = New KeyboardHook()
        mHook = New MouseHook()
        fHook = New FocusHook(dictionary)
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
    'Función que devuelve un diccionario con el contenido del fichero .ini'
    Private Function ReadIni()
        Dim ini As New FicherosINI("C:\Users\jmmanrique\Desktop\config.ini")
        Dim arrayEvents() As String = ini.GetSection("TIPOS_EVENTOS")
        'Como se guardan pares de valores {llave = valor} el Step es igual a 2'
        For i = 0 To arrayEvents.Length - 1 Step 2
            dictionary.Add(arrayEvents(i), Convert.ToInt32(arrayEvents(i + 1)))
        Next
        'Se agrega el contador del cambio de foco'
        dictionary.Add("CounterFocus", ini.GetInteger("FOCO", "CounterFocus"))
        Return dictionary
    End Function

    'Sobrecargamos el Window Procedure para recibir mensajes del Clipboard'
    Protected Overloads Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Select Case m.Msg
            Case WM_DRAWCLIPBOARD 'process Clipboard'
                GetClipboardData()
                SendMessage(nextClipViewer, m.Msg, m.WParam, m.LParam)
            Case WM_CHANGECBCHAIN 'remove viewer'
                If m.WParam = nextClipViewer Then 'wParam = hWndRemove = hWnd2
                    nextClipViewer = m.LParam 'lParam = hWnd3
                Else
                    SendMessage(nextClipViewer, m.Msg, m.WParam, m.LParam)
                End If
            Case Else
                'unhandled window message'
                MyBase.WndProc(m)
        End Select
    End Sub

    'En este método se captura la información que hay contenida en el Clipboard'
    Private Sub GetClipboardData()
        Try
            Dim iData As New DataObject
            iData = Clipboard.GetDataObject
            If iData.ContainsText Then 'ANSI TEXT'
                RaiseEvent ClipboardData(iData.GetData(DataFormats.Text, True).ToString())
            ElseIf iData.ContainsImage Then 'IMAGE FORMAT'
                RaiseEvent ClipboardData(iData.GetData(DataFormats.Bitmap, True).ToString())
            Else
                'Do nothing'
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
    End Sub

    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        'De esta manera no interfiere con en el resto de eventos'
        Static focusKey As String
        If focusKey <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            focusKey = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    'Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Keys) Handles kbHook.CombKey
    '   ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
    'End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        Static focusWheel As String
        If focusWheel <> pathTitle Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
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

    Private Sub ClipboardEvent(ByVal clipboardText As String) Handles Me.ClipboardData
        Static lastcbtext As String
        If clipboardText <> lastcbtext Then
            ListBox1.Items.Add(Now.ToString + "#" + Convert.ToInt32(TypeAction.CopyApp).ToString + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastcbtext = clipboardText
        Else
            'do nothing'
        End If
    End Sub

    Private Sub UnregisterClipboardViewer()
        ChangeClipboardChain(Me.Handle, nextClipViewer)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Or mHook IsNot Nothing Then
            kbHook.Dispose()
            mHook.Dispose()
        End If
        Try
            fHook.FocusThread.Abort()
        Catch ThreadAbortException As Exception
            Debug.WriteLine(ThreadAbortException.Message)
        End Try
        UnregisterClipboardViewer()
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

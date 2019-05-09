Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    Private Const WM_SETTEXT = &HC
    'Esta variable es la encargada de controlar que el foco sea el mismo y no se repitan mismas acciones'
    Private lastFocus As String
    'Esta variable se encarga de controlar la última acción registrada'
    Private lastAction As Integer
    'Dictionary<String,Integer> del fichero .ini'
    Private dictionaryIni As New Dictionary(Of String, Integer)
    'Variable para controlar el último origen del copy'
    Private lastOrigin As String
    'Dictionary<String,Integer> para controlar el número de focos que se van registrando en la aplicación'
    Private dictionaryFocus As New Dictionary(Of String, Integer)
    'Delegado que se encarga de llamar al método de manera asíncrona'
    Private Delegate Sub AddItemCallBack(ByVal item As String)
    'Obtiene el Handle de la ventana actual para el clipboard'
    Private nextClipViewer As IntPtr
    'Evento del Clipboard'
    Public Event ClipboardData(ByVal clipboardText As String)
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(
     ByVal lpClassName As String,
     ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr,
                      ByVal childAfter As IntPtr,
                      ByVal lclassName As String,
                      ByVal windowTitle As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dictionaryIni = ReadIni()
        StartHooks()
        StartClipboard()
    End Sub
    'Se inicializan los hooks'
    Private Sub StartHooks()
        'Se inicializa el foco principal de la aplicación'
        lastFocus = GetPathName()
        'diccionario vacio significa que el archivo .ini no se ha encontrado'
        If dictionaryIni.Count = 0 Then
            Application.Exit()
        Else
            kbHook = New KeyboardHook(dictionaryIni)
            mHook = New MouseHook(dictionaryIni)
            fHook = New FocusHook(dictionaryIni)
            CheckHooks()
        End If
    End Sub
    'Commprueba que se hayan instalado correctamente los hooks'
    Private Sub CheckHooks()
        Try
            If mHook.HHookID = IntPtr.Zero Then
                Throw New Exception("Could not set mouse hook")
                mHook.Dispose()
            ElseIf kbHook.HHookID = IntPtr.Zero Then
                Throw New Exception("Could not set keyboard hook")
                kbHook.Dispose()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Application.Exit()
        End Try
    End Sub
    'Se inicializa el clipboard'
    Private Sub StartClipboard()
        Clipboard.Clear()
        nextClipViewer = SetClipboardViewer(Me.Handle)
    End Sub
    'Función que devuelve un diccionario con el contenido del fichero .ini'
    Private Function ReadIni()
        Dim path = Application.StartupPath + "\config.ini"
        CheckAndLoadFile(path)
        Return dictionaryIni
    End Function
    'Este método se encarga de comprobar si existe el archivo y de crear un diccionario del archivo .ini'
    Private Sub CheckAndLoadFile(path As String)
        Try
            If File.Exists(path) Then
                Dim ini As New FicherosINI(path)
                Dim arrayEvents() As String = ini.GetSection("TIPOS_EVENTOS")
                'Como se guardan pares de valores {llave = valor} el Step es igual a 2'
                For i = 0 To arrayEvents.Length - 1 Step 2
                    dictionaryIni.Add(arrayEvents(i), Convert.ToInt32(arrayEvents(i + 1)))
                Next
                'Se agrega el contador del cambio de foco'
                dictionaryIni.Add("CounterFocus", ini.GetInteger("FOCO", "CounterFocus"))
            Else
                Throw New Exception("No se puede abrir el archivo. Compruebe que la ruta del archivo .ini es válida.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Environment.Exit(0)
        End Try
    End Sub
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

    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        Static focusKey As String
        If pathTitle <> focusKey And pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            focusKey = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal pathTitle As String) Handles kbHook.CombKey
        Static lastkey As Keys
        If vKey <> lastkey And pathTitle <> explorer Then
            'ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " [" + key.ToString + "+" + vKey.ToString + "] en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            lastkey = vKey
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        Static focusWheel As String
        If pathTitle <> focusWheel And pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            focusWheel = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_FocusRise(ByVal typeAction As Integer, ByVal pathTitle As String) Handles fHook.FocusRise
        'Estas dos variables se encargan de contar los focos registrados'
        Static counterFocusApp As Integer = typeAction
        Static counterLastFocus As Integer = typeAction
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle And pathTitle <> explorer Then
            'Si ya se ha registrado un foco determinado se busca en el diccionario de focos y se actualiza el foco'
            If dictionaryFocus.ContainsKey(pathTitle) Then
                Dim focusRegistered As Integer = dictionaryFocus.Where(Function(p) p.Key = pathTitle).FirstOrDefault.Value
                'AddItemToList(Now.ToString + "#" + focusRegistered.ToString + " en App: " + pathTitle + "#" + user)
                AddItemToList(Now.ToString + "#" + focusRegistered.ToString + "#" + user)
                lastFocus = pathTitle
                lastAction = focusRegistered
                'Se inicializa el foco con el número correspondiente del archivo .ini'
            ElseIf counterFocusApp = counterLastFocus Then
                AddItemToList(Now.ToString + "#" + counterFocusApp.ToString + "#" + user)
                'AddItemToList(Now.ToString + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                dictionaryFocus.Add(pathTitle, counterFocusApp)
                counterFocusApp = counterLastFocus + 1 'de esta manera los eventos se registrarán de manera creciente'
                lastFocus = pathTitle
                counterLastFocus = counterFocusApp
                lastAction = counterFocusApp
            Else 'En caso contrario se actualizan el foco y el contador, además de registrarse en el diccionario'
                AddItemToList(Now.ToString + "#" + counterFocusApp.ToString + "#" + user)
                'AddItemToList(Now.ToString + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                dictionaryFocus.Add(pathTitle, counterFocusApp)
                counterFocusApp = counterLastFocus + 1
                lastFocus = pathTitle
                counterLastFocus = counterFocusApp
                lastAction = counterFocusApp
            End If
        Else
            'do nothing'
        End If
    End Sub

    Private Sub ClipboardEvent() Handles Me.ClipboardData
        Dim pathTitle As String = GetPathName()
        Dim typeAction As Integer = SearchValue(dictionaryIni, "Copy")
        If typeAction <> lastAction And pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            lastOrigin = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteEvent(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal pathTitle As String) Handles kbHook.PasteEvent
        If pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
        Else
            'do nothing'
        End If
    End Sub

    Private Sub UnregisterClipboardViewer()
        ChangeClipboardChain(Me.Handle, nextClipViewer)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
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

    Private Sub btnHook_Click(sender As Object, e As EventArgs) Handles btnHook.Click
        Dim SendText As String
        Dim notepad As IntPtr, editx As IntPtr

        SendText = "Hello this will write to notepad!"
        notepad = FindWindow("notepad", vbNullString)
        editx = FindWindowEx(notepad, 0&, "edit", vbNullString)
        Call SendMessage(editx, WM_SETTEXT, 0&, SendText)
    End Sub
End Class
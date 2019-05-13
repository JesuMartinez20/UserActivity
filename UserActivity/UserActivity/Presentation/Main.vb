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
    'Módulo que inicializa al formulario'
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadIni()
        ReadFocusDict()
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
    'Método que devuelve un diccionario con el contenido del fichero .ini'
    Private Sub ReadIni()
        CheckAndLoadFile(pathIni)
    End Sub
    'Método encargado de leer el diccionario de focos registrados hasta el momento'
    Private Sub ReadFocusDict()
        Dim sLine As String = ""
        Dim arrText As New ArrayList()
        If File.Exists(pathFocusDict) Then
            Dim sr As New System.IO.StreamReader(pathFocusDict)
            Do
                sLine = sr.ReadLine()
                If Not sLine Is Nothing Then
                    arrText.Add(sLine)
                    GetFocusAndID(sLine)
                End If
            Loop Until sLine Is Nothing
            sr.Close()
            'si existe el archivo, recuperamos la última entrada registrada y la incrementamos en 1'
            Dim kvp As KeyValuePair(Of String, Integer) = dictionaryFocus.Last
            lastAction = kvp.Value + 1
        Else 'en caso contrario la actualizamos a 0'
            lastAction = 0
        End If
    End Sub
    'Método encargado de almacenar cada línea del archivo FocusDictionary.txt en el dictionaryFocus (foco e identificador)'
    Private Sub GetFocusAndID(ByVal line As String)
        Dim intPos As Integer
        Dim focus As String
        Dim idFocus As String
        intPos = InStr(1, line, "#") 'posicion de "#"
        focus = Mid(line, 1, intPos - 1) 'Se extrae desde el inicio hasta la posicion de la coma -1 
        idFocus = Mid(line, intPos + 1) 'Se extrae desde la posicion de la coma + 1 hasta el final
        dictionaryFocus.Add(focus, Convert.ToInt32(idFocus))
    End Sub
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
            ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
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
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + user)
            ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " [" + key.ToString + "+" + vKey.ToString + "] en App: " + pathTitle + "#" + user)
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
            ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            focusWheel = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_FocusRise(ByVal typeAction As Integer, ByVal pathTitle As String) Handles fHook.FocusRise
        'Esta variable se encarga de contar los focos registrados'
        Dim counterFocusApp As Integer = typeAction
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle And pathTitle <> explorer Then
            'Si ya se ha registrado un foco determinado se busca en el diccionario de focos y se actualiza el foco actual'
            If dictionaryFocus.ContainsKey(pathTitle) Then
                Dim focusRegistered As Integer = dictionaryFocus.Where(Function(p) p.Key = pathTitle).FirstOrDefault.Value
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + focusRegistered.ToString + " en App: " + pathTitle + "#" + user)
                AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + focusRegistered.ToString + "#" + user)
                lastFocus = pathTitle
            ElseIf lastAction = 0 Then 'Si no existe el archivo FocusDictionary.txt, se inicializa el foco con el número correspondiente del archivo .ini'
                AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + "#" + user)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                dictionaryFocus.Add(pathTitle, counterFocusApp)
                SaveFocusDictionary(pathTitle, counterFocusApp)
                lastFocus = pathTitle
                lastAction = counterFocusApp + 1 'de esta manera se actualiza la última acción realizada'
            Else 'Si el foco no está registrado, el identificador corresponderá al del último almacenado en FocusDictionary.txt'
                counterFocusApp = lastAction
                AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + "#" + user)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                dictionaryFocus.Add(pathTitle, counterFocusApp)
                SaveFocusDictionary(pathTitle, counterFocusApp)
                lastFocus = pathTitle
                lastAction = lastAction + 1 'actualizamos la variable lastAction'
            End If
        Else
            'do nothing'
        End If
    End Sub

    Private Sub ClipboardEvent() Handles Me.ClipboardData
        Dim pathTitle As String = GetPathName()
        Dim typeAction As Integer = SearchValue(dictionaryIni, "Copy")
        If typeAction <> lastAction And pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
            lastOrigin = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteEvent(ByVal typeAction As Integer, ByVal key As Keys, ByVal vKey As Keys, ByVal pathTitle As String) Handles kbHook.PasteEvent
        If pathTitle <> explorer Then
            ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + user)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            ListBox1.TopIndex = ListBox1.Items.Count - 1
            lastAction = typeAction
        Else
            'do nothing'
        End If
    End Sub
    'Este método se encarga de crear un archivo o añadir contendido, si ya existe, del diccionario de focos'
    Private Sub SaveFocusDictionary(ByVal pathTitle As String, ByVal counterFocusApp As Integer)
        'Si el archivo ya está creado, se añade el contenido'
        If File.Exists(pathFocusDict) Then
            Dim sw As New System.IO.StreamWriter(pathFocusDict, True)
            sw.WriteLine(pathTitle + "#" + counterFocusApp.ToString)
            sw.Close()
        Else 'en caso contrario se crea e insertan los parámetros del método'
            Dim sw As New System.IO.StreamWriter(pathFocusDict)
            sw.WriteLine(pathTitle + "#" + counterFocusApp.ToString)
            sw.Close()
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
End Class
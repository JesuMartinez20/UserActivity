Imports System.IO

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    'Esta variable guarda el último foco capturado'
    Private lastFocus As String
    'Esta variable guarda la última acción registrada'
    Private lastAction As Integer
    'Esta variable guarda el útimo ID de foco registrado hasta el momento'
    Private lastIDFocus As Integer
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
#Region "EJECUCIÓN DEL PROGRAMA"
    'Módulo que inicializa al formulario'
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LaunchAppMinimized()
        ReadIni()
        ReadFocusDict()
        StartHooks()
        StartClipboard()
    End Sub

    Private Sub LaunchAppMinimized()
        Try
            If Me.WindowState = FormWindowState.Minimized Then
                Me.Visible = False
                NotifyIcon.Visible = True
                NotifyIcon.ShowBalloonTip(1000, "Notify Icon", "UserActivity is running...", ToolTipIcon.Info)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
#End Region
#Region "INICIALIZACIONES"
    'Se inicializan los hooks'
    Private Sub StartHooks()
        'Se inicializa el foco principal de la aplicación'
        'lastFocus = GetPathName()
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
#End Region
#Region "LECTURA ARCHIVO INI"
    'Método que lee un archivo .ini si existe y crea un diccionario con su contenido'
    Private Sub ReadIni()
        Try
            If File.Exists(pathIni) And File.ReadAllLines(pathIni).Length <> 0 Then
                Dim ini As New FicherosINI(pathIni)
                ReadPathLogAndFocusDict(ini)
                Dim arrayEvents() As String = ini.GetSection("TIPOS_EVENTOS")
                'Como se guardan pares de valores {llave = valor} el Step es igual a 2'
                For i = 0 To arrayEvents.Length - 1 Step 2
                    dictionaryIni.Add(arrayEvents(i), Convert.ToInt32(arrayEvents(i + 1)))
                Next
                'El número de inicio del foco será el número de la última entrada del .ini +1
                dictionaryIni.Add("InitActivaApp", Convert.ToInt32(arrayEvents.Last) + 1)
                'Se agrega el contador del cambio de foco'
                dictionaryIni.Add("CounterFocus", ini.GetInteger("FOCO", "CounterFocus"))
            Else
                Throw New Exception("No se puede abrir el archivo." & vbNewLine &
                                    "Compruebe que exista el archivo configBD.ini.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region
#Region "LECTURA FICHERO Y ALMACENAMIENTO EN FICHERO DE LOG Y CATÁLOGO FOCOS"
    'Método encargado de leer, si existe, el diccionario de focos registrados hasta el momento'
    Private Sub ReadFocusDict()
        Dim sLine As String = ""
        Dim sr As StreamReader
        If File.Exists(pathFocusDict) Then
            sr = New StreamReader(pathFocusDict)
            While Not sr.EndOfStream
                sLine = sr.ReadLine()
                If Not sLine Is Nothing Then
                    GetFocusAndID(sLine)
                End If
            End While
            sr.Close()
            'si contiene datos el archivo, recuperamos la última entrada registrada y la incrementamos en 1'
            If File.ReadAllLines(pathFocusDict).Length <> 0 Then
                lastIDFocus = dictionaryFocus.Last.Value
            End If
        Else 'en caso contrario la actualizamos a 0'
            lastIDFocus = 0
        End If
    End Sub
    'Método encargado de almacenar cada línea del archivo FocusDictionary.txt en el dictionaryFocus (foco e identificador)'
    Private Sub GetFocusAndID(ByVal line As String)
        Dim intPos As Integer
        Dim focus As String
        Dim idFocus As String
        intPos = InStr(1, line, "#") 'posicion de "#"
        'Si intPos es mayor que 0 significa que no se ha interpretrado como valor de la línea: ""
        If intPos > 0 Then
            focus = Mid(line, 1, intPos - 1) 'Se extrae desde el inicio hasta la posicion de la coma -1 
            idFocus = Mid(line, intPos + 1) 'Se extrae desde la posicion de la coma + 1 hasta el final
            dictionaryFocus.Add(focus, Convert.ToInt32(idFocus))
        Else
            'do nothing'
        End If
    End Sub
    'Establece la ruta del archivo log y del diccionario de focos'
    Private Sub ReadPathLogAndFocusDict(ByRef ini As FicherosINI)
        Dim pathLog As String = ini.GetString("FICHERO", "Log")
        If pathLog.Equals("pathExe") Then 'se establece la ruta por defecto'
            pathEvents = Application.StartupPath + "\EventosRegistrados.log"
        Else
            pathEvents = pathLog
            LogIsValid(pathEvents)
        End If
        Dim pathDictionary As String = ini.GetString("FICHERO", "CatalogoFocos")
        If pathDictionary.Equals("pathExe") Then
            pathFocusDict = Application.StartupPath + "\CatalogoFocos.txt"
        Else
            pathFocusDict = pathDictionary
            DictIsValid(pathFocusDict)
        End If
    End Sub
    'Comprueba que el path del fichero no es un directorio y que el fichero tengan la extensión (*.log)'
    Private Shared Sub LogIsValid(ByVal pathFile As String)
        Try
            If Path.GetExtension(pathFile) <> ".log" Or Directory.Exists(pathFile) Then
                Throw New Exception("Fichero no válido. Compruebe que la ruta no se trate de un directorio o que el fichero tenga la extensión adecuada (*.log)")
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Environment.Exit(0)
        End Try
    End Sub
    'Comprueba que el path del diccionario no es un directorio y que el fichero tengan la extensión (*.txt)'
    Private Shared Sub DictIsValid(ByVal pathFile As String)
        Try
            If Path.GetExtension(pathFile) <> ".txt" Or Directory.Exists(pathFile) Then
                Throw New Exception("Fichero no válido. Compruebe que la ruta no se trate de un directorio o que el fichero tenga la extensión adecaduda (*.txt)")
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Environment.Exit(0)
        End Try
    End Sub

    Private Sub SaveEvents(ByVal action As Integer)
        'Si el archivo ya está creado, se añade el contenido'
        If File.Exists(pathEvents) Then
            Dim sw As New System.IO.StreamWriter(pathEvents, True)
            sw.WriteLine(action.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        Else 'en caso contrario se crea e insertan los parámetros del método'
            Dim sw As New System.IO.StreamWriter(pathEvents)
            sw.WriteLine(action.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        End If
    End Sub
    'Se encarga de guardar eventos'
    Private Sub SaveEventsAndFocusDict(ByVal pathTitle As String, ByVal counterFocusApp As Integer)
        SaveEvents(counterFocusApp)
        dictionaryFocus.Add(pathTitle, counterFocusApp)
        SaveFocusDictionary(pathTitle, counterFocusApp)
    End Sub
    'Este método se encarga de crear un archivo o añadir contendido si ya existe del diccionario de focos'
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
#End Region
#Region "EVENTOS"
    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        Static focusKey As String
        If pathTitle <> focusKey And pathTitle.ToLower <> explorer.ToLower Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + DateTime.Now.ToString + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            SaveEvents(typeAction)
            lastAction = typeAction
            focusKey = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal vKey As Keys, ByVal pathTitle As String) Handles kbHook.CombKey
        Static lastkey As Keys
        If vKey <> lastkey And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " [" + key.ToString + "+" + vKey.ToString + "] en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            SaveEvents(typeAction)
            lastAction = typeAction
            lastkey = vKey
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        Static focusWheel As String
        If pathTitle <> focusWheel And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            SaveEvents(typeAction)
            lastAction = typeAction
            focusWheel = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_FocusRise(ByVal typeAction As Integer, ByVal pathTitle As String) Handles fHook.FocusRise
        'Esta variable se encarga de contar los focos registrados'
        Dim counterFocusApp As Integer = typeAction
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle And pathTitle <> explorer Then
            'Si ya se ha registrado un foco determinado se busca en el diccionario de focos y se actualiza el foco actual'
            If dictionaryFocus.ContainsKey(pathTitle) Then
                Dim focusRegistered As Integer = dictionaryFocus.Where(Function(p) p.Key = pathTitle).FirstOrDefault.Value
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + focusRegistered.ToString + " en App: " + pathTitle + "#" + userName)
                'AddItemToList(focusRegistered.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + user)
                SaveEvents(focusRegistered)
                UpdateFocusAndAction(pathTitle, focusRegistered)
            ElseIf lastIDFocus = 0 Then 'Si no existe el archivo FocusDictionary.txt, se inicializa el foco con el número correspondiente del archivo .ini'
                'AddItemToList(counterFocusApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                SaveEventsAndFocusDict(pathTitle, counterFocusApp)
                UpdateFocusAndAction(pathTitle, counterFocusApp)
                lastIDFocus = counterFocusApp
            Else 'Si el foco no está registrado, el identificador corresponderá al del último almacenado en FocusDictionary.txt'
                lastIDFocus += 1
                'AddItemToList(counterFocusApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                SaveEventsAndFocusDict(pathTitle, lastIDFocus)
                UpdateFocusAndAction(pathTitle, lastIDFocus)
            End If
        Else
            'do nothing'
        End If
    End Sub
    'Se encarga de actualizar los contadores de eventos de foco y de acción'
    Private Sub UpdateFocusAndAction(ByVal pathTitle As String, ByVal counterFocusApp As Integer)
        lastFocus = pathTitle
        lastAction = counterFocusApp
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

    Private Sub ClipboardEvent() Handles Me.ClipboardData
        Dim pathTitle As String = GetPathName()
        Dim typeAction As Integer = SearchValue(dictionaryIni, "Copy")
        If typeAction <> lastAction And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            SaveEvents(typeAction)
            lastAction = typeAction
            lastOrigin = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteEvent(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.PasteEvent
        If pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            SaveEvents(typeAction)
            lastAction = typeAction
        Else
            'do nothing'
        End If
    End Sub
#End Region

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
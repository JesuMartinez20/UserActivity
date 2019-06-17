Imports System.IO

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    Private flagBD As Boolean = True
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
        'ReadFocusCatalog()
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Try
            Me.Visible = True
            Me.WindowState = FormWindowState.Normal
            NotifyIcon.Visible = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Application.Exit()
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
        lastFocus = GetPathName()
        'diccionario vacio significa que el archivo .ini no se ha encontrado'
        If dictionaryIni.Count = 0 Or flagBD = False Then
            Application.Exit()
        Else
            kbHook = New KeyboardHook(dictionaryIni)
            mHook = New MouseHook(dictionaryIni)
            fHook = New FocusHook(dictionaryIni, dictionaryFocus)
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                ReadBD(ini)
                Dim arrayEvents() As String = ini.GetSection("TIPOS_EVENTOS")
                CheckAndInsertActions(arrayEvents)
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
            Application.Exit()
        End Try
    End Sub
    'Se leen los parámtetros necesarios para establecer la conexión con la BD'
    Private Sub ReadBD(ByRef ini As FicherosINI)
        Dim arrayBD() As String = ini.GetSection("BD")
        Dim conexionBD As AgentBD
        Try
            conexionBD = New AgentBD(arrayBD(1), arrayBD(3), arrayBD(5), arrayBD(7))
        Catch ex As Exception
            MessageBox.Show("No se ha podido conectar con la base de datos." & vbNewLine &
                            "Compruebe los parámetros en el archivo configBD.ini", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            flagBD = False
        End Try
    End Sub
#End Region
#Region "EVENTOS"
    Private Sub kbHook_KeyDown(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.KeyDown
        Static focusKey As String
        Dim ev As Events
        If pathTitle <> focusKey And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            ev = FillEvent(typeAction, pathTitle)
            InsertEvent(ev)
            lastAction = typeAction
            focusKey = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal typeAction As Integer, ByVal vKey As Keys, ByVal pathTitle As String) Handles kbHook.CombKey
        Static lastkey As Keys
        Dim ev As Events
        If vKey <> lastkey And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + userName)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            ev = FillEvent(typeAction, pathTitle)
            InsertEvent(ev)
            lastAction = typeAction
            lastkey = vKey
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal typeAction As Integer, ByVal pathTitle As String) Handles mHook.MouseWheel
        Static focusWheel As String
        Dim ev As Events
        If pathTitle <> focusWheel And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            ev = FillEvent(typeAction, pathTitle)
            InsertEvent(ev)
            lastAction = typeAction
            focusWheel = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_FocusRise(ByVal typeAction As Integer, ByVal pathTitle As String) Handles fHook.FocusRise
        Dim counterInitApp As Integer = typeAction
        Dim ev As Events
        Dim focus As Focus
        'Comparamos que el foco actual es diferente del foco más antiguo (lastfocus)'
        If lastFocus <> pathTitle And pathTitle <> explorer Then
            focus = New Focus
            Try
                'Se realiza una consulta cada vez que se cambie de foco para comprobar si existe en la BD'
                focus.ReadFocus()
                If focus.DaoAction.FocusCatalog.Count = 0 Then 'Si el catálogo está vacío, se inicializa guardando tanto el evento foco como su ID en el catálogo'
                    focus = FillFocus(pathTitle, counterInitApp)
                    InsertFocus(focus)
                    ev = FillEvent(counterInitApp, pathTitle)
                    'AddItemToList(counterInitApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                    InsertFocusEvent(ev)
                    UpdateFocusAndAction(pathTitle, counterInitApp)
                    lastIDFocus = counterInitApp 'se inicializa con el ID de foco respectivo del archivo .ini'
                ElseIf focus.DaoAction.FocusCatalog.ContainsKey(pathTitle) Then 'Si la referencia ya existe, únicamente se guarda el evento foco'
                    Dim focusInBD As Integer = focus.DaoAction.FocusCatalog.Where(Function(p) p.Key = pathTitle).FirstOrDefault.Value
                    'AddItemToList(focusInBD.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                    ev = FillEvent(focusInBD, pathTitle)
                    InsertFocusEvent(ev)
                    UpdateFocusAndAction(pathTitle, focusInBD)
                Else 'Si se trata de un evento no registrado, se guarda también tanto el evento como las referencias en el catálogo'
                    lastIDFocus += 1 'el siguiente evento registrado será el último + 1'
                    focus = FillFocus(pathTitle, lastIDFocus)
                    InsertFocus(focus)
                    ev = FillEvent(lastIDFocus, pathTitle)
                    'AddItemToList(lastIDFocus.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                    InsertFocusEvent(ev)
                    UpdateFocusAndAction(pathTitle, lastIDFocus)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(0)
            End Try
        Else
            'do nothing'
        End If
    End Sub
    'Se encarga de actualizar el último foco y la última acción añadida'
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub

    Private Sub ClipboardEvent() Handles Me.ClipboardData
        Dim pathTitle As String = GetPathName()
        Dim typeAction As Integer = SearchValue(dictionaryIni, "Copy")
        Dim ev As Events
        If typeAction <> lastAction And pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            ev = FillEvent(typeAction, pathTitle)
            InsertEvent(ev)
            lastAction = typeAction
            lastOrigin = pathTitle
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteEvent(ByVal typeAction As Integer, ByVal pathTitle As String) Handles kbHook.PasteEvent
        Dim pasteEv As PasteEvent
        If pathTitle <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            pasteEv = FillPasteEvent(typeAction, pathTitle, lastOrigin)
            InsertPasteEvent(pasteEv)
            lastAction = typeAction
        Else
            'do nothing'
        End If
    End Sub
#End Region
#Region "MÉTODOS PARA LA CORRECTA INTERACCIÓN CON LA BASE DE DATOS"
    'Se guarda el contenido de la tabla catalogo_focos en un diccionario y el úlitmo ID de foco añadido'
    Private Sub ReadFocusCatalog()
        Dim focus As New Focus
        Dim kvp As KeyValuePair(Of String, Integer)
        Try
            focus.ReadFocus()
            'Si el catálogo de focos contiene elementos se ingresan en el diccionario de focos'
            If focus.DaoAction.FocusCatalog.Count <> 0 Then
                For Each kvp In focus.DaoAction.FocusCatalog
                    dictionaryFocus.Add(kvp.Key, kvp.Value)
                Next
                lastIDFocus = focus.DaoAction.FocusCatalog.Last.Value
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Devuelve un objeto con todos los datos del foco'
    Private Function FillFocus(pathTitle As String, idFocus As Integer) As Focus
        Dim focus As New Focus
        focus.IdFocus = idFocus
        focus.Focus = pathTitle.Replace("\", "\\")
        Return focus
    End Function
    'Inserta los focos en la tabla catalogo_focos'
    Private Sub InsertFocus(focus As Focus)
        Try
            focus.InsertFocus()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Inserta los eventos en la tabla eventos de la bd'
    Private Sub InsertFocusEvent(ev As Events)
        Try
            ev.InsertFocusEvent()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Se encarga de comprobar si la tabla acciones está completa con los datos del .ini, en caso contrario, los inserta'
    Private Sub CheckAndInsertActions(arrayEvents() As String)
        Dim action As New Action
        Try
            action.ReadAction()
            If action.DaoAction.ActionsList.Count = 0 Then
                InsertAction(arrayEvents, action)
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Inserta las acciones en la tabla acciones de la bd'
    Private Sub InsertAction(arrayEvents() As String, action As Action)
        'Como se guardan pares de valores {llave = valor} el Step es igual a 2'
        For i = 0 To arrayEvents.Length - 1 Step 2
            action.IdAction = Convert.ToInt32(arrayEvents(i + 1))
            action.Action = arrayEvents(i)
            Try
                action.InsertAction()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(0)
            End Try
        Next
    End Sub
    'Devuelve un evento completo de tipo: type,scroll,combkey y focus'
    Private Function FillEvent(typeAction As Integer, origin As String) As Events
        Dim ev As Events = New Events
        ev.Fecha = Now.ToString("yyyy-MM-dd HH:mm:ss")
        ev.IdAction = typeAction
        'De esta manera se inserta correctamente el path en la base de datos'
        ev.AppOrigin = origin.Replace("\", "\\")
        ev.User = userName
        Return ev
    End Function
    'Inserta los eventos en la tabla eventos de la bd'
    Private Sub InsertEvent(ev As Events)
        Try
            ev.InsertEvent()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Devuelve un evento completo de tipo paste'
    Private Function FillPasteEvent(typeAction As Integer, destiny As String, origin As String) As PasteEvent
        Dim pasteEv As PasteEvent = New PasteEvent
        pasteEv.Fecha = Now.ToString("yyyy-MM-dd HH:mm:ss")
        pasteEv.IdAction = typeAction
        'De esta manera se inserta correctamente el path en la base de datos'
        pasteEv.AppOrigin = origin.Replace("\", "\\")
        pasteEv.AppDestiny = destiny.Replace("\", "\\")
        pasteEv.User = userName
        Return pasteEv
    End Function
    'Inserta los eventos paste en la tabla eventos_paste de la bd'
    Private Sub InsertPasteEvent(pasteEV As PasteEvent)
        Try
            pasteEV.InsertPasteEvent()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
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
        AgentBD.CloseBD()
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
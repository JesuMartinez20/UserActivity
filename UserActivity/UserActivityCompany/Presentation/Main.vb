Imports System.IO

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    'Variable encargada de comprobar el cambio de foco'
    Private focusThreshold As Integer
    'Esta variable gestiona la aplicación activa más reciente'
    Private currentAppName As String
    'Esta variable gestiona la acción más reciente registrada'
    Private currentActionId As Integer
    'Esta variable gestiona el útimo id de aplicación registrada hasta el momento'
    Private lastAppId As Integer
    'Diccionario encargado de almacenar las diferentes acciones (incluida la apps activas) y sus correspondientes ids''
    Private actions As New Dictionary(Of String, Integer)
    'Variable para controlar el origen del último copy'
    Private lastCopyOrigin As String
    'Array con las acciones recuperadas del archivo .ini'
    Private arrayActions() As String
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
        ReadActionsCatalog()
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
        currentAppName = GetPathName()
        'diccionario vacío significa que el archivo .ini no se ha encontrado'
        If actions.Count = 0 Then
            Application.Exit()
        Else
            kbHook = New KeyboardHook(actions)
            mHook = New MouseHook(actions)
            fHook = New FocusHook(focusThreshold)
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
    'Método encargado de leer un archivo .ini si existe y almacena su contenido en un diccionario'
    Private Sub ReadIni()
        Try
            If File.Exists(pathIni) And File.ReadAllLines(pathIni).Length <> 0 Then
                Dim ini As New FicherosINI(pathIni)
                ReadFilesPath(ini)
                arrayActions = ini.GetSection("ACCIONES")
                For i = 0 To arrayActions.Length - 1
                    actions.Add(arrayActions(i), i + 1)
                Next
                'Se almacena el valor correspondiente al campo FocusThresHold del fichero .ini'
                focusThreshold = ini.GetInteger("FOCO", "FocusThresHold")
            Else
                Throw New Exception("No se puede abrir el archivo." & vbNewLine &
                                    "Compruebe que exista el archivo configBD.ini.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region
#Region "LECTURA DE FICHEROS"
    'Método encargado de leer si existe, el catalogo de apps registradas hasta el momento'
    Private Sub ReadActionsCatalog()
        Dim sLine As String = "", hasData As Boolean
        Dim sr As StreamReader
        If File.Exists(pathActionCatalog) Then
            sr = New StreamReader(pathActionCatalog)
            While Not sr.EndOfStream
                sLine = sr.ReadLine()
                If Not sLine Is Nothing Then
                    GetActionAndID(sLine)
                    hasData = True
                End If
            End While
            sr.Close()
            'si contiene datos el archivo, recuperamos la última entrada registrada'
            If hasData Then
                lastAppId = actions.Last.Value
            Else
                lastAppId = 0
            End If
        Else 'Si no existe registro, se inicializa a 0 el último ID de acción'
            lastAppId = 0
        End If
    End Sub
    'Método encargado de extraer la app y su correspondiente id del archivo CatalogoApps.txt'
    Private Sub GetActionAndID(ByVal line As String)
        Dim intPos As Integer
        Dim app As String
        Dim idApp As String
        intPos = InStr(1, line, "#") 'posicion de "#"
        'Si intPos es mayor que 0 significa que no se ha interpretrado como valor de la línea: ""
        If intPos > 0 Then
            app = Mid(line, 1, intPos - 1) 'Se extrae desde el inicio hasta la posicion de la coma -1 
            idApp = Mid(line, intPos + 1) 'Se extrae desde la posicion de la coma + 1 hasta el final
            actions.Add(app, Convert.ToInt32(idApp))
        Else
            'do nothing'
        End If
    End Sub
    'Establece la ruta del archivo log y del catalogo de apps'
    Private Sub ReadFilesPath(ByRef ini As FicherosINI)
        Dim myPath As String = ini.GetString("FICHERO", "ActionLog")
        If myPath.Equals("pathExe") Then 'se establece la ruta por defecto'
            pathActionLog = Application.StartupPath + "\EventosRegistrados.log"
        Else
            pathActionLog = myPath
            PathIsValid(pathActionLog)
        End If
        myPath = ini.GetString("FICHERO", "ActionCatalog")
        If myPath.Equals("pathExe") Then
            pathActionCatalog = Application.StartupPath + "\CatalogoApps.txt"
        Else
            pathActionCatalog = myPath
            CatalogIsValid(pathActionCatalog)
        End If
    End Sub
#End Region
#Region "COMPROBACIÓN FICHEROS"
    'Comprueba que el path del fichero no es un directorio y que el fichero tenga la extensión (*.log)'
    Private Shared Sub PathIsValid(ByVal pathFile As String)
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
    'Comprueba que el path del diccionario no es un directorio y que el fichero tenga la extensión (*.txt)'
    Private Shared Sub CatalogIsValid(ByVal pathFile As String)
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
#End Region
#Region "EVENTOS"
    Private Sub kbHook_KeyDown(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.KeyDown
        Static lastAppName As String 'Controla la app activa'
        If appName <> lastAppName And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + DateTime.Now.ToString + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            LogAction(actionId)
            lastAppName = appName
        Else
            'do nothing'
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal actionId As Integer, ByVal vKey As Keys, ByVal appName As String) Handles kbHook.CombKey
        Static lastkey As Keys
        If vKey <> lastkey And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " [" + key.ToString + "+" + vKey.ToString + "] en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            LogAction(actionId)
            lastkey = vKey
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal actionId As Integer, ByVal appName As String) Handles mHook.MouseWheel
        Static lastAppName As String
        If appName <> lastAppName And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            LogAction(actionId)
            lastAppName = appName
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseLeftDown(ByVal typeAction As Integer, ByVal point As Point) Handles mHook.MouseLeftDown
        LogClicks(typeAction, point)
    End Sub

    Private Sub mHook_MouseRightDown(ByVal typeAction As Integer, ByVal point As Point) Handles mHook.MouseRightDown
        LogClicks(typeAction, point)
    End Sub

    Private Sub fHook_AppRise(ByVal appName As String) Handles fHook.AppRise
        Dim iniAppId As Integer
        'Se compara la aplicación activa actualmente con la activa por última vez, si es diferente, se gestionará dependiendo de 3 casos'
        If currentAppName <> appName And appName <> explorer Then
            If actions.ContainsKey(appName) Then 'Si la app ya está registrada, se recuperan su nombre y su id, posteriormente se registra'
                Dim actionIdRegistered As Integer = actions.Where(Function(p) p.Key = appName).FirstOrDefault.Value
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + focusRegistered.ToString + " en App: " + pathTitle + "#" + userName)
                'AddItemToList(focusRegistered.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + user)
                LogAction(actionIdRegistered)
                UpdateCurrentAction(appName, actionIdRegistered)
            ElseIf lastAppId = 0 Then 'Si no existe previamente un catálogo de apps, se inicializa guardando tanto la aplicación activa como su id correspondiente'
                'AddItemToList(counterFocusApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                iniAppId = arrayActions.Count + 1 'El Id de inicio será el número de acciones registradas +1'
                LogAction(iniAppId)
                actions.Add(appName, iniAppId)
                LogCatalog(appName, iniAppId)
                UpdateCurrentAction(appName, iniAppId)
                lastAppId = iniAppId 'se actualiza el id de aplicación activa'
            Else 'Si la app no está registrada en el catálogo, se procede a su registro'
                lastAppId += 1 'la siguiente app a registrar será el última almacenada + 1'
                'AddItemToList(counterFocusApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
                'AddItemToList(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + counterFocusApp.ToString + " en App: " + pathTitle + "#" + user)
                LogAction(lastAppId)
                actions.Add(appName, lastAppId)
                LogCatalog(appName, lastAppId)
                UpdateCurrentAction(appName, lastAppId)
            End If
        Else
            'do nothing'
        End If
    End Sub
    'Se encarga de actualizar los contadores de eventos de foco y de acción'
    Private Sub UpdateCurrentAction(ByVal newAppName As String, ByVal newActionId As Integer)
        currentAppName = newAppName
        currentActionId = newActionId
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

    Private Sub ClipboardAction() Handles Me.ClipboardData
        Dim originApp As String = GetPathName()
        Dim actionId As Integer = SearchValue(actions, "Copy")
        If actionId <> currentActionId And originApp <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            LogAction(actionId)
            currentActionId = actionId
            lastCopyOrigin = originApp
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteAction(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.PasteAction
        If appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss,fff") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            LogAction(actionId)
            currentActionId = actionId
        Else
            'do nothing'
        End If
    End Sub
#End Region
#Region "ALMACENAMIENTO DE LAS ACCIONES Y APPS"
    Private Sub LogAction(ByVal actionId As Integer)
        'Si el archivo ya está creado, se añade el contenido'
        If File.Exists(pathActionLog) Then
            Dim sw As New System.IO.StreamWriter(pathActionLog, True)
            sw.WriteLine(actionId.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        Else 'en caso contrario se crea e insertan los valores correspondientes'
            Dim sw As New System.IO.StreamWriter(pathActionLog)
            sw.WriteLine(actionId.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        End If
    End Sub

    Private Sub LogClicks(ByVal actionId As Integer, ByVal pt As Point)
        'Si el archivo ya está creado, se añade el contenido'
        If File.Exists(pathActionLog) Then
            Dim sw As New System.IO.StreamWriter(pathActionLog, True)
            sw.WriteLine(actionId.ToString + " X:" + pt.X.ToString + " Y:" + pt.Y.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        Else 'en caso contrario se crea e insertan los valores correspondientes'
            Dim sw As New System.IO.StreamWriter(pathActionLog)
            sw.WriteLine(actionId.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        End If
    End Sub

    Private Sub LogCatalog(ByVal appName As String, ByVal idApp As Integer)
        'Si el archivo ya está creado, se añade el contenido'
        If File.Exists(pathActionCatalog) Then
            Dim sw As New System.IO.StreamWriter(pathActionCatalog, True)
            sw.WriteLine(appName + "#" + idApp.ToString)
            sw.Close()
        Else 'en caso contrario se crea e insertan los parámetros del método'
            Dim sw As New System.IO.StreamWriter(pathActionCatalog)
            sw.WriteLine(appName + "#" + idApp.ToString)
            sw.Close()
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
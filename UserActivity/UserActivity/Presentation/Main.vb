Imports System.IO

Public Class Main
    Private WithEvents kbHook As KeyboardHook
    Private WithEvents mHook As MouseHook
    Private WithEvents fHook As FocusHook
    'Private flagBD As Boolean = true
    'Variable encargada de comprobar el cambio de foco'
    Private focusThreshold As Integer
    'Esta variable gestiona la aplicación activa más reciente'
    Private currentAppName As String
    'Esta variable gestiona la acción más reciente registrada'
    Private currentActionId As Integer
    'Esta variable gestiona el útimo Id de aplicación registrada hasta el momento'
    Private lastAppId As Integer
    'Diccionario encargado de almacenar las diferentes acciones (incluida la app activa) y sus correspondientes Ids'
    Private actions As New Dictionary(Of String, Integer)
    'Variable encargada de gestionar el origen del último copy'
    Private lastCopyOrigin As String
    'Diccionario encargardo de gestionar las aplicaciones registradas hasta el momento'
    Private appsRegistered As New Dictionary(Of String, Integer)
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
        ReadCatalogApps()
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
        currentAppName = GetPathName()
        'diccionario vacio significa que el archivo .ini no se ha encontrado'
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
    'Método que lee un archivo .ini si existe y se extrae su contenido'
    Private Sub ReadIni()
        Try
            If File.Exists(pathIni) And File.ReadAllLines(pathIni).Length <> 0 Then
                Dim ini As New FicherosINI(pathIni)
                'ReadBD(ini)
                'Se almacenan las acciones correspondiente a la sección ACCIONES del fichero .ini'
                arrayActions = ini.GetSection("ACCIONES")
                CheckCatalogActions(arrayActions)
                'Se vuelcan las acciones del array a un diccionario llamado actions para su posterior gestión'
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
            Application.Exit()
        End Try
    End Sub
    'Se leen los parámtetros necesarios para establecer la conexión con la BD'
    'Private Sub ReadBD(ByRef ini As FicherosINI)
    'Dim arrayBD() As String = ini.GetSection("BD")
    'Dim conexionBD As AgentBD
    'Try
    '       conexionBD = New AgentBD(arrayBD(1), arrayBD(3), arrayBD(5), arrayBD(7))
    'Catch ex As Exception
    '       MessageBox.Show("No se ha podido conectar con la base de datos." & vbNewLine &
    '                      "Compruebe los parámetros en el archivo configBD.ini", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '     flagBD = False
    'End Try
    'End Sub
#End Region
#Region "EVENTOS"
    Private Sub kbHook_KeyDown(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.KeyDown
        Static lastAppName As String 'Controla la app activa'
        Dim action As Action
        If appName <> lastAppName And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            action = SaveAction(actionId, appName)
            InsertAction(action)
            lastAppName = appName
        Else
            'do nothing'
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal actionId As Integer, ByVal vKey As Keys, ByVal appName As String) Handles kbHook.CombKey
        Static lastkey As Keys 'Controla la segunda tecla activa procedente de una combinación de teclas'
        Dim action As Action
        If vKey <> lastkey And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App:" + pathTitle + "#" + userName)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            action = SaveAction(actionId, appName)
            InsertAction(action)
            lastkey = vKey
        Else
            'do nothing'
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal actionId As Integer, ByVal appName As String) Handles mHook.MouseWheel
        Static lastAppName As String 'Controla la app activa'
        Dim action As Action
        If appName <> lastAppName And appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "#" + " en App:" + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            action = SaveAction(actionId, appName)
            InsertAction(action)
            lastAppName = appName
        Else
            'do nothing'
        End If
    End Sub

    Private Sub fHook_AppRise(ByVal appName As String) Handles fHook.AppRise
        Dim iniAppId As Integer
        Dim action As Action
        Dim app As Catalog_Apps
        'Se compara la apliación activa actualmente con la activa por última vez, si es diferente, se gestionará dependiendo de 3 casos'
        If currentAppName <> appName And appName <> explorer Then
            If appsRegistered.Count = 0 Then 'Si el catálogo está vacío se inicializa guardando tanto la aplicación activa como su Id correspondiente'
                iniAppId = arrayActions.Count + 1 'El Id de inicio será el número de acciones registradas +1'
                app = SaveApp(appName, iniAppId)
                InsertAppCatalog(app)
                appsRegistered.Add(appName, iniAppId)
                action = SaveAction(iniAppId, appName)
                'AddItemToList(counterInitApp.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                InsertApp(action)
                UpdateCurrentAction(appName, iniAppId)
                lastAppId = iniAppId 'se actualiza el Id de aplicación activa'
            ElseIf appsRegistered.ContainsKey(appName) Then 'Si la app ya está registrada, se recuperan su nombre y su id, posteriormente se registra'
                Dim actionIdRegistered As Integer = appsRegistered.Where(Function(p) p.Key = appName).FirstOrDefault.Value 'id de la app registrada en el diccionario'
                'AddItemToList(focusInBD.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                action = SaveAction(actionIdRegistered, appName)
                InsertApp(action)
                UpdateCurrentAction(appName, actionIdRegistered)
            Else 'Si la app no está registrada en el catálogo, se procede a registrar dicha app donde sea necesario'
                lastAppId += 1 'la siguiente app a registrar será el última almacenada + 1'
                app = SaveApp(appName, lastAppId)
                InsertAppCatalog(app)
                appsRegistered.Add(appName, lastAppId)
                action = SaveAction(lastAppId, appName)
                'AddItemToList(lastIDFocus.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
                InsertApp(action)
                UpdateCurrentAction(appName, lastAppId)
            End If
        Else
            'do nothing'
        End If
    End Sub
    'Se encarga de actualizar la aplicación activa actulamente, así como la acción'
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
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub

    Private Sub ClipboardAction() Handles Me.ClipboardData
        Dim originApp As String = GetPathName()
        Dim actionId As Integer = SearchValue(actions, "Copy")
        Dim action As Action
        If actionId <> currentActionId And originApp <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + "en App: " + pathTitle + "#" + user)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            action = SaveAction(actionId, originApp)
            InsertAction(action)
            currentActionId = actionId
            lastCopyOrigin = originApp
        Else
            'do nothing'
        End If
    End Sub

    Private Sub PasteAction(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.PasteAction
        Dim p As ActionPaste
        If appName <> explorer Then
            'ListBox1.Items.Add(typeAction.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            'ListBox1.Items.Add(Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + typeAction.ToString + " en App:" + pathTitle + "#" + user + "#" + "Origen: " + lastOrigin)
            'ListBox1.TopIndex = ListBox1.Items.Count - 1
            p = SavePasteAction(actionId, lastCopyOrigin, appName)
            InsertPasteAction(p)
            currentActionId = actionId
        Else
            'do nothing'
        End If
    End Sub
#End Region
#Region "MÉTODOS PARA LA CORRECTA INTERACCIÓN CON LA BASE DE DATOS"
    'Se lee el catalogo de apps provenientes de la BD y se vuelca la información en un diccionario llamado appsRegistered para evitar consultas innecesarias a la BD'
    Private Sub ReadCatalogApps()
        Dim ca As New Catalog_Apps
        Dim kvp As KeyValuePair(Of String, Integer)
        Try
            ca.ReadCatalogApps()
            'Si el catálogo de apps contiene elementos se ingresan en el diccionario de apps'
            If ca.DaoCatalog.AppCatalog.Count <> 0 Then
                For Each kvp In ca.DaoCatalog.AppCatalog
                    appsRegistered.Add(kvp.Key, kvp.Value)
                Next
                lastAppId = ca.DaoCatalog.AppCatalog.Last.Value 'Se recupera el último id de la última aplicación registrada en la BD para gestionar los ids posteriores'
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show("leer foco", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Recupera toda la información de una app activa'
    Private Function SaveApp(appName As String, appId As Integer) As Catalog_Apps
        Dim ca As New Catalog_Apps
        ca.IdApp = appId
        'ca.App = appName.Replace("\", "\\")
        ca.App = appName
        Return ca
    End Function
    'Inserta una app en la tabla catalogo_apps de la bd'
    Private Sub InsertAppCatalog(ca As Catalog_Apps)
        Try
            ca.InsertApp()
        Catch ex As Exception
            MessageBox.Show("Insertar foco diccionario", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Inserta una app en la tabla acciones_app de la bd'
    Private Sub InsertApp(action As Action)
        Try
            action.InsertAppAction()
        Catch ex As Exception
            MessageBox.Show("Insertar foco evento", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Comprueba el catálogo de acciones provenientes de la BD. Si existen acciones no se hace nada, en caso contrario, se insertan las acciones del fichero .ini en la BD'
    Private Sub CheckCatalogActions(arrayActions() As String)
        Dim ca As New Catalog_Actions
        Try
            ca.ReadCatalogActions()
            If ca.DaoCatalog.ActionsList.Count = 0 Then
                InsertActionCatalog(arrayActions, ca)
            Else
                'do nothing'
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Inserta las acciones en la tabla catalogo_acciones de la bd'
    Private Sub InsertActionCatalog(arrayActions() As String, ca As Catalog_Actions)
        For i = 0 To arrayActions.Length - 1
            ca.IdAction = i + 1
            ca.Action = arrayActions(i)
            Try
                ca.InsertAction()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(0)
            End Try
        Next
    End Sub
    'Recupera toda la información de una acción concreta'
    Private Function SaveAction(actionId As Integer, appName As String) As Action
        Dim action As Action = New Action
        action.Fecha = Now.ToString("yyyy-MM-dd HH:mm:ss")
        action.IdAction = actionId
        'De esta manera se inserta correctamente el path en la base de datos'
        'action.App = appName.Replace("\", "\\")
        action.App = appName
        action.User = userName
        Return action
    End Function
    'Inserta una acción concreta en la tabla acciones de la bd'
    Private Sub InsertAction(action As Action)
        Try
            action.InsertAction()
        Catch ex As Exception
            MessageBox.Show("insertar evento", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub
    'Devuelve un evento completo de tipo paste'
    Private Function SavePasteAction(actionId As Integer, originApp As String, destinyApp As String) As ActionPaste
        Dim p As ActionPaste = New ActionPaste
        p.Fecha = Now.ToString("yyyy-MM-dd HH:mm:ss")
        p.IdAction = actionId
        'De esta manera se inserta correctamente el path en la base de datos'
        'p.AppOrigin = originApp.Replace("\", "\\")
        p.AppOrigin = originApp
        'p.AppDestiny = destinyApp.Replace("\", "\\")
        p.AppDestiny = destinyApp
        p.User = userName
        Return p
    End Function
    'Inserta los eventos paste en la tabla eventos_paste de la bd'
    Private Sub InsertPasteAction(p As ActionPaste)
        Try
            p.InsertPasteAction()
        Catch ex As Exception
            MessageBox.Show("insertar evento paste", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
        AgentBD.getAgent.Conexion.Close()
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
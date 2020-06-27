'********************************************************'
'Aplicación elaborada por: Jesús Martínez Manrique'
'Versión Live'
'********************************************************'
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
    'Esta variable gestiona el útimo id de la aplicación registrada hasta el momento'
    Private lastAppId As Integer
    'Diccionario encargado de gestionar las acciones registradas hasta el momento'
    Private actions As New Dictionary(Of String, Integer)
    'Diccionario encargardo de gestionar las aplicaciones registradas hasta el momento'
    Private appsRegistered As New Dictionary(Of String, Integer)
    'Array con las acciones recuperadas del archivo .ini'
    Private arrayActions() As String
    'Obtiene el Handle de la ventana actual para el clipboard'
    Private nextClipViewer As IntPtr
    'Evento del Clipboard'
    Public Event ClipboardData(ByVal clipboardText As String)
#Region "EJECUCIÓN DEL PROGRAMA"
    'Módulo que inicializa al formulario'
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LaunchAppMinimized()
        ReadIni()
        ReadCatalogActions()
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
    'Método encargado de leer un archivo .ini, si existe, y almacena su contenido en un diccionario'
    Private Sub ReadIni()
        Try
            If File.Exists(pathIni) And File.ReadAllLines(pathIni).Length > 0 Then
                Dim ini As New INIFiles(pathIni)
                ReadFilesPath(ini)

                arrayActions = ini.GetSection("ACCIONES")
                For i = 0 To arrayActions.Length - 1
                    actions.Add(arrayActions(i), i + 1)
                Next
                'Se almacena el valor correspondiente al campo FocusThresHold del fichero .ini'
                focusThreshold = ini.GetInteger("FOCO", "FocusThresHold")
            Else
                Throw New Exception("No se puede abrir el archivo." & vbNewLine &
                                    "Compruebe que exista el archivo config.ini.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region
#Region "LECTURA DE FICHEROS"
    'Método encargado de leer, si existe, el catalogo de apps registradas hasta el momento'
    Private Sub ReadCatalogActions()
        Dim sLine As String = ""
        Dim sLineArray As String()
        Dim sr As StreamReader

        If File.Exists(pathCatalogActions) Then
            sr = New StreamReader(pathCatalogActions)
            While Not sr.EndOfStream
                sLine = sr.ReadLine()
                If Not sLine Is Nothing Then
                    sLineArray = sLine.Split(":")
                    Dim app = sLineArray(1)
                    Dim idApp = sLineArray(0)
                    appsRegistered.Add(app, Convert.ToInt32(idApp))
                End If
            End While
            sr.Close()
            lastAppId = appsRegistered.Last.Value 'se recupera el id de la última aplicación registarda
        End If
    End Sub

    'Establece la ruta del archivo log y del catalogo de apps'
    Private Sub ReadFilesPath(ByRef ini As INIFiles)
        Dim myPath As String = ini.GetString("FICHERO", "ActionLog")
        If myPath.Equals("pathExe") Then 'se establece la ruta por defecto'
            pathLogActions = Application.StartupPath + "\AccionesRegistradas.log"
        Else
            pathLogActions = myPath
            PathIsValid(pathLogActions)
        End If

        myPath = ini.GetString("FICHERO", "ActionCatalog")
        If myPath.Equals("pathExe") Then
            pathCatalogActions = Application.StartupPath + "\CatalogoApps.txt"
        Else
            pathCatalogActions = myPath
            CatalogIsValid(pathCatalogActions)
        End If
    End Sub
#End Region
#Region "COMPROBACIÓN FICHEROS"
    'Comprueba que el path del fichero no es un directorio y que el fichero tenga la extensión (*.log)'
    Private Shared Sub PathIsValid(ByVal pathFile As String)
        Try
            If Path.GetExtension(pathFile) <> ".log" Or Directory.Exists(pathFile) Then
                Throw New Exception("Fichero no válido. Compruebe que la ruta no se trate de un directorio o que el fichero tenga la extensión adecuada (*.log)")
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

        If appName <> lastAppName And appName <> explorer And appName <> userActivity Then
            ActionLog(actionId)
            lastAppName = appName
        End If
    End Sub

    Private Sub kbHook_CombKey(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.CombKey
        If actionId <> currentActionId And appName <> explorer And appName <> userActivity Then
            ActionLog(actionId)
            currentActionId = actionId
        End If
    End Sub

    Private Sub mHook_MouseWheel(ByVal actionId As Integer, ByVal appName As String) Handles mHook.MouseWheel
        Static lastAppName As String

        If appName <> lastAppName And appName <> explorer And appName <> userActivity Then
            ActionLog(actionId)
            lastAppName = appName
        End If
    End Sub

    Private Sub fHook_AppRise(ByVal appName As String) Handles fHook.AppRise
        Dim iniAppId As Integer
        'Se compara la aplicación activa actualmente con la activa por última vez, si es diferente, se gestionará dependiendo de 3 casos'
        If currentAppName <> appName And appName <> explorer And appName <> userActivity Then
            If appsRegistered.Count = 0 Then 'Si el catálogo está vacío se inicializa guardando tanto la aplicación activa como su id correspondiente'
                iniAppId = arrayActions.Count + 1 'El id de inicio será el número de acciones registradas + 1'
                ActionLog(iniAppId)
                appsRegistered.Add(appName, iniAppId)
                CatalogLog(appName, iniAppId)
                UpdateCurrentAction(appName, iniAppId)
                lastAppId = iniAppId 'se actualiza el id de aplicación activa'
            ElseIf appsRegistered.ContainsKey(appName) Then 'Si la app ya está registrada, se recuperan su nombre y su id'
                Dim actionIdRegistered As Integer = appsRegistered.Where(Function(p) p.Key = appName).FirstOrDefault.Value
                ActionLog(actionIdRegistered)
                UpdateCurrentAction(appName, actionIdRegistered)
            Else 'Si la app no está registrada en el catálogo, se procede a registrar la misma'
                lastAppId += 1 'la siguiente app a registrar será el última almacenada + 1'
                ActionLog(lastAppId)
                appsRegistered.Add(appName, lastAppId)
                CatalogLog(appName, lastAppId)
                UpdateCurrentAction(appName, lastAppId)
            End If
        End If
    End Sub
    'Se encarga de actualizar los contadores de eventos de foco y de acción'
    Private Sub UpdateCurrentAction(ByVal newAppName As String, ByVal newActionId As Integer)
        currentAppName = newAppName
        currentActionId = newActionId
    End Sub
    'Se sobrecarga el Window Procedure para recibir mensajes del Clipboard'
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
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0)
        End Try
    End Sub

    Private Sub ClipboardAction() Handles Me.ClipboardData
        Dim originApp As String = GetPathName()
        Dim actionId As Integer = SearchValue(actions, "Copy")

        If actionId <> currentActionId And originApp <> explorer And originApp <> userActivity Then
            ActionLog(actionId)
            currentActionId = actionId
        End If
    End Sub

    Private Sub PasteAction(ByVal actionId As Integer, ByVal appName As String) Handles kbHook.PasteAction
        If appName <> explorer And appName <> userActivity Then
            ActionLog(actionId)
            currentActionId = actionId
        End If
    End Sub
#End Region
#Region "ALMACENAMIENTO DE LAS ACCIONES Y APPS"
    Private Sub ActionLog(ByVal actionId As Integer)
        'Si el archivo ya está creado, se añade al contenido del mismo'
        If File.Exists(pathLogActions) Then
            Dim sw As New System.IO.StreamWriter(pathLogActions, True)
            sw.WriteLine(actionId.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        Else 'en caso contrario se crean e insertan los valores correspondientes'
            Dim sw As New System.IO.StreamWriter(pathLogActions)
            sw.WriteLine(actionId.ToString + "#" + Now.ToString("yyyy-MM-dd HH:mm:ss") + "#" + userName)
            sw.Close()
        End If
    End Sub

    Private Sub CatalogLog(ByVal appName As String, ByVal idApp As Integer)
        'Si el archivo ya está creado, se añade al contenido del mismo'
        If File.Exists(pathCatalogActions) Then
            Dim sw As New System.IO.StreamWriter(pathCatalogActions, True)
            sw.WriteLine(idApp.ToString + ":" + appName)
            sw.Close()
        Else 'en caso contrario se crean e insertan los valores correspondientes'
            Dim sw As New System.IO.StreamWriter(pathCatalogActions)
            sw.WriteLine(idApp.ToString + ":" + appName)
            sw.Close()
        End If
    End Sub
#End Region
    Private Sub UnregisterClipboardViewer()
        ChangeClipboardChain(Me.Handle, nextClipViewer)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If Not fHook Is Nothing Then
                fHook.FocusThread.Abort()
            Else
                Application.Exit()
            End If
        Catch ThreadAbortException As Exception
            Debug.WriteLine(ThreadAbortException.Message)
        End Try
        UnregisterClipboardViewer()
    End Sub
End Class
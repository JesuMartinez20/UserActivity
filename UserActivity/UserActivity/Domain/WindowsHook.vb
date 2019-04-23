Public Class WindowsHook
    Private historyID As New List(Of Integer)
    Private historyWindowName As New List(Of String)
    Private historyProName As New List(Of String)

    Private degBy As Control
    Private delay As Integer
    Private mainThread As System.Threading.Thread
    'Se declaran los eventos de Windows'
    Public Event WindowsClosed(ByVal processName As String, ByVal processID As Integer)
    Public Event WindowsOpened(ByVal processName As String, ByVal processID As Integer)
    Public Event WindowsModified(ByVal processName As String, ByVal processID As Integer)
    Private Delegate Sub WindowsDelegate(ByVal processName As String, ByVal processID As Integer)
    'Constructor'
    Public Sub New(ByRef con As Control, Optional _delay As Integer = 1000)
        LoadHistory()
        degBy = con
        delay = _delay
        mainThread = New System.Threading.Thread(AddressOf Runner)
        mainThread.Start()
    End Sub
    'Llamada asíncrona para delegar el control a este método'
    Private Sub WinOpenDelegate(ByVal _processName As String, ByVal _processID As Integer)
        If degBy.InvokeRequired Then
            degBy.Invoke(New WindowsDelegate(AddressOf WinOpenDelegate), _processName, _processID)
        Else
            RaiseEvent WindowsOpened(_processName, _processID)
        End If
    End Sub

    Private Sub WinClosedDelegate(ByVal _processName As String, ByVal _processID As integer)
        If degBy.InvokeRequired Then
            degBy.Invoke(New WindowsDelegate(AddressOf WinClosedDelegate), _processName, _processID)
        Else
            RaiseEvent WindowsClosed(_processName, _processID)
        End If
    End Sub

    'Hilo que llama al método RunCheck() cada seg'
    Private Sub Runner()
        While True
            RunCheck()
            System.Threading.Thread.Sleep(delay)
        End While
    End Sub
    'Método que comprueba los procesos activos del sistema cada seg'
    Private Sub RunCheck()
        Dim tempID As New List(Of Integer)
        Dim tempWindowName As New List(Of String)
        Dim tempProName As New List(Of String)
        For Each p As Process In Process.GetProcesses
            tempID.Add(p.Id)
            tempProName.Add(p.MainWindowTitle)
            tempWindowName.Add(p.ProcessName)
        Next
        Dim temp As Integer
        For i As Integer = 0 To tempID.Count - 1
            'Obtiene el índice correspondiente al nombre del proceso en la variable temp'
            temp = historyID.IndexOf(tempID.Item(i))
            If temp <> -1 Then
                'si se encuentra en la lista temp entonces significa que son el mismo proceso, por lo que se borra de la lista history'
                historyID.RemoveAt(temp)
                historyProName.RemoveAt(temp)
                historyWindowName.RemoveAt(temp)
            Else
                'si es la primera vez que aparece este proceso se lanza'
                WinOpenDelegate(tempProName.Item(i), tempID.Item(i))
                'RaiseEvent WindowsOpened(tempWindowName.Item(i), tempProName.Item(i))
            End If
        Next
        'en la lista historyID aparecerán los procesos con más inactividad, es decir, los que no se usan o se han cerrado'
        For i As Integer = 0 To historyID.Count - 1
            'se lanza el evento de cerrar el proceso'
            WinClosedDelegate(historyWindowName.Item(i), historyID.Item(i))
            'RaiseEvent WindowsClosed(tempWindowName.Item(i), tempProName.Item(i))
        Next
        'actualizamos las dos listas, inicializándolas con los mismos elementos'
        historyID = tempID
        historyProName = tempProName
        historyWindowName = tempWindowName
    End Sub
    'Método que recopila los procesos activos del sistema al crear la instancia del objeto de esta clase'
    Private Sub LoadHistory()
        historyID.Clear()
        historyWindowName.Clear()
        historyProName.Clear()
        For Each p As Process In Process.GetProcesses()
            historyID.Add(p.Id)
            historyProName.Add(p.MainWindowTitle)
            historyWindowName.Add(p.ProcessName)
        Next
    End Sub
    'Se para el hilo'
    Public Sub StopAll()
        If mainThread IsNot Nothing Then
            mainThread.Abort()
        End If
    End Sub
End Class

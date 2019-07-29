Imports System.Threading

Public Class FocusHook
    Private _focusThread As Threading.Thread
    Private _focusThresHold As Integer
    'Getter'
    Public Property FocusThread As Thread
        Get
            Return _focusThread
        End Get
        Set(value As Thread)
            _focusThread = value
        End Set
    End Property
    'Constructor'
    Public Sub New(ByVal focusThresHold As Integer)
        _focusThread = New Threading.Thread(AddressOf GetFocusInfo)
        _focusThread.Start()
        _focusThresHold = focusThresHold
    End Sub
    'Evento de foco'
    Public Event AppRise(ByVal appName As String)
    'Método que cada cierto tiempo lanza un evento de la aplicación activa'
    Public Sub GetFocusInfo()
        While True
            Dim newAppName As String = GetAppName()
            'Si no se consigue la aplicación activa no se hace nada'
            If newAppName = Nothing Then
                'do nothing'
            Else
                RaiseEvent AppRise(newAppName)
            End If
            System.Threading.Thread.Sleep(_focusThresHold)
        End While
    End Sub
End Class

Imports System.Threading

Public Class FocusHook
    Private _focusThread As Threading.Thread
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
    Public Sub New()
        _focusThread = New Threading.Thread(AddressOf GetFocusInfo)
        _focusThread.Start()
    End Sub
    'Evento de foco'
    Public Event FocusRise(ByVal action As Integer, ByVal focus As String)
    'Método que cada cierto tiempo lanza un evento del foco actual'
    Public Sub GetFocusInfo()
        While True
            Dim currentFocus As String = GetPathName()
            Dim action As Integer = TypeAction.ActivaApp
            RaiseEvent FocusRise(action, currentFocus)
            System.Threading.Thread.Sleep(5000)
        End While
    End Sub
End Class

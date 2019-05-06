Imports System.Threading

Public Class FocusHook
    Private _focusThread As Threading.Thread
    Private _dictionary As Dictionary(Of String, Integer)
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
    Public Sub New(ByVal dictionary As Dictionary(Of String, Integer))
        _focusThread = New Threading.Thread(AddressOf GetFocusInfo)
        _focusThread.Start()
        _dictionary = dictionary
    End Sub
    'Evento de foco'
    Public Event FocusRise(ByVal action As Integer, ByVal focus As String)
    'Método que cada cierto tiempo lanza un evento del foco actual'
    Public Sub GetFocusInfo()
        While True
            Dim currentFocus As String = GetPathName()
            Dim action As Integer = SearchValue(_dictionary, "InitActivaApp")
            Dim counter As Integer = SearchValue(_dictionary, "CounterFocus")
            'Si no se consigue capturar el foco actual o se trata del Explorer no se lanza'
            If currentFocus = Nothing Or currentFocus.Equals("C:\WINDOWS\Explorer.EXE") Then
                'do nothing'
            Else
                RaiseEvent FocusRise(action, currentFocus)
            End If
            System.Threading.Thread.Sleep(counter)
        End While
    End Sub
End Class

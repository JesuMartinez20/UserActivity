Public Class Events
    Private _fecha As String
    Private _idAction As Integer
    Private _appOrigin As String
    Private _user As String
    Private _daoEvent As DAOEvent
    'Constructor'
    Public Sub New()
        Me._daoEvent = New DAOEvent
    End Sub
#Region "GETTER Y SETTER"
    Public Property Fecha As String
        Get
            Return _fecha
        End Get
        Set(value As String)
            _fecha = value
        End Set
    End Property

    Public Property IdAction As Integer
        Get
            Return _idAction
        End Get
        Set(value As Integer)
            _idAction = value
        End Set
    End Property

    Public Property AppOrigin As String
        Get
            Return _appOrigin
        End Get
        Set(value As String)
            _appOrigin = value
        End Set
    End Property

    Public Property User As String
        Get
            Return _user
        End Get
        Set(value As String)
            _user = value
        End Set
    End Property

    Public Property DaoTSEvent As DAOEvent
        Get
            Return _daoEvent
        End Get
        Set(value As DAOEvent)
            _daoEvent = value
        End Set
    End Property
#End Region
    Public Sub InsertEvent()
        Me._daoEvent.InsertEvent(Me)
    End Sub

    Public Sub InsertFocusEvent()
        Me._daoEvent.InsertFocusEvent(Me)
    End Sub
End Class

Imports UserActivity

Public Class TypeScrollEvent
    Private _fecha As String
    Private _idAction As Integer
    Private _appOrigin As String
    Private _user As String
    Private _daoTSEvent As DAOTSEvent

    Public Sub New()
        Me._daoTSEvent = New DAOTSEvent
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

    Public Property DaoTSEvent As DAOTSEvent
        Get
            Return _daoTSEvent
        End Get
        Set(value As DAOTSEvent)
            _daoTSEvent = value
        End Set
    End Property
#End Region
    Public Sub InsertEvent()
        Me._daoTSEvent.InsertDAOEvent(Me)
    End Sub

    Public Sub InsertAction(ByVal action As String)
        Me._daoTSEvent.InsertDAOAction(Me, action)
    End Sub
End Class

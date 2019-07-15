Public Class Action
    Private _fecha As String
    Private _idAction As Integer
    Private _app As String
    Private _user As String
    Private _daoAction As DAOAction
    'Constructor'
    Public Sub New()
        Me._daoAction = New DAOAction
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

    Public Property App As String
        Get
            Return _app
        End Get
        Set(value As String)
            _app = value
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

    Public Property DaoAction As DAOAction
        Get
            Return _daoAction
        End Get
        Set(value As DAOAction)
            _daoAction = value
        End Set
    End Property
#End Region
    Public Sub InsertAction()
        Me._daoAction.InsertAction(Me)
    End Sub

    Public Sub InsertAppAction()
        Me._daoAction.InsertAppAction(Me)
    End Sub
End Class

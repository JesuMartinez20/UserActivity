Imports UserActivity

Public Class Action
    Private _idAction As Integer
    Private _action As String
    Private _daoAction As DAOAction

    Public Sub New()
        Me._daoAction = New DAOAction
    End Sub
#Region "GETTER Y SETTER"
    Public Property IdAction As Integer
        Get
            Return _idAction
        End Get
        Set(value As Integer)
            _idAction = value
        End Set
    End Property

    Public Property Action As String
        Get
            Return _action
        End Get
        Set(value As String)
            _action = value
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

    Public Sub Read()
        Me._daoAction.Read()
    End Sub
End Class

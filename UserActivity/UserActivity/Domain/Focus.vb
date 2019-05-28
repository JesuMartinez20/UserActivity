Imports UserActivity

Public Class Focus
    Private _idFocus As Integer
    Private _focus As String
    Private _daoAction As DAOAction
    'Constructor'
    Public Sub New()
        Me.DaoAction = New DAOAction
    End Sub
#Region "GETTER Y SETTER"
    Public Property IdFocus As Integer
        Get
            Return _idFocus
        End Get
        Set(value As Integer)
            _idFocus = value
        End Set
    End Property

    Public Property Focus As String
        Get
            Return _focus
        End Get
        Set(value As String)
            _focus = value
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
    Public Sub ReadFocus()
        Me._daoAction.ReadFocus()
    End Sub

    Public Sub InsertFocus()
        Me._daoAction.InsertFocus(Me)
    End Sub
End Class

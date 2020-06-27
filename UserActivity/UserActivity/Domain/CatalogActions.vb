Imports UserActivity

Public Class CatalogActions
    Private _idAction As Integer
    Private _action As String
    Private _daoCatalog As DAOCatalog
    'Constructor'
    Public Sub New()
        Me._daoCatalog = New DAOCatalog
    End Sub
#Region "GETTER Y SETTER"
    Public Property IdAction As Integer
        Get
            Return Me._idAction
        End Get
        Set(value As Integer)
            Me._idAction = value
        End Set
    End Property

    Public Property Action As String
        Get
            Return Me._action
        End Get
        Set(value As String)
            Me._action = value
        End Set
    End Property

    Public Property DaoCatalog As DAOCatalog
        Get
            Return Me._daoCatalog
        End Get
        Set(value As DAOCatalog)
            Me._daoCatalog = value
        End Set
    End Property
#End Region
    Public Sub InsertAction()
        Me._daoCatalog.InsertAction(Me)
    End Sub

    Public Sub ReadCatalogActions()
        Me._daoCatalog.ReadAction()
    End Sub
End Class

Imports UserActivity

Public Class Catalog_Actions
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

    Public Property DaoCatalog As DAOCatalog
        Get
            Return _daoCatalog
        End Get
        Set(value As DAOCatalog)
            _daoCatalog = value
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

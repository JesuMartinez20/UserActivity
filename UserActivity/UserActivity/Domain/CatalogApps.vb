Imports UserActivity

Public Class CatalogApps
    Private _idApp As Integer
    Private _app As String
    Private _daoCatalog As DAOCatalog
    'Constructor'
    Public Sub New()
        Me._daoCatalog = New DAOCatalog
    End Sub
#Region "GETTER Y SETTER"
    Public Property IdApp As Integer
        Get
            Return Me._idApp
        End Get
        Set(value As Integer)
            Me._idApp = value
        End Set
    End Property

    Public Property App As String
        Get
            Return Me._app
        End Get
        Set(value As String)
            Me._app = value
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
    Public Sub ReadCatalogApps()
        Me._daoCatalog.ReadApp()
    End Sub

    Public Sub InsertApp()
        Me._daoCatalog.InsertApp(Me)
    End Sub
End Class

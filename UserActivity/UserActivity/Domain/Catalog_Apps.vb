Imports UserActivity

Public Class Catalog_Apps
    Private _idApp As Integer
    Private _app As String
    Private _daoCatalog As DAOCatalog
    'Constructor'
    Public Sub New()
        Me.DaoCatalog = New DAOCatalog
    End Sub
#Region "GETTER Y SETTER"
    Public Property IdApp As Integer
        Get
            Return _idApp
        End Get
        Set(value As Integer)
            _idApp = value
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

    Public Property DaoCatalog As DAOCatalog
        Get
            Return _daoCatalog
        End Get
        Set(value As DAOCatalog)
            _daoCatalog = value
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

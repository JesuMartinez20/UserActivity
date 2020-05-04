Public Class ActionPaste
    Private _fecha As String
    Private _idAction As Integer
    Private _appOrigin As String
    Private _appDestiny As String
    Private _user As String
    Private _daoAction As DAOAction
    'Constructor'
    Public Sub New()
        Me._daoAction = New DAOAction
    End Sub
#Region "GETTER Y SETTER"
    Public Property Fecha As String
        Get
            Return Me._fecha
        End Get
        Set(value As String)
            Me._fecha = value
        End Set
    End Property

    Public Property IdAction As Integer
        Get
            Return Me._idAction
        End Get
        Set(value As Integer)
            Me._idAction = value
        End Set
    End Property

    Public Property AppOrigin As String
        Get
            Return Me._appOrigin
        End Get
        Set(value As String)
            Me._appOrigin = value
        End Set
    End Property

    Public Property AppDestiny As String
        Get
            Return Me._appDestiny
        End Get
        Set(value As String)
            Me._appDestiny = value
        End Set
    End Property

    Public Property User As String
        Get
            Return Me._user
        End Get
        Set(value As String)
            Me._user = value
        End Set
    End Property
#End Region
    Public Sub InsertPasteAction()
        Me._daoAction.InsertPasteAction(Me)
    End Sub
End Class

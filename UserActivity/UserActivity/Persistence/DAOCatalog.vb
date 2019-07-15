Imports UserActivity

Public Class DAOCatalog
    Private _actionsList As List(Of Integer)
    Private _appCatalog As Dictionary(Of String, Integer) 'Este diccionario guardará el contenido de la tabla catalogo_apps'
#Region "GETTER"
    Public Property ActionsList As List(Of Integer)
        Get
            Return _actionsList
        End Get
        Set(value As List(Of Integer))
            _actionsList = value
        End Set
    End Property

    Public Property AppCatalog As Dictionary(Of String, Integer)
        Get
            Return _appCatalog
        End Get
        Set(value As Dictionary(Of String, Integer))
            _appCatalog = value
        End Set
    End Property
#End Region
    Public Sub New()
        _actionsList = New List(Of Integer)
        _appCatalog = New Dictionary(Of String, Integer)
    End Sub

    Public Function InsertAction(ByVal ca As Catalog_Actions) As Integer
        Return AgentBD.getAgent.Insert("INSERT INTO catalogo_acciones VALUES(" & ca.IdAction & ",'" & ca.Action & "');")
    End Function

    Public Sub ReadAction()
        'Buenas prácticas mencionan el uso de Using para cerrar y liberar recursos cuando se hace una consulta'
        Using reader As MySql.Data.MySqlClient.MySqlDataReader = AgentBD.getAgent.Read("SELECT id_accion FROM catalogo_acciones")
            If reader.HasRows Then
                While reader.Read
                    Me._actionsList.Add(reader.GetInt32(0))
                End While
            End If
        End Using
    End Sub

    Public Function InsertApp(ByVal app As Catalog_Apps) As Integer
        Return AgentBD.getAgent.Insert("INSERT INTO catalogo_apps VALUES(" & app.IdApp & ",'" & app.App & "');")
    End Function

    Public Sub ReadApp()
        Using reader As MySql.Data.MySqlClient.MySqlDataReader = AgentBD.getAgent.Read("SELECT * FROM catalogo_apps")
            If reader.HasRows Then
                While reader.Read
                    Me.AppCatalog.Add(reader.GetString(1), reader.GetInt32(0))
                End While
            End If
        End Using
    End Sub
End Class

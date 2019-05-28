Imports UserActivity

Public Class DAOAction
    Private _actionsList As List(Of Integer)
    Private _focusCatalog As Dictionary(Of String, Integer) 'Este diccionario guardará el contenido de catalogo_focos'
#Region "GETTER"
    Public Property ActionsList As List(Of Integer)
        Get
            Return _actionsList
        End Get
        Set(value As List(Of Integer))
            _actionsList = value
        End Set
    End Property

    Public Property FocusCatalog As Dictionary(Of String, Integer)
        Get
            Return _focusCatalog
        End Get
        Set(value As Dictionary(Of String, Integer))
            _focusCatalog = value
        End Set
    End Property
#End Region
    Public Sub New()
        _actionsList = New List(Of Integer)
        _focusCatalog = New Dictionary(Of String, Integer)
    End Sub

    Public Function InsertAction(ByVal action As Action) As Integer
        Return AgentBD.Insert("INSERT INTO acciones VALUES(" & action.IdAction & ",'" & action.Action & "');")
    End Function

    Public Sub ReadAction()
        'Buenas prácticas mencionan el uso de Using para cerrar y liberar recursos cuando se hace una consulta'
        Using reader As MySql.Data.MySqlClient.MySqlDataReader = AgentBD.Read("SELECT id_accion FROM acciones")
            If reader.HasRows Then
                While reader.Read
                    Me._actionsList.Add(reader.GetInt32(0))
                End While
            End If
        End Using
    End Sub

    Public Function InsertFocus(ByVal focus As Focus) As Integer
        Return AgentBD.Insert("INSERT INTO catalogo_focos VALUES(" & focus.IdFocus & ",'" & focus.Focus & "');")
    End Function

    Public Sub ReadFocus()
        Using reader As MySql.Data.MySqlClient.MySqlDataReader = AgentBD.Read("SELECT * FROM catalogo_focos")
            If reader.HasRows Then
                While reader.Read
                    Me.FocusCatalog.Add(reader.GetString(1), reader.GetInt32(0))
                End While
            End If
        End Using
    End Sub
End Class

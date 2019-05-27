Imports UserActivity

Public Class DAOAction
    Private _actionsList As List(Of Integer)
#Region "GETTER"
    Public Property ActionsList As List(Of Integer)
        Get
            Return _actionsList
        End Get
        Set(value As List(Of Integer))
            _actionsList = value
        End Set
    End Property
#End Region
    Public Sub New()
        _actionsList = New List(Of Integer)
    End Sub

    Public Function InsertAction(ByVal action As Action) As Integer
        Return AgentBD.Insert("INSERT INTO acci VALUES(" & action.IdAction & ",'" & action.Action & "');")
    End Function

    Public Sub Read()
        'Buenas prácticas mencionan el uso de Using para cerrar y liberar recursos cuando se hace una consulta'
        Using reader As MySql.Data.MySqlClient.MySqlDataReader = AgentBD.Read("SELECT id_accion FROM accion")
            If reader.HasRows Then
                While reader.Read
                    Me._actionsList.Add(reader.GetInt32(0))
                End While
            End If
        End Using
    End Sub
End Class

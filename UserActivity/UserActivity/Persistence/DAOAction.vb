Imports UserActivity

Public Class DAOAction
    Private actionsList As List(Of Integer)

    Public Sub New()
        actionsList = New List(Of Integer)
    End Sub

    Public Function InsertAction(ByVal action As Action) As Integer
        Return AgentBD.Insert("INSERT INTO accion VALUES(" & action.IdAction & ",'" & action.Action & "');")
    End Function

    Public Sub Read()
        Dim reader As MySql.Data.MySqlClient.MySqlDataReader
        reader = AgentBD.Read("SELECT id_accion FROM accion")
        If reader.HasRows Then
            While reader.Read
                Me.actionsList.Add(reader.GetInt32(0))
            End While
        End If
    End Sub
End Class

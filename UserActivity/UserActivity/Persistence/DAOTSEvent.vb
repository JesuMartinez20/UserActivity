Public Class DAOTSEvent
    Public Function InsertDAOEvent(ByVal tsEvent As TypeScrollEvent) As Integer
        Return AgentBD.Insert("INSERT INTO evento_type_scroll (fecha,id_accion,app_origen,usuario) VALUES('" & tsEvent.Fecha & "'," & tsEvent.IdAction & ",'" & tsEvent.AppOrigin & "','" & tsEvent.User & "');")
    End Function

    Public Function InsertDAOAction(ByVal tsEvent As TypeScrollEvent, ByVal action As String) As Integer
        Return AgentBD.Insert("INSERT INTO accion VALUES(" & tsEvent.IdAction & ",'" & action & "');")
    End Function
End Class

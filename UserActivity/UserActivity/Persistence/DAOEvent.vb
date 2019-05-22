Public Class DAOEvent
    Public Function InsertTSEvent(ByVal tsEvent As TypeScrollEvent) As Integer
        Return AgentBD.Insert("INSERT INTO evento_type_scroll (fecha,id_accion,app_origen,usuario) VALUES('" & tsEvent.Fecha & "'," & tsEvent.IdAction & ",'" & tsEvent.AppOrigin & "','" & tsEvent.User & "');")
    End Function
End Class

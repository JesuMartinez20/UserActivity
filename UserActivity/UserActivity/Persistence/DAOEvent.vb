Imports UserActivity

Public Class DAOEvent
    Public Function InsertEvent(ByVal ev As GeneralEvent) As Integer
        Return AgentBD.Insert("INSERT INTO eveno VALUES" &
            "('" & ev.Fecha & "'," & ev.IdAction & ",'" & ev.AppOrigin & "','" & ev.User & "');")
    End Function

    Public Function InsertPasteEvent(ByVal pasteEv As PasteEvent) As Integer
        Return AgentBD.Insert("INSERT INTO evento_paste VALUES" &
                              "('" & pasteEv.Fecha & "'," & pasteEv.IdAction & ",'" & pasteEv.AppOrigin & "'," &
                              "'" & pasteEv.AppDestiny & "','" & pasteEv.User & "');")
    End Function
End Class

Imports UserActivity

Public Class DAOEvent
    Public Function InsertEvent(ByVal ev As Events) As Integer
        Return AgentBD.Insert("INSERT INTO eventos VALUES" &
            "('" & ev.Fecha & "'," & ev.IdAction & ",'" & ev.AppOrigin & "','" & ev.User & "');")
    End Function

    Public Function InsertPasteEvent(ByVal pasteEv As PasteEvent) As Integer
        Return AgentBD.Insert("INSERT INTO eventos_paste VALUES" &
                              "('" & pasteEv.Fecha & "'," & pasteEv.IdAction & ",'" & pasteEv.AppOrigin & "'," &
                              "'" & pasteEv.AppDestiny & "','" & pasteEv.User & "');")
    End Function

    Public Function InsertFocusEvent(ByVal ev As Events) As Integer
        Return AgentBD.Insert("INSERT INTO eventos_foco VALUES" &
            "('" & ev.Fecha & "'," & ev.IdAction & ",'" & ev.AppOrigin & "','" & ev.User & "');")
    End Function
End Class

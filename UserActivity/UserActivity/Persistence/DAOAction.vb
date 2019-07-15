Imports UserActivity

Public Class DAOAction
    Public Function InsertAction(ByVal action As Action) As Integer
        Return AgentBD.getAgent.Insert("INSERT INTO acciones VALUES" &
            "('" & action.Fecha & "'," & action.IdAction & ",'" & action.App & "','" & action.User & "');")
    End Function

    Public Function InsertPasteAction(ByVal p As ActionPaste) As Integer
        Return AgentBD.getAgent.Insert("INSERT INTO acciones_paste VALUES" &
                              "('" & p.Fecha & "'," & p.IdAction & ",'" & p.AppOrigin & "'," &
                              "'" & p.AppDestiny & "','" & p.User & "');")
    End Function

    Public Function InsertAppAction(ByVal action As Action) As Integer
        Return AgentBD.getAgent.Insert("INSERT INTO acciones_app VALUES" &
            "('" & action.Fecha & "'," & action.IdAction & ",'" & action.App & "','" & action.User & "');")
    End Function
End Class

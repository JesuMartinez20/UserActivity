Imports MySql.Data.MySqlClient
Public Class AgentBD
    Private Shared conexion As MySqlConnection
    'Constructor'
    Public Sub New(ByRef server As String, ByRef userID As String, ByRef password As String, ByRef database As String)
        conexion = New MySqlConnection
        conexion.ConnectionString = "server=" & server & ";" & "user id=" & userID & ";" & "password=" & password & ";" & "database=" & database & ";"
        conexion.Open()
    End Sub

    Public Shared Function Insert(ByVal sql As String) As Integer
        Dim com As New MySqlCommand(sql, conexion)
        Return com.ExecuteNonQuery()
    End Function
    'Método para cerrar la bd'
    Public Shared Sub CloseBD()
        conexion.Close()
        conexion.Dispose()
    End Sub
End Class

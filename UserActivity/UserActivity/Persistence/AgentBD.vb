Imports MySql.Data.MySqlClient
Public Class AgentBD
    Private _conexion As MySqlConnection
    Private Shared minstance As AgentBD
    Private Shared server As String
    Private Shared userID As String
    Private Shared password As String
    Private Shared database As String
    'GETTER'
    Public Property Conexion As MySqlConnection
        Get
            Return _conexion
        End Get
        Set(value As MySqlConnection)
            _conexion = value
        End Set
    End Property
    'Constructor'
    Public Sub New()
        ReadBD()
        conexion = New MySqlConnection
        conexion.ConnectionString = "server=" & server & ";" & "user id=" & userID & ";" & "password=" & password & ";" & "database=" & database & ";"
        conexion.Open()
    End Sub
    'Se leen los parámtetros necesarios para establecer la conexión con la BD'
    Private Sub ReadBD()
        Dim ini As New FicherosINI(pathIni)
        Dim arrayBD() As String = ini.GetSection("BD")
        server = arrayBD(1)
        userID = arrayBD(3)
        password = arrayBD(5)
        database = arrayBD(7)
    End Sub
    'Esta función recupera la instancia creada con la BD'
    Public Shared Function getAgent() As AgentBD
        If minstance Is Nothing Then
            minstance = New AgentBD
        End If
        Return minstance
    End Function

    Public Function Insert(ByVal sql As String) As Integer
        Dim com As New MySqlCommand(sql, conexion)
        Return com.ExecuteNonQuery()
    End Function

    Public Function Read(ByVal sql As String) As MySqlDataReader
        Dim com As New MySqlCommand(sql, conexion)
        Return com.ExecuteReader()
    End Function
End Class

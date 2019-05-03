Imports System.Runtime.InteropServices
Imports System.Text

Public Class FicherosINI
     Private strFilename As String
    ' Funciones del API de Windows para ficheros INI
    Private Declare Ansi Function WritePrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String,
      ByVal lpKeyName As String, ByVal lpString As String,
      ByVal lpFileName As String) As Integer

    Private Declare Ansi Function FlushPrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer,
      ByVal lpKeyName As Integer, ByVal lpString As Integer,
      ByVal lpFileName As String) As Integer
    ' Constructor, para aceptar el fichero INI'
    Public Sub New(ByVal Filename As String)
        strFilename = Filename
    End Sub
    'Propiedad sólo lectura con nombre de fichero'
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property
    'Función para leer cadena de texto (string) de fichero INI'
    Public Function GetString(ByVal sFileName As String, ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "") As String
        Dim charCount As Integer
        Dim sRetVal As StringBuilder
        '
        sRetVal = New StringBuilder(256)
        '
        charCount = GetPrivateProfileString(sSection, sKeyName, sDefault, sRetVal, sRetVal.Capacity, sFileName)
        If charCount = 0 Then
            Return sDefault
        Else
            Return sRetVal.ToString
        End If
    End Function
    'Función para leer un valor numérico del fichero INI'
    Public Function GetInteger(ByVal Section As String, ByVal Key As String, Optional ByVal nDefault As Integer = 0) As Integer
        Return GetPrivateProfileInt(Section, Key, nDefault, strFilename)
    End Function
    'Función para leer una sección entera de un fichero INI'
    Public Function GetSection(ByVal sSection As String) As String()
        Dim aSeccion() As String
        Dim n As Integer
        '
        ReDim aSeccion(0)
        '
        Dim sBuffer = New String(ChrW(0), 32767)
        '
        n = GetPrivateProfileSection(sSection, sBuffer, sBuffer.Length, strFilename)
        '
        If n > 0 Then
            '
            ' Cortar la cadena al número de caracteres devueltos
            ' menos el último que indican el final de la cadena
            sBuffer = sBuffer.Substring(0, n - 1).TrimEnd()
            ' Cada elemento estará separado por un Chr(0)
            ' y cada valor estará en la forma: clave = valor
            aSeccion = sBuffer.Split(New Char() {ChrW(0), "="c})
        End If
        ' Devolver el array
        Return aSeccion
    End Function
    'Función para leer un valor booleano de fichero INI
    Public Function GetBoolean(ByVal Section As String, ByVal Key As String, ByVal [Default] As Boolean) As Boolean
        Return (GetPrivateProfileInt(Section, Key, CInt([Default]), strFilename) = 1)
    End Function
    'Función para escribir valor de cadena (string) en fichero INI
    Public Sub WriteString(ByVal Section As String, ByVal Key As String, ByVal Value As String)
        WritePrivateProfileString(Section, Key, Value, strFilename)
        Flush()
    End Sub
    'Función para escribir valor numérico en fichero INI
    Public Sub WriteInteger(ByVal Section As String, ByVal Key As String, ByVal Value As Integer)
        ' Writes an integer to your INI file
        WriteString(Section, Key, CStr(Value))
        Flush()
    End Sub
    'Función para escribir valor booleano en fichero INI
    Public Sub WriteBoolean(ByVal Section As String, ByVal Key As String, ByVal Value As Boolean)
        ' Writes a boolean to your INI file
        WriteString(Section, Key, CStr(Math.Abs(CInt(Value))))
        Flush()
    End Sub
    'Guarda los cambios de la caché en fichero INI
    Private Sub Flush()
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub
End Class

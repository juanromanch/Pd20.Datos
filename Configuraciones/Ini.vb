Public Class Ini
    ' Funciones API
    Private Declare Ansi Function GetPrivateProfileString Lib “kernel32.dll” Alias “GetPrivateProfileStringA” (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As System.Text.StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

    Private Declare Ansi Function WritePrivateProfileString Lib “kernel32.dll” Alias “WritePrivateProfileStringA” (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    Private Declare Ansi Function GetPrivateProfileInt Lib “kernel32.dll” Alias “GetPrivateProfileIntA” (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer

    Private Declare Ansi Function FlushPrivateProfileString Lib “kernel32.dll” Alias “WritePrivateProfileStringA” (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer

    Dim strFilename As String

    ' Constructor, acepta un nombre de fichero (si no existe se creará)
    Public Sub New(ByVal Filename As String)
        strFilename = Filename
    End Sub

    ' Propiedad para Read-only
    ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    Public Function ObtenerString(ByVal Seccion As String, ByVal Clave As String, ByVal Defecto As String) As String
        ' Devuelve una cadena desde tu archivo INI
        Dim intCharCount As Integer
        Dim objResult As New System.Text.StringBuilder(256)
        intCharCount = GetPrivateProfileString(Seccion, Clave, Defecto, objResult, objResult.Capacity, strFilename)

        If intCharCount > 0 Then
            ObtenerString = Left(objResult.ToString, intCharCount)

        Else
            ObtenerString = ""
        End If
    End Function

    Public Function ObtenerInteger(ByVal Seccion As String, ByVal Clave As String, ByVal Defecto As Integer) As Integer
        ' Devuelve un número desde tu archivo INI
        Return GetPrivateProfileInt(Seccion, Clave, Defecto, strFilename)
    End Function

    Public Function ObtenerBoolean(ByVal Seccion As String, ByVal Clave As String, ByVal Defecto As Boolean) As Boolean
        ' Devuelve un valo lógico desde un archivo INI
        Return ObtenerString(Seccion, Clave, Defecto)
    End Function

    Public Sub EscrbirString(ByVal Seccion As String, ByVal Clave As String, ByVal Valor As String)
        ' Escribe una cadena a un archivo INI
        WritePrivateProfileString(Seccion, Clave, Valor, strFilename)
        Flush()
    End Sub

    Public Sub EscrbirInteger(ByVal Seccion As String, ByVal Clave As String, ByVal Valor As Integer)
        ' Escribe un número a un archivo INI
        EscrbirString(Seccion, Clave, CStr(Valor))
        Flush()
    End Sub

    Public Sub EscrbirBoolean(ByVal Seccion As String, ByVal Clave As String, ByVal Valor As Boolean)
        ' Escribe un valor logico a un arhcivo INI
        EscrbirString(Seccion, Clave, Valor)
        Flush()
    End Sub

    Private Sub Flush()
        ' Stores all the cached changes to your INI file
        ' Guarda todos los cambios al archivo INI
        FlushPrivateProfileString(0, 0, 0, strFilename)
    End Sub
End Class

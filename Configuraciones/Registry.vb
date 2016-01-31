Module Registry

    'Códigos de error y sus valores: help.netop.com/support/errorcodes/win32_error_codes.htm

    'ADVAPI32 Registry API Bas File.
    ' This file was not writen by me but I like to thank who did write it

    ' --------------------------------------------------------------------
    ' ADVAPI32
    ' --------------------------------------------------------------------
    ' function prototypes, constants, and type definitions
    ' for Windows 32-bit Registry API

    Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
    Public Const HKEY_CURRENT_USER As Integer = &H80000001
    Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Public Const HKEY_USERS As Integer = &H80000003
    Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
    Public Const ERROR_SUCCESS As Short = 0
    Public Const ERROR_NONE As Short = 0
    Public sKeys As Collection

    Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String,
        ByRef lpcbName As Integer, ByVal lpReserved As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByRef lpftLastWriteTime As Integer) As Integer


    Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer


    Public Sub GetKeyNames(ByVal hKey As Integer, ByVal strPath As String)
        Dim Cnt, TKey As Integer 'Cnt es un contador para acceder a la subclaves, empieza como 0 y se incrementa hasta que se llegue al final
        Dim StrBuff, StrKey As String
        RegOpenKey(hKey, strPath, TKey)
        Do
            StrBuff = New String(vbNullChar, 255) 'A pointer to a buffer that receives the name of the subkey, including the terminating null character
            'RegEnumKeyEx devuelve el valor ERROR_SUCCESS que equivale a 0 cuando tiene exito. Si falla devuelve un codigo de error
            If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, 0) <> 0 Then Exit Do
            Cnt = Cnt + 1
            StrKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
            sKeys.Add(StrKey)
        Loop
    End Sub

End Module
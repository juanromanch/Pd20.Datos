Imports System.Security.Cryptography

''' <summary>
''' Clase para encriptar y desencriptar cadenas de texto con TripleDES
''' </summary>
Public NotInheritable Class Encriptador
    Private TripleDes As New TripleDESCryptoServiceProvider

    Public Sub New(contra)
        TripleDes.Key = TruncateHash(contra, TripleDes.KeySize \ 8)
        TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
    End Sub

    ''' <summary>
    ''' Establece la clave de encriptación
    ''' </summary>
    Public WriteOnly Property Clave As String
        Set(value As String)
            TripleDes.Key = TruncateHash(value, TripleDes.KeySize \ 8)
            TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
        End Set

    End Property
    Private Function TruncateHash(
    ByVal _clave As String,
    ByVal longitud As Integer) As Byte()

        Dim sha1 As New SHA1CryptoServiceProvider

        ' Hash the clave.
        Dim keyBytes() As Byte =
            System.Text.Encoding.Unicode.GetBytes(_clave)
        Dim hash() As Byte = sha1.ComputeHash(keyBytes)

        ' Truncate or pad the hash.
        ReDim Preserve hash(longitud - 1)
        Return hash
    End Function

    ''' <summary>
    ''' Encripta una cadena de texto claro
    ''' </summary>
    ''' <param name="plaintext">Texto claro a encriptar</param>
    ''' <returns>Cadena de texto encriptado</returns>
    Public Function Encriptar(
    ByVal plaintext As String) As String

        ' Convert the plaintext string to a byte array.
        Dim textoplano() As Byte =
            System.Text.Encoding.Unicode.GetBytes(plaintext)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the encoder to write to the stream.
        Dim encStream As New CryptoStream(ms,
            TripleDes.CreateEncryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        encStream.Write(textoplano, 0, textoplano.Length)
        encStream.FlushFinalBlock()

        ' Convert the encrypted stream to a printable string.
        Return Convert.ToBase64String(ms.ToArray)
    End Function
    ''' <summary>
    ''' Desencripta una cadena de texto encriptado
    ''' </summary>
    ''' <param name="textoencriptado">Cadena de texto encriptad</param>
    ''' <returns>Texto desencriptado</returns>
    Public Function Desencriptar(
    ByVal textoencriptado As String) As String

        ' Convert the encrypted text string to a byte array.
        Dim bytesEncriptados() As Byte = Convert.FromBase64String(textoencriptado)

        ' Create the stream.
        Dim ms As New System.IO.MemoryStream
        ' Create the decoder to write to the stream.
        Dim decStream As New CryptoStream(ms,
            TripleDes.CreateDecryptor(),
            System.Security.Cryptography.CryptoStreamMode.Write)

        ' Use the crypto stream to write the byte array to the stream.
        decStream.Write(bytesEncriptados, 0, bytesEncriptados.Length)
        decStream.FlushFinalBlock()

        ' Convert the plaintext stream to a string.
        Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
    End Function
End Class

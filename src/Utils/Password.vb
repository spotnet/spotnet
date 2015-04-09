Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Friend Class cEncPass

    Private iv() As Byte
    Private key() As Byte
    Public Function Encrypt(ByVal plainText As String) As String

        If plainText Is Nothing Then Return Nothing

        ' Declare a UTF8Encoding object so we may use the GetByte 
        ' method to transform the plainText into a Byte array. 
        Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
        Dim inputInBytes() As Byte = utf8encoder.GetBytes(plainText)

        ' Create a new TripleDES service provider 
        Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

        ' The ICryptTransform interface uses the TripleDES 
        ' crypt provider along with encryption key and init vector 
        ' information 
        Dim cryptoTransform As ICryptoTransform =
        tdesProvider.CreateEncryptor(Me.key, Me.iv)

        ' All cryptographic functions need a stream to output the 
        ' encrypted information. Here we declare a memory stream 
        ' for this purpose. 
        Dim encryptedStream As MemoryStream = New MemoryStream()
        Dim cryptStream As CryptoStream = New CryptoStream(encryptedStream,
        cryptoTransform, CryptoStreamMode.Write)

        ' Write the encrypted information to the stream. Flush the information 
        ' when done to ensure everything is out of the buffer. 
        cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
        cryptStream.FlushFinalBlock()
        encryptedStream.Position = 0

        ' Read the stream back into a Byte array and return it to the calling 
        ' method. 
        Dim result(CInt(encryptedStream.Length - 1)) As Byte
        encryptedStream.Read(result, 0, CInt(encryptedStream.Length))
        cryptStream.Close()
        Return Convert.ToBase64String(result) & "=" ' force = teken, voor detectie OldDecrypt

    End Function

    Public Function Decrypt(ByVal inputInByts As String) As String

        If Len(Trim$(inputInByts)) = 0 Then Return ""

        inputInByts = inputInByts.Substring(0, inputInByts.Length - 1) ' Strip extra '='

        Try

            ' UTFEncoding is used to transform the decrypted Byte Array 
            ' information back into a string. 
            Dim utf8encoder As UTF8Encoding = New UTF8Encoding()
            Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider()

            Dim inputInBytes() As Byte = Convert.FromBase64String(inputInByts)

            ' As before we must provide the encryption/decryption key along with 
            ' the init vector. 
            Dim cryptoTransform As ICryptoTransform =
            tdesProvider.CreateDecryptor(Me.key, Me.iv)

            ' Provide a memory stream to decrypt information into 
            Dim decryptedStream As MemoryStream = New MemoryStream()
            Dim cryptStream As CryptoStream = New CryptoStream(decryptedStream,
            cryptoTransform, CryptoStreamMode.Write)
            cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
            cryptStream.FlushFinalBlock()
            decryptedStream.Position = 0

            ' Read the memory stream and convert it back into a string 
            Dim result(CInt(decryptedStream.Length - 1)) As Byte
            decryptedStream.Read(result, 0, CInt(decryptedStream.Length))
            cryptStream.Close()
            Dim myutf As UTF8Encoding = New UTF8Encoding()
            Return myutf.GetString(result)

        Catch ex As Exception

            Tools.Foutje(ex.Message)
            Return vbNullString
        End Try


    End Function

End Class

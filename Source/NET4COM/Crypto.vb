Imports System.Security.Cryptography
Imports System.IO
Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Crypto.InterfaceId)> _
Public Interface ICrypto
    Function DecryptTextFromFile(ByVal fileName As String, ByVal key As String) As String
    Function EncryptTextToFile(ByVal textToEncrypt As String, ByVal fileName As String) As String
End Interface

<ComVisible(True)> _
<Guid(Crypto.ClassId)> _
<ComDefaultInterface(GetType(ICrypto))> _
Public Class Crypto
    Implements ICrypto

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "f8ef4741-2dfa-4d56-bcbb-076451fdeb2f"
    Public Const InterfaceId As String = "540b7f2c-4fc8-4108-a38a-ae60cb39bd35"
    Public Const EventsId As String = "fcdbf7d6-1b4e-4202-8975-6eac61c1c223"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private RijndaelAlg As Rijndael = Rijndael.Create

    Protected Function DecryptTextFromFile(ByVal fileName As String, ByVal key As String) As String Implements ICrypto.DecryptTextFromFile
        ' Set the key to use to decrypt
        Dim byteKey As Byte()
        byteKey = Convert.FromBase64String(key)
        RijndaelAlg.Key = byteKey

        ' Create or open the specified file. 
        Dim fStream As FileStream = File.Open(fileName, FileMode.OpenOrCreate)

        ' Create a CryptoStream using the FileStream 
        ' and the passed key and initialization vector (IV).
        Dim cStream As New CryptoStream(fStream, _
                                        RijndaelAlg.CreateDecryptor(RijndaelAlg.Key, RijndaelAlg.IV), _
                                        CryptoStreamMode.Read)

        ' Create a StreamReader using the CryptoStream.
        Dim sReader As New StreamReader(cStream)

        ' Read the data from the stream 
        ' to decrypt it.
        Dim val As String = Nothing

        Try
            val = sReader.ReadLine()
        Finally
            ' Close the streams and
            ' close the file.
            sReader.Close()
            cStream.Close()
            fStream.Close()
        End Try

        ' Return the string. 
        Return val
    End Function


    Protected Function EncryptTextToFile(ByVal textToEncrypt As String, ByVal fileName As String) As String Implements ICrypto.EncryptTextToFile
        ' Encrypt text to a file using the file name, key, and IV.
        ' Create or open the specified file.
        Dim fStream As FileStream = File.Open(fileName, FileMode.OpenOrCreate)

        ' Create a CryptoStream using the FileStream 
        ' and the passed key and initialization vector (IV).
        Dim cStream As New CryptoStream(fStream, _
                                       RijndaelAlg.CreateEncryptor(RijndaelAlg.Key, RijndaelAlg.IV), _
                                       CryptoStreamMode.Write)

        ' Create a StreamWriter using the CryptoStream.
        Dim sWriter As New StreamWriter(cStream)

        Try
            ' Write the data to the stream 
            ' to encrypt it.
            sWriter.WriteLine(textToEncrypt)
        Finally
            ' Close the streams and
            ' close the file.
            sWriter.Close()
            cStream.Close()
            fStream.Close()
        End Try

        Return Convert.ToBase64String(RijndaelAlg.Key)
    End Function

End Class



Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Compression.InterfaceId)> _
Public Interface ICompression
    Sub Compress(ByVal source As String, ByVal destination As String)
    Sub Decompress(ByVal source As String, ByVal destination As String)
End Interface

<ComVisible(True)> _
<Guid(Compression.ClassId)> _
<ComDefaultInterface(GetType(ICompression))> _
Public Class Compression
    Implements ICompression

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a0de0132-94e5-4511-a45f-719ce0fccf6c"
    Public Const InterfaceId As String = "8b2a333e-b859-430a-86c3-d9d818b7d929"
    Public Const EventsId As String = "734e0af8-9956-4f1b-9592-631d2b3af3a3"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Sub Compress(ByVal source As String, ByVal destination As String) Implements ICompression.Compress
        Dim fileBytes As Byte() = My.Computer.FileSystem.ReadAllBytes(source)
        Dim fs As New FileStream(destination, FileMode.Create)
        ' Use the newly created file stream for the compressed data.
        Dim compressedzipStream As New GZipStream(fs, CompressionMode.Compress) ', True)
        compressedzipStream.Write(fileBytes, 0, fileBytes.Length)
        ' Close the stream.
        compressedzipStream.Close()
        fs.Close()
    End Sub

    Protected Sub Decompress(ByVal source As String, ByVal destination As String) Implements ICompression.Decompress
        ' Determine the uncompressed size of the file;
        ' we can't rely on the compressed size because
        ' it's not the true size
        Dim inputFile As FileStream = Nothing
        'Dim compressedZipStream As GZipStream = Nothing

        inputFile = New FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read)

        Dim compressedZipStream As GZipStream = New GZipStream(inputFile, CompressionMode.Decompress)

        Dim offset As Integer = 0
        Dim totalBytes As Integer = 0
        Dim smallBuffer(100) As Byte

        While True
            Dim bytesRead As Integer = compressedZipStream.Read(smallBuffer, 0, 100)
            If bytesRead = 0 Then
                Exit While
            End If
            offset += bytesRead
            totalBytes += bytesRead
        End While
        compressedZipStream.Close()

        ' Open and read the contents of the file now that
        ' we know the uncompressed size
        Dim fs As New FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read)

        ' Use the newly created file stream for the decompressed data.
        Dim decompressedzipStream As New GZipStream(fs, CompressionMode.Decompress)
        Dim buffer(totalBytes - 1) As Byte
        decompressedzipStream.Read(buffer, 0, totalBytes)
        My.Computer.FileSystem.WriteAllBytes(destination, buffer, False)
        ' Close the stream.
        decompressedzipStream.Close()
        fs.Close()

    End Sub

End Class



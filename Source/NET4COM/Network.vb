Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Network.InterfaceId)> _
Public Interface INetwork
    Function IsAvailable() As Boolean
    Function Ping(ByVal address As String) As Boolean
    Sub DownloadFile(ByVal address As String, ByVal destinationFileName As String)
    Sub UploadFile(ByVal address As String, ByVal destinationFileName As String)
    Sub SendEmail(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal messageText As String, ByVal SMTPServer As String)
End Interface

<ComVisible(True)> _
<Guid(Network.ClassId)> _
<ComDefaultInterface(GetType(INetwork))> _
Public Class Network
    Implements INetwork

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "a96dc74a-1c82-40af-86ee-688fa7efc9f3"
    Public Const InterfaceId As String = "ddf119da-7f74-4827-bf23-15cf7fa6d783"
    Public Const EventsId As String = "6460cbeb-e8b8-4005-8c87-3253490c27f6"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Function IsAvailable() As Boolean Implements INetwork.IsAvailable
        Return My.Computer.Network.IsAvailable
    End Function

    Protected Function Ping(ByVal address As String) As Boolean Implements INetwork.Ping
        Return My.Computer.Network.Ping(address)
    End Function

    Protected Sub DownloadFile(ByVal address As String, ByVal destinationFileName As String) Implements INetwork.DownloadFile
        My.Computer.Network.DownloadFile(address, destinationFileName)
    End Sub

    Protected Sub UploadFile(ByVal address As String, ByVal destinationFileName As String) Implements INetwork.UploadFile
        My.Computer.Network.UploadFile(address, destinationFileName, "", "")
    End Sub

    Protected Sub SendEmail(ByVal sender As String, ByVal recipient As String, ByVal subject As String, ByVal messageText As String, ByVal SMTPServer As String) Implements INetwork.SendEmail
        Using mailMessage As New MailMessage(sender, recipient, subject, messageText)
            Dim emailClient As New SmtpClient(SMTPServer)
            emailClient.Send(mailMessage)
        End Using
    End Sub
End Class


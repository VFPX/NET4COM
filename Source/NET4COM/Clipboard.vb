Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Clipboard.InterfaceId)> _
Public Interface IClipboard
    Sub Clear()
    Function ContainsAudio() As Boolean
    Function ContainsData(ByVal customDataFormat As String) As Boolean
    Function ContainsImage() As Boolean
    Function ContainsText() As Boolean
    Function ContainsFileDropList() As Boolean
    Function GetText() As String
    Sub SetText(ByVal text As String)
End Interface

<ComVisible(True)> _
<Guid(Clipboard.ClassId)> _
<ComDefaultInterface(GetType(IClipboard))> _
Public Class Clipboard
    Implements IClipboard

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "b9d550c0-e90c-47b0-a95d-f77bd9bacc8e"
    Public Const InterfaceId As String = "364b59e1-986c-48a7-b2e4-e9f56f30aa2f"
    Public Const EventsId As String = "afdfb454-00e7-4bcf-aab7-e6847e16ce0b"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Sub Clear() Implements IClipboard.Clear
        My.Computer.Clipboard.Clear()
    End Sub

    Protected Function ContainsAudio() As Boolean Implements IClipboard.ContainsAudio
        Return My.Computer.Clipboard.ContainsAudio()
    End Function

    Protected Function ContainsData(ByVal customDataFormat As String) As Boolean Implements IClipboard.ContainsData
        Return My.Computer.Clipboard.ContainsData(customDataFormat)
    End Function

    Protected Function ContainsImage() As Boolean Implements IClipboard.ContainsImage
        Return My.Computer.Clipboard.ContainsImage()
    End Function

    Protected Function ContainsText() As Boolean Implements IClipboard.ContainsText
        Return My.Computer.Clipboard.ContainsText()
    End Function

    Protected Function ContainsFileDropList() As Boolean Implements IClipboard.ContainsFileDropList
        Return My.Computer.Clipboard.ContainsFileDropList()
    End Function

    Protected Function GetText() As String Implements IClipboard.GetText
        Return My.Computer.Clipboard.GetText()
    End Function

    Protected Sub SetText(ByVal text As String) Implements IClipboard.SetText
        My.Computer.Clipboard.SetText(text)
    End Sub

End Class



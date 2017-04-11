Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Mouse.InterfaceId)> _
Public Interface IMouse
    Function ButtonsSwapped() As Boolean
    Function WheelExists() As Boolean
    Function WheelScrollLines() As Integer
End Interface

<ComVisible(True)> _
<Guid(Mouse.ClassId)> _
<ComDefaultInterface(GetType(IMouse))> _
Public Class Mouse
    Implements IMouse

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "9ad99cd3-e715-4388-be12-c36fb68b7e23"
    Public Const InterfaceId As String = "cfa8fe7b-1a63-4118-887e-8f6625def8b4"
    Public Const EventsId As String = "94d20fa6-bd8a-4f78-a7f8-734166be1f37"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Function ButtonsSwapped() As Boolean Implements IMouse.ButtonsSwapped
        Return My.Computer.Mouse.ButtonsSwapped()
    End Function

    Protected Function WheelExists() As Boolean Implements IMouse.WheelExists
        Return My.Computer.Mouse.WheelExists()
    End Function

    Protected Function WheelScrollLines() As Integer Implements IMouse.WheelScrollLines
        Return My.Computer.Mouse.WheelScrollLines()
    End Function

End Class



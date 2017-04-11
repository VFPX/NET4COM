Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Printer.InterfaceId)> _
Public Interface IPrinter
    Sub PrintImage(ByVal imageFile As String)
End Interface

<ComVisible(True)> _
<Guid(Printer.ClassId)> _
<ComDefaultInterface(GetType(IPrinter))> _
Public Class Printer
    Implements IPrinter

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "8cea2414-80a1-4cff-b492-80786fc72795"
    Public Const InterfaceId As String = "1fd5072e-b3ce-4906-9e39-bf4f1f49e8c1"
    Public Const EventsId As String = "facc3645-035f-418e-b892-64384a7757aa"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private imageFileName As String

    Protected Sub PrintImage(ByVal imageFile As String) Implements IPrinter.PrintImage
        Me.imageFileName = imageFile
        Dim pd As New PrintDocument()
        AddHandler pd.PrintPage, AddressOf Me.pd_PrintPage
        pd.Print()
    End Sub

    ' The PrintPage event is raised for each page to be printed.
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim g As Graphics = ev.Graphics
        g.DrawImage(Image.FromFile(Me.imageFileName), 5, 5)
    End Sub

End Class



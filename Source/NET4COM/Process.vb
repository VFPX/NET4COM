Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Process.InterfaceId)> _
Public Interface IProcess
    Sub StopApp(ByVal processName As String)
    Sub StartApp(ByVal fileName As String)
End Interface

<ComVisible(True)> _
<Guid(Process.ClassId)> _
<ComDefaultInterface(GetType(IProcess))> _
Public Class Process
    Implements IProcess

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "1285e378-7313-447f-bb98-d8a5931725f2"
    Public Const InterfaceId As String = "0a02562d-e94b-435c-b26b-8c05deeb9978"
    Public Const EventsId As String = "5b95835a-76e2-449e-b8fe-d9da30aa889a"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Sub StopApp(ByVal processName As String) Implements IProcess.StopApp
        Dim apps() As System.Diagnostics.Process = _
            System.Diagnostics.Process.GetProcessesByName(processName)
        For Each app As System.Diagnostics.Process In apps
            app.CloseMainWindow()
        Next
    End Sub

    Protected Sub StartApp(ByVal fileName As String) Implements IProcess.StartApp
        System.Diagnostics.Process.Start(fileName)
    End Sub

End Class



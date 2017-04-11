Imports System.ServiceProcess
Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(WindowsServices.InterfaceId)> _
Public Interface IWindowsServices
    Sub StartService(ByVal service As String)
    Sub StopService(ByVal service As String)
    Sub PauseService(ByVal service As String)
    Sub ContinueService(ByVal service As String)
End Interface

<ComVisible(True)> _
<Guid(WindowsServices.ClassId)> _
<ComDefaultInterface(GetType(IWindowsServices))> _
Public Class WindowsServices
    Implements IWindowsServices

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "29d80b04-50e3-4548-9aca-87f1877e055c"
    Public Const InterfaceId As String = "832d787f-0fa9-4d6b-bab2-bf78a9ed5413"
    Public Const EventsId As String = "75071fef-45f8-4067-9513-454a3f57ff03"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Sub StartService(ByVal service As String) Implements IWindowsServices.StartService
        Using controller As New ServiceController(service)
            controller.Start()
        End Using
    End Sub

    Sub StopService(ByVal service As String) Implements IWindowsServices.StopService
        Using controller As New ServiceController(service)
            controller.Stop()
        End Using
    End Sub

    Sub PauseService(ByVal service As String) Implements IWindowsServices.PauseService
        Using controller As New ServiceController(service)
            controller.Pause()
        End Using
    End Sub

    Sub ContinueService(ByVal service As String) Implements IWindowsServices.ContinueService
        Using controller As New ServiceController(service)
            controller.Continue()
        End Using
    End Sub

End Class



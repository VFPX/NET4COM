Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(PowerStatus.InterfaceId)> _
Public Interface IPowerStatus
    ReadOnly Property BatteryLifePercent() As Single
    ReadOnly Property BatteryFullLifetime() As Integer
    ReadOnly Property BatteryLifeRemaining() As Integer
    ReadOnly Property BatteryChargeStatus() As Integer
    ReadOnly Property PowerLineStatus() As Integer
End Interface

<ComVisible(True)> _
<Guid(PowerStatus.ClassId)> _
<ComDefaultInterface(GetType(IPowerStatus))> _
Public Class PowerStatus
    Implements IPowerStatus

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "561aa7cc-a5d6-4584-96c7-1215c96139c3"
    Public Const InterfaceId As String = "ddf3bfa0-fbee-44bb-b0aa-52530b6694bb"
    Public Const EventsId As String = "3beff989-7ff3-4ed6-bc7c-4fa1369c9597"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Dim power As System.Windows.Forms.PowerStatus = System.Windows.Forms.SystemInformation.PowerStatus

    Protected ReadOnly Property BatteryLifePercent() As Single Implements IPowerStatus.BatteryLifePercent
        Get
            Return power.BatteryLifePercent
        End Get
    End Property

    Protected ReadOnly Property BatteryFullLifetime() As Integer Implements IPowerStatus.BatteryFullLifetime
        Get
            Return power.BatteryFullLifetime
        End Get
    End Property

    Protected ReadOnly Property BatteryLifeRemaining() As Integer Implements IPowerStatus.BatteryLifeRemaining
        Get
            Return power.BatteryLifeRemaining
        End Get
    End Property

    Protected ReadOnly Property BatteryChargeStatus() As Integer Implements IPowerStatus.BatteryChargeStatus
        Get
            Return power.BatteryChargeStatus
        End Get
    End Property

    Protected ReadOnly Property PowerLineStatus() As Integer Implements IPowerStatus.PowerLineStatus
        Get
            Return power.PowerLineStatus
        End Get
    End Property

End Class



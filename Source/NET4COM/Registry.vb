Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Registry.InterfaceId)> _
Public Interface IRegistry
    Function GetValue(ByVal keyName As String, ByVal valueName As String) As String
    Sub SetValue(ByVal keyName As String, ByVal valueName As String, ByVal value As String)
    Sub CreateKey(ByVal hive As String, ByVal subKey As String)
    Sub DeleteKey(ByVal hive As String, ByVal subKey As String)
End Interface

    <ComVisible(True)> _
    <Guid(Registry.ClassId)> _
    <ComDefaultInterface(GetType(IRegistry))> _
    Public Class Registry
        Implements IRegistry

#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "8d4971ce-68dd-4879-8bf5-682bfaf9cdc1"
        Public Const InterfaceId As String = "ca6e4a85-5354-4df5-85cd-9d3bdecb8798"
        Public Const EventsId As String = "a7f1267b-a2c6-44fd-9fac-b13ffcdd15eb"
#End Region

        ' A creatable COM class must have a Public Sub New() 
        ' with no parameters, otherwise, the class will not be 
        ' registered in the COM registry and cannot be created 
        ' via CreateObject.
        Public Sub New()
            MyBase.New()
        End Sub

        Protected Function GetValue(ByVal keyName As String, ByVal valueName As String) As String Implements IRegistry.GetValue
            Return My.Computer.Registry.GetValue(keyName, valueName, Nothing).ToString()
        End Function

        Protected Sub SetValue(ByVal keyName As String, ByVal valueName As String, ByVal value As String) Implements IRegistry.SetValue
            My.Computer.Registry.SetValue(keyName, valueName, value)
        End Sub

        Protected Sub CreateKey(ByVal hive As String, ByVal subKey As String) Implements IRegistry.CreateKey
        Select Case hive.ToUpper(Globalization.CultureInfo.CurrentCulture)
            Case My.Computer.Registry.ClassesRoot.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.ClassesRoot.CreateSubKey(subKey)
            Case My.Computer.Registry.CurrentConfig.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.CurrentConfig.CreateSubKey(subKey)
            Case My.Computer.Registry.CurrentUser.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.CurrentUser.CreateSubKey(subKey)
            Case My.Computer.Registry.DynData.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.DynData.CreateSubKey(subKey)
            Case My.Computer.Registry.LocalMachine.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.LocalMachine.CreateSubKey(subKey)
            Case My.Computer.Registry.PerformanceData.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.PerformanceData.CreateSubKey(subKey)
            Case My.Computer.Registry.Users.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.Users.CreateSubKey(subKey)
        End Select
        End Sub

    Protected Sub DeleteKey(ByVal hive As String, ByVal subKey As String) Implements IRegistry.DeleteKey
        Select Case hive.ToUpper(Globalization.CultureInfo.CurrentCulture)
            Case My.Computer.Registry.ClassesRoot.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.ClassesRoot.DeleteSubKey(subKey)
            Case My.Computer.Registry.CurrentConfig.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.CurrentConfig.DeleteSubKey(subKey)
            Case My.Computer.Registry.CurrentUser.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.CurrentUser.DeleteSubKey(subKey)
            Case My.Computer.Registry.DynData.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.DynData.DeleteSubKey(subKey)
            Case My.Computer.Registry.LocalMachine.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.LocalMachine.DeleteSubKey(subKey)
            Case My.Computer.Registry.PerformanceData.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.PerformanceData.DeleteSubKey(subKey)
            Case My.Computer.Registry.Users.Name.ToUpper(Globalization.CultureInfo.CurrentCulture)
                My.Computer.Registry.Users.DeleteSubKey(subKey)
        End Select
    End Sub

End Class



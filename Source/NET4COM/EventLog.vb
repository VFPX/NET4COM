Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(EventLog.InterfaceId)> _
Public Interface IEventLog
    Function LogEntries(ByVal logName As String) As String()
    Function LogEntriesForSource(ByVal logName As String, ByVal sourceName As String) As String()
End Interface

<ComVisible(True)> _
<Guid(EventLog.ClassId)> _
<ComDefaultInterface(GetType(IEventLog))> _
Public Class EventLog
    Implements IEventLog

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4eda587b-542a-491d-8dc1-6fd6ad8c417e"
    Public Const InterfaceId As String = "f080a257-0d40-4086-9617-38f23cc96083"
    Public Const EventsId As String = "f1cc35de-5317-4866-aa5e-6c0dbd3b1616"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Function LogEntries(ByVal logName As String) As String() Implements IEventLog.LogEntries

        Using log As New System.Diagnostics.EventLog(logName)
            Dim arrayString As ArrayList = New ArrayList(log.Entries.Count)
            For Each entry As EventLogEntry In log.Entries
                arrayString.Add(entry.Message)
            Next
            Return CType(arrayString.ToArray(GetType(String)), String())
        End Using
    End Function

    Protected Function LogEntriesForSource(ByVal logName As String, ByVal sourceName As String) As String() Implements IEventLog.LogEntriesForSource

        Using log As New System.Diagnostics.EventLog(logName, ".", sourceName)
            Dim arrayString As ArrayList = New ArrayList(log.Entries.Count)
            For Each entry As EventLogEntry In log.Entries
                arrayString.Add(entry.Message)
            Next
            Return CType(arrayString.ToArray(GetType(String)), String())
        End Using
    End Function

End Class



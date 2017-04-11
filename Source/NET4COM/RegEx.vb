Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(RegEx.InterfaceId)> _
Public Interface IRegEx
    Function Replace(ByVal input As String, ByVal pattern As String, ByVal replacement As String) As String
    Function IsMatch(ByVal input As String, ByVal pattern As String) As Boolean
    Function Matches(ByVal input As String, ByVal pattern As String) As String()
    Function Split(ByVal input As String, ByVal pattern As String) As String()
End Interface

<ComVisible(True)> _
<Guid(RegEx.ClassId)> _
<ComDefaultInterface(GetType(IRegEx))> _
Public Class RegEx
    Implements IRegEx

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "6a7c2014-997a-4739-8f1e-e8497dd48917"
    Public Const InterfaceId As String = "7bd77ab3-9e77-485e-a67f-af41740d45a8"
    Public Const EventsId As String = "b0cd3fea-2e5b-4849-a0da-79fa67ec5615"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Function Replace(ByVal input As String, ByVal pattern As String, ByVal replacement As String) As String Implements IRegEx.Replace
        Return System.Text.RegularExpressions.Regex.Replace(input, pattern, replacement)
    End Function

    Protected Function IsMatch(ByVal input As String, ByVal pattern As String) As Boolean Implements IRegEx.IsMatch
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)
    End Function

    Protected Function Matches(ByVal input As String, ByVal pattern As String) As String() Implements IRegEx.Matches
        Dim mc As System.Text.RegularExpressions.MatchCollection
        mc = System.Text.RegularExpressions.Regex.Matches(input, pattern)
        Dim x As Integer
        Dim numItems As Integer = mc.Count
        Dim arrayString As ArrayList = New ArrayList(numItems)
        For x = 0 To (mc.Count - 1)
            arrayString.Add(mc(x).Groups(0).Value)
        Next
        Return CType(arrayString.ToArray(GetType(String)), String())
    End Function

    Protected Function Split(ByVal input As String, ByVal pattern As String) As String() Implements IRegEx.Split
        Return System.Text.RegularExpressions.Regex.Split(input, pattern)
    End Function

End Class



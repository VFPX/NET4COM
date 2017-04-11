Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(SpecialDirectories.InterfaceId)> _
Public Interface ISpecialDirectories
    Function GetMyDocuments() As String
    Function GetDesktop() As String
    Function GetMyMusic() As String
    Function GetMyPictures() As String
    Function GetProgramFiles() As String
    Function GetPrograms() As String
    Function GetTempPath() As String
    Function GetAllUsersApplicationData() As String
    Function CurrentUserApplicationData() As String
End Interface

<ComVisible(True)> _
<Guid(SpecialDirectories.ClassId)> _
<ComDefaultInterface(GetType(ISpecialDirectories))> _
Public Class SpecialDirectories
    Implements ISpecialDirectories

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "9e25ef29-bd19-43ca-a6d4-a83c020ca18c"
    Public Const InterfaceId As String = "1f84fa84-25f0-4e7b-8164-554c2f2121aa"
    Public Const EventsId As String = "92f0a14e-b4ca-4a32-8f61-8a130024d142"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function CurrentUserApplicationData() As String Implements ISpecialDirectories.CurrentUserApplicationData
        Return My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData
    End Function

    Public Function GetAllUsersApplicationData() As String Implements ISpecialDirectories.GetAllUsersApplicationData
        Return My.Computer.FileSystem.SpecialDirectories.AllUsersApplicationData
    End Function

    Public Function GetDesktop() As String Implements ISpecialDirectories.GetDesktop
        Return My.Computer.FileSystem.SpecialDirectories.Desktop
    End Function

    Public Function GetMyDocuments() As String Implements ISpecialDirectories.GetMyDocuments
        Return My.Computer.FileSystem.SpecialDirectories.MyDocuments
    End Function

    Public Function GetMyMusic() As String Implements ISpecialDirectories.GetMyMusic
        Return My.Computer.FileSystem.SpecialDirectories.MyMusic
    End Function

    Public Function GetMyPictures() As String Implements ISpecialDirectories.GetMyPictures
        Return My.Computer.FileSystem.SpecialDirectories.MyPictures
    End Function

    Public Function GetProgramFiles() As String Implements ISpecialDirectories.GetProgramFiles
        Return My.Computer.FileSystem.SpecialDirectories.ProgramFiles
    End Function

    Public Function GetPrograms() As String Implements ISpecialDirectories.GetPrograms
        Return My.Computer.FileSystem.SpecialDirectories.Programs
    End Function

    Public Function GetTempPath() As String Implements ISpecialDirectories.GetTempPath
        Return My.Computer.FileSystem.SpecialDirectories.Temp
    End Function
End Class



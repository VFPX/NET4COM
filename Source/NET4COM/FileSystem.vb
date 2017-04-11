Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(FileSystem.InterfaceId)> _
Public Interface IFileSystem
    Sub CopyDirectory(ByVal source As String, ByVal destination As String)
    Sub CopyFile(ByVal source As String, ByVal destination As String)
    Sub MoveDirectory(ByVal source As String, ByVal destination As String)
    Sub MoveFile(ByVal source As String, ByVal destination As String)
    Function GetVolumeLabel(ByVal drive As String) As String
    Function GetTotalFreeDriveSpace(ByVal drive As String) As Long
    Function GetTotalDriveSpace(ByVal drive As String) As Long
End Interface

<ComVisible(True)> _
<Guid(FileSystem.ClassId)> _
<ComDefaultInterface(GetType(IFileSystem))> _
Public Class FileSystem
    Implements IFileSystem

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "7885154a-63c8-4cb8-b36a-ea30c16c1d28"
    Public Const InterfaceId As String = "54696944-982a-4590-8f3e-4770076ed491"
    Public Const EventsId As String = "e3c5faed-0f56-4b0f-b49e-76548a026009"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Sub CopyDirectory(ByVal source As String, ByVal destination As String) Implements IFileSystem.CopyDirectory
        My.Computer.FileSystem.CopyDirectory(source, destination)
    End Sub

    Protected Sub CopyFile(ByVal source As String, ByVal destination As String) Implements IFileSystem.CopyFile
        My.Computer.FileSystem.CopyFile(source, destination)
    End Sub

    Protected Function GetTotalDriveSpace(ByVal drive As String) As Long Implements IFileSystem.GetTotalDriveSpace
        Return My.Computer.FileSystem.GetDriveInfo(drive).TotalSize
    End Function

    Protected Function GetTotalFreeDriveSpace(ByVal drive As String) As Long Implements IFileSystem.GetTotalFreeDriveSpace
        Return My.Computer.FileSystem.GetDriveInfo(drive).TotalFreeSpace
    End Function

    Protected Function GetVolumeLabel(ByVal drive As String) As String Implements IFileSystem.GetVolumeLabel
        Return My.Computer.FileSystem.GetDriveInfo(drive).VolumeLabel
    End Function

    Protected Sub MoveDirectory(ByVal source As String, ByVal destination As String) Implements IFileSystem.MoveDirectory
        My.Computer.FileSystem.MoveDirectory(source, destination)
    End Sub

    Protected Sub MoveFile(ByVal source As String, ByVal destination As String) Implements IFileSystem.MoveFile
        My.Computer.FileSystem.MoveFile(source, destination)
    End Sub
End Class



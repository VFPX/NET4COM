Imports System.Runtime.InteropServices

<ComVisible(True)> _
<Guid(Audio.InterfaceId)> _
Public Interface IAudio
    Sub Play(ByVal soundFile As String, ByVal playMode As Integer)
    Sub StopAudio()
    Sub PlaySystemSound(ByVal systemSound As Integer)
End Interface

<ComVisible(True)> _
<Guid(Audio.ClassId)> _
<ComDefaultInterface(GetType(IAudio))> _
Public Class Audio
    Implements IAudio

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "fa1154fb-3956-4c92-bba0-ea7f97aa1859"
    Public Const InterfaceId As String = "b16708b5-7d19-4296-8809-61709a7014e3"
    Public Const EventsId As String = "024c5e2d-4b43-4f5b-9481-bd7585b6b021"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Protected Sub Play(ByVal soundFile As String, ByVal playMode As Integer) Implements IAudio.Play
        Select Case playMode
            Case 0
                My.Computer.Audio.Play(soundFile, AudioPlayMode.WaitToComplete)
            Case 1
                My.Computer.Audio.Play(soundFile, AudioPlayMode.Background)
            Case 2
                My.Computer.Audio.Play(soundFile, AudioPlayMode.BackgroundLoop)
        End Select
    End Sub 'Play

    Protected Sub StopAudio() Implements IAudio.StopAudio
        My.Computer.Audio.Stop()
    End Sub

    Protected Sub PlaySystemSound(ByVal systemSound As Integer) Implements IAudio.PlaySystemSound
        Select Case systemSound
            Case 0
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Asterisk)
            Case 1
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Beep)
            Case 2
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Exclamation)
            Case 3
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Hand)
            Case 4
                My.Computer.Audio.PlaySystemSound(System.Media.SystemSounds.Question)
        End Select
    End Sub

End Class



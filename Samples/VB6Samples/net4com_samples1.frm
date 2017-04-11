VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NET4COM Samples"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNetworkStatus 
      Caption         =   "Network Status"
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdRegExFormat 
      Caption         =   "Format 4255551234"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton cmdStopApp 
      Caption         =   "Stop Notepad"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartApp 
      Caption         =   "Start Notepad"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNetworkStatus_Click()
Dim oNetwork As NET4COM.Network
Set oNetwork = CreateObject("NET4COM.Network")
If oNetwork.IsAvailable() Then
    MsgBox "Network is available."
Else
    MsgBox "Network is not available."
End If
End Sub

Private Sub cmdRegExFormat_Click()
Dim oRegEx As NET4COM.RegEx
Set oRegEx = CreateObject("NET4COM.RegEx")
MsgBox (oRegEx.Replace("4255551234", "(\d{3})(\d{3})(\d{4})", "($1) $2-$3"))
End Sub

Private Sub cmdStartApp_Click(Index As Integer)
Dim oProcess As NET4COM.Process
Set oProcess = New NET4COM.Process
oProcess.StartApp ("Notepad.exe")
End Sub

Private Sub cmdStopApp_Click(Index As Integer)
Dim oProcess As NET4COM.Process
Set oProcess = New NET4COM.Process
oProcess.StopApp ("notepad")
End Sub


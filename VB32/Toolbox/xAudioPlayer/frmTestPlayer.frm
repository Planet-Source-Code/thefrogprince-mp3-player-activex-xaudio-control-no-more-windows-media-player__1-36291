VERSION 5.00
Object = "{D6059D49-FD16-4C5A-9F0F-90D319318521}#1.1#0"; "XAudioPlayer.ocx"
Begin VB.Form frmTestPlayer 
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin XAudioPlayer.ctlXaudioPlayer ctlXaudioPlayer1 
      Height          =   615
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
   End
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Text            =   "C:\Temp\AudioGalaxy\celine dion - A New Day Has Come.mp3"
      Top             =   660
      Width           =   4605
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Test Stop"
      Height          =   435
      Left            =   1560
      TabIndex        =   1
      Top             =   1260
      Width           =   1395
   End
   Begin VB.CommandButton btnTest 
      Caption         =   "Test Start"
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   1395
   End
End
Attribute VB_Name = "frmTestPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnStop_Click()
    Me.ctlXaudioPlayer1.StopPlayer
    
End Sub

Private Sub btnTest_Click()
    
    Me.ctlXaudioPlayer1.Play Me.txtFile.Text
    
End Sub

Private Sub ctlXaudioPlayer1_DebugError(tErr As XAudioPlayer.XA_DebugInfo)
'    MsgBox "Debug error." & vbCrLf & vbCrLf & tErr.mMessage
End Sub

Private Sub ctlXaudioPlayer1_Error(tErr As XAudioPlayer.XA_ErrorInfo)
    MsgBox "Player error." & vbCrLf & vbCrLf & tErr.mMessage
End Sub

Private Sub Form_Terminate()
    MsgBox "Begin form terminate"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Begin unload"
    
End Sub





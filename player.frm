VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "mswebdvd.dll"
Begin VB.Form player 
   Caption         =   "MyDVD Personal Edition"
   ClientHeight    =   4425
   ClientLeft      =   1845
   ClientTop       =   450
   ClientWidth     =   5535
   Icon            =   "player.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   1080
      Top             =   3960
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   720
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   3960
   End
   Begin MSWEBDVDLibCtl.MSWebDVD MSWebDVD1 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _cx             =   9763
      _cy             =   7858
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   983055
      WindowlessActivation=   0   'False
   End
End
Attribute VB_Name = "player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
MSWebDVD1.Width = player.Width
MSWebDVD1.Height = player.Height - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "To show this windows again, please press the button ''Video'' at the controller window."
End Sub

Private Sub Timer1_Timer()
Me.Caption = "MyDvD Personal Edition - Playing /"
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Me.Caption = "MyDvD Personal Edition - Playing |"
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Me.Caption = "MyDvD Personal Edition - Playing \"
Timer3.Enabled = False
Timer4.Enabled = True
End Sub

Private Sub Timer4_Timer()
Me.Caption = "MyDvD Personal Edition - Playing -"
Timer4.Enabled = False
Timer1.Enabled = True
End Sub

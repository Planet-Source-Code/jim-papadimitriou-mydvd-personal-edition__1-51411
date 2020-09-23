VERSION 5.00
Begin VB.Form Controller 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "MyDVD Personal Edition"
   ClientHeight    =   1710
   ClientLeft      =   1110
   ClientTop       =   5040
   ClientWidth     =   7320
   FillStyle       =   5  'Downward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   7320
   Begin VB.CommandButton Command45 
      Caption         =   "About"
      Height          =   255
      Left            =   5640
      TabIndex        =   62
      Top             =   240
      Width           =   735
   End
   Begin VB.CheckBox Fulls 
      Caption         =   "Turn Fullscreen on"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox SubtitlesOnOff 
      Caption         =   "Subtitles On/Off"
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox time 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   51
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2760
      Top             =   1320
   End
   Begin VB.Frame Speeds 
      Caption         =   "Select Speed"
      Height          =   855
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   7335
      Begin VB.CommandButton Command42 
         Caption         =   "-20"
         Height          =   255
         Left            =   6840
         TabIndex        =   46
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command41 
         Caption         =   "20"
         Height          =   255
         Left            =   6840
         TabIndex        =   45
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command40 
         Caption         =   "-19"
         Height          =   255
         Left            =   6480
         TabIndex        =   44
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command39 
         Caption         =   "19"
         Height          =   255
         Left            =   6480
         TabIndex        =   43
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command38 
         Caption         =   "-18"
         Height          =   255
         Left            =   6120
         TabIndex        =   42
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command37 
         Caption         =   "18"
         Height          =   255
         Left            =   6120
         TabIndex        =   41
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command36 
         Caption         =   "-17"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command35 
         Caption         =   "17"
         Height          =   255
         Left            =   5760
         TabIndex        =   39
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command34 
         Caption         =   "-16"
         Height          =   255
         Left            =   5400
         TabIndex        =   38
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command33 
         Caption         =   "16"
         Height          =   255
         Left            =   5400
         TabIndex        =   37
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command32 
         Caption         =   "-15"
         Height          =   255
         Left            =   5040
         TabIndex        =   36
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command31 
         Caption         =   "15"
         Height          =   255
         Left            =   5040
         TabIndex        =   35
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command30 
         Caption         =   "-14"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command29 
         Caption         =   "14"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command28 
         Caption         =   "-13"
         Height          =   255
         Left            =   4320
         TabIndex        =   32
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command27 
         Caption         =   "13"
         Height          =   255
         Left            =   4320
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command26 
         Caption         =   "-12"
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command25 
         Caption         =   "12"
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command24 
         Caption         =   "-11"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command23 
         Caption         =   "11"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command22 
         Caption         =   "-10"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command21 
         Caption         =   "10"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command20 
         Caption         =   "-9"
         Height          =   255
         Left            =   2880
         TabIndex        =   24
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command19 
         Caption         =   "9"
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command18 
         Caption         =   "-8"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command17 
         Caption         =   "8"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command16 
         Caption         =   "-7"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "7"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Caption         =   "-6"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Caption         =   "6"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-5"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "5"
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "-4"
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "4"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-3"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "3"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-2"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "2"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-1"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "1"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Closedvdprog 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6720
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CheckBox UnMute 
      Caption         =   "Mute"
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Dir"
      Height          =   255
      Left            =   1800
      TabIndex        =   47
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Titles"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton gotoRoot 
      Caption         =   "Root"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Stop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.CheckBox PlayPause 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.CheckBox VideoWin 
      Caption         =   "Video"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bookmark"
      Height          =   495
      Left            =   2640
      TabIndex        =   58
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton Command44 
         Caption         =   "Restore"
         Height          =   255
         Left            =   600
         TabIndex        =   60
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Save"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Total DvD PlayTime"
      Height          =   255
      Left            =   3600
      TabIndex        =   56
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Played"
      Height          =   255
      Left            =   720
      TabIndex        =   57
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Available Volumes:"
      Height          =   375
      Left            =   2520
      TabIndex        =   55
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Current Chapter"
      Height          =   375
      Left            =   1560
      TabIndex        =   54
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3240
      TabIndex        =   50
      Top             =   600
      Width           =   375
   End
   Begin VB.Label CChap 
      Height          =   255
      Left            =   2160
      TabIndex        =   49
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "0:00:00:00"
      Height          =   255
      Left            =   600
      TabIndex        =   48
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Check1_Click()

End Sub

Private Sub Closedvdprog_Click()
End
End Sub


Private Sub Command1_Click()
 player.MSWebDVD1.ShowMenu "2"
End Sub

Private Sub Command10_Click()
player.MSWebDVD1.PlayBackwards "4"
End Sub

Private Sub Command11_Click()
player.MSWebDVD1.PlayForwards "5"
End Sub

Private Sub Command12_Click()
player.MSWebDVD1.PlayBackwards "5"
End Sub

Private Sub Command13_Click()
player.MSWebDVD1.PlayForwards "6"
End Sub

Private Sub Command14_Click()
player.MSWebDVD1.PlayBackwards "6"
End Sub

Private Sub Command15_Click()
player.MSWebDVD1.PlayForwards "7"
End Sub

Private Sub Command16_Click()
player.MSWebDVD1.PlayBackwards "7"
End Sub

Private Sub Command17_Click()
player.MSWebDVD1.PlayForwards "8"
End Sub

Private Sub Command18_Click()
player.MSWebDVD1.PlayBackwards "8"
End Sub

Private Sub Command19_Click()
player.MSWebDVD1.PlayForwards "9"
End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Command20_Click()
player.MSWebDVD1.PlayBackwards "9"
End Sub

Private Sub Command21_Click()
player.MSWebDVD1.PlayForwards "10"
End Sub

Private Sub Command22_Click()
player.MSWebDVD1.PlayBackwards "10"
End Sub

Private Sub Command23_Click()
player.MSWebDVD1.PlayForwards "11"
End Sub

Private Sub Command24_Click()
player.MSWebDVD1.PlayBackwards "11"
End Sub

Private Sub Command25_Click()
player.MSWebDVD1.PlayForwards "12"
End Sub

Private Sub Command26_Click()
player.MSWebDVD1.PlayBackwards "12"
End Sub

Private Sub Command27_Click()
player.MSWebDVD1.PlayForwards "13"
End Sub

Private Sub Command28_Click()
player.MSWebDVD1.PlayBackwards "13"
End Sub

Private Sub Command29_Click()
player.MSWebDVD1.PlayForwards "14"
End Sub

Private Sub Command3_Click()
player.MSWebDVD1.PlayForwards "1"
End Sub

Private Sub Command30_Click()
player.MSWebDVD1.PlayBackwards "14"
End Sub

Private Sub Command31_Click()
player.MSWebDVD1.PlayForwards "15"
End Sub

Private Sub Command32_Click()
player.MSWebDVD1.PlayBackwards "15"
End Sub

Private Sub Command33_Click()
player.MSWebDVD1.PlayForwards "16"
End Sub

Private Sub Command34_Click()
player.MSWebDVD1.PlayBackwards "16"
End Sub

Private Sub Command35_Click()
player.MSWebDVD1.PlayForwards "17"
End Sub

Private Sub Command36_Click()
player.MSWebDVD1.PlayBackwards "17"
End Sub

Private Sub Command37_Click()
player.MSWebDVD1.PlayForwards "18"
End Sub

Private Sub Command38_Click()
player.MSWebDVD1.PlayBackwards "18"
End Sub

Private Sub Command39_Click()
player.MSWebDVD1.PlayForwards "19"
End Sub

Private Sub Command4_Click()
player.MSWebDVD1.PlayBackwards "1"
End Sub

Private Sub Command40_Click()
player.MSWebDVD1.PlayBackwards "19"
End Sub

Private Sub Command41_Click()
player.MSWebDVD1.PlayForwards "20"
End Sub

Private Sub Command42_Click()
player.MSWebDVD1.PlayBackwards "20"
End Sub

Private Sub Command43_Click()
On Error Resume Next
player.MSWebDVD1.SaveBookmark
End Sub

Private Sub Command44_Click()
On Error Resume Next
player.MSWebDVD1.RestoreBookmark
End Sub

Private Sub Command45_Click()
frmAbout.Show
End Sub

Private Sub Command5_Click()
player.MSWebDVD1.PlayForwards "2"
End Sub

Private Sub Command6_Click()
player.MSWebDVD1.PlayBackwards "2"
End Sub

Private Sub Command7_Click()
player.MSWebDVD1.PlayForwards "3"
End Sub

Private Sub Command8_Click()
player.MSWebDVD1.PlayBackwards "3"
End Sub

Private Sub Command9_Click()
player.MSWebDVD1.PlayForwards "4"
End Sub

Private Sub Form_Load()
On Error Resume Next
player.Show
Form1.dir.Text = GetSetting("MyDVD", "Settings", "DVDDir")
If Form1.dir.Text = "" Then
   MsgBox "You will need to set the directory of DVD. There are instructions on the popup"
   Form1.Show
   PlayPause.Visible = False
   Speeds.Visible = False
Else
   player.MSWebDVD1.DVDDirectory = Form1.dir.Text
   PlayPause.Visible = True
   Speeds.Visible = True
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Fulls_Click()
If Fulls.Value = 0 Then
   player.MSWebDVD1.FullScreenMode = False
   Fulls.Caption = "Turn fullscreen on"
Else
    player.MSWebDVD1.FullScreenMode = True
    Fulls.Caption = "Turn fullscreen off"
    End If
End Sub

Private Sub gotoRoot_Click()
On Error Resume Next
 player.MSWebDVD1.ShowMenu "3"
End Sub

Private Sub PlayPause_Click()
On Error Resume Next
If PlayPause.Value = 0 Then
   PlayPause.Caption = "Play"
   player.MSWebDVD1.Pause
   Label4.Caption = dvd.TotalTitleTime
   'Timer1.Enabled = False
Else
    PlayPause.Caption = "Pause"
    player.MSWebDVD1.Play
    'Timer1.Enabled = True
End If
End Sub

Private Sub Speed_Change()

End If
End Sub

Private Sub Stop_Click()
On Error Resume Next
player.MSWebDVD1.Stop
PlayPause.Value = 0
End Sub

Private Sub SubtitlesOnOff_Click()
On Error GoTo Erro
If SubtitlesOnOff.Value = 0 Then
   player.MSWebDVD1.SubpictureOn = False
Else
    player.MSWebDVD1.SubpictureOn = True
End If
Erro:
   MsgBox "There is no disc in the drive or this operation is prohibited!"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label3.Caption = player.MSWebDVD1.CurrentTime
CChap.Caption = player.MSWebDVD1.CurrentChapter
Label6.Caption = player.MSWebDVD1.CurrentTitle
time.Text = player.MSWebDVD1.TotalTitleTime
End Sub

Private Sub UnMute_Click()
If UnMute.Value = 0 Then
    player.MSWebDVD1.Mute = False
    UnMute.Caption = "Mute"
Else
    player.MSWebDVD1.Mute = True
    UnMute.Caption = "unMute"
End If
End Sub

Private Sub VideoWin_Click()
If VideoWin.Value = 1 Then
    player.Show
Else
    player.Hide
End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Set Directory of Playing"
   ClientHeight    =   1410
   ClientLeft      =   1920
   ClientTop       =   3315
   ClientWidth     =   5730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SET"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox dir 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":000C
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
player.MSWebDVD1.DVDDirectory = dir.Text
SaveSetting "MyDVD", "Settings", "DVDDir", dir.Text
MsgBox "The program will now close because of changes to take effect. Please start it again."
Me.Hide
End
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

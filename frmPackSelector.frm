VERSION 5.00
Begin VB.Form frmPackSelector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select a Level Pack to load"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   720
      Top             =   3240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   2610
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   1230
      TabIndex        =   1
      Top             =   2610
      Width           =   1035
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   60
      Pattern         =   "*.pak"
      TabIndex        =   0
      Top             =   60
      Width           =   2265
   End
End
Attribute VB_Name = "frmPackSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.ClearScore
LevelPack = File1.FileName
frmLevels.Show
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub File1_DblClick()
If File1.ListIndex <> -1 Then Command1_Click
End Sub

Private Sub Form_Load()
frmMain.Enabled = False
File1.Path = App.Path
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
End Sub

Private Sub Timer1_Timer()
If File1.ListIndex = -1 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If

End Sub

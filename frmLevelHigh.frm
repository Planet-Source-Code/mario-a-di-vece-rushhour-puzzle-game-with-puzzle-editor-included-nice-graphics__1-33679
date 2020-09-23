VERSION 5.00
Begin VB.Form frmLevelHigh 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You have a High Score!"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Solution"
      Enabled         =   0   'False
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next Level"
      Default         =   -1  'True
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit Game"
      Height          =   375
      Left            =   3210
      TabIndex        =   0
      Top             =   2220
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   240
      Picture         =   "frmLevelHigh.frx":0000
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2670
      TabIndex        =   5
      Top             =   270
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2670
      TabIndex        =   4
      Top             =   630
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Moves:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2670
      TabIndex        =   3
      Top             =   1020
      Width           =   1305
   End
End
Attribute VB_Name = "frmLevelHigh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
YourMoves = frmMain.List2.ListCount
YourTime = X
Dim Levels As Integer
Open App.Path & "\" & LevelPack For Input As #1
Input #1, Levels
Close #1

If LevelToLoad = Levels Then
    Command2.Enabled = False
Else
    Command2.Enabled = True
End If

CalculateScore

frmMain.Enabled = False
Label1.Caption = "Your Score: " & Score
Label2.Caption = "Your Time: " & F
Label3.Caption = "Your Moves: " & frmMain.List2.ListCount



End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
End Sub

Public Sub CalculateScore()

If TargetTime * TargetMoves > YourTime * YourMoves Then
    Score = Round(((TargetTime * TargetMoves) / (YourTime * YourMoves)) * 1000, 0)
End If

If TargetTime * TargetMoves = YourTime * YourMoves Then
    Score = 1000
End If

If TargetTime * TargetMoves < YourTime * YourMoves Then
    Score = Round(((YourTime * YourMoves) / (TargetTime * TargetMoves)) * 1000, 0)
End If

End Sub


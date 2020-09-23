VERSION 5.00
Begin VB.Form frmEndLevel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "You Win!"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Quit Game"
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next Level"
      Default         =   -1  'True
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save Solution"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1710
      Width           =   1335
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1305
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
      Left            =   2640
      TabIndex        =   1
      Top             =   690
      Width           =   1140
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
      Left            =   2640
      TabIndex        =   0
      Top             =   330
      Width           =   1230
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   210
      Picture         =   "frmEndLevel.frx":0000
      Top             =   240
      Width           =   1995
   End
End
Attribute VB_Name = "frmEndLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
If Command1.Caption = "&Save Solution" Then
frmSaveSolution.Show
Else
frmMain.RemoveHighlight
LevelToLoad = LevelToLoad
CancelLoad = False
If CancelLoad = False Then

For i = 1 To 29
    frmMain.YourCar(i).Left = frmMain.imgSquare(0).Left
Next i
frmMain.List2.Clear

frmMain.LoadNewGame
frmMain.Command1.Enabled = False
frmMain.Command2.Enabled = False
frmMain.Command3.Enabled = False
frmMain.Command4.Enabled = False
frmMain.Frame1.Caption = "Controls"

End If
Unload Me
End If
End Sub

Private Sub Command2_Click()
frmMain.RemoveHighlight
LevelToLoad = LevelToLoad + 1
CancelLoad = False
If CancelLoad = False Then

For i = 1 To 29
    frmMain.YourCar(i).Left = frmMain.imgSquare(0).Left
Next i
frmMain.List2.Clear

frmMain.LoadNewGame
frmMain.Command1.Enabled = False
frmMain.Command2.Enabled = False
frmMain.Command3.Enabled = False
frmMain.Command4.Enabled = False
frmMain.Frame1.Caption = "Controls"

End If
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
Unload frmMain
End Sub

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

If Score > 3000 Then
    Command1.Enabled = True
    Me.Caption = "You Win! HIGH SCORE ACHIEVED!"
    Command1.Caption = "&Save Solution"

Else
    Command1.Enabled = True
    Me.Caption = "You Win!"
    Command1.Caption = "&Try Again"
End If



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

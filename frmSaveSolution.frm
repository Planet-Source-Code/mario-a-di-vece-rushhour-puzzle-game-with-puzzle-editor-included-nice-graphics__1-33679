VERSION 5.00
Begin VB.Form frmSaveSolution 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Solution Dialog"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   1650
      TabIndex        =   5
      Top             =   2610
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   405
      Left            =   2940
      TabIndex        =   4
      Top             =   2610
      Width           =   1155
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   1650
      Pattern         =   "*.sol"
      TabIndex        =   1
      Top             =   780
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   390
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saved Solutions:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1470
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Solution Name:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "frmSaveSolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
For i = 0 To File1.ListCount - 1
    File1.ListIndex = i
    If File1.FileName = Text1.Text & ".sol" Then
        GoTo FileIsRepeated
    End If
Next i

Open App.Path & "\" & Text1.Text & ".sol" For Output As #1
Print #1, "Rush Hour 1.0 Solution file"
Print #1, "Please DO NOT MODIFY!"
Print #1, "Level Pack = " & LevelPack
Print #1, "Level Number = " & LevelToLoad
Print #1, "Target Moves = " & TargetMoves
Print #1, "Target Time = " & TargetTime
Print #1, "Your Moves = " & YourMoves
Print #1, "Your Time = " & YourTime

For i = 0 To frmMain.List2.ListCount - 1
    frmMain.List2.ListIndex = i
    Print #1, frmMain.List2.Text
Next i

Print #1, "EOF"

Close #1

MsgBox "Solution was saved.", vbInformation, "Rush Hour"

Unload Me

Exit Sub

FileIsRepeated:

MsgBox "The solution name you specified is already in use.", vbCritical, "Error"
Text1.SetFocus

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
File1.Path = App.Path
frmEndLevel.Enabled = False
'Text1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmEndLevel.Enabled = True

End Sub

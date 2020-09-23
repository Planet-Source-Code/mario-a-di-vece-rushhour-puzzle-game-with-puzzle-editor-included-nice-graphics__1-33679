VERSION 5.00
Begin VB.Form frmLevels 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Puzzle"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   1350
      TabIndex        =   2
      Top             =   600
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2445
   End
End
Attribute VB_Name = "frmLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmMain.RemoveHighlight
LevelToLoad = Combo1.Text
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

Private Sub Command2_Click()
CancelLoad = True
Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrorFound
Dim Levels As Integer
Open App.Path & "\" & LevelPack For Input As #1
Input #1, Levels
Close #1

frmMain.Enabled = False
For i = 1 To Levels
Combo1.AddItem i
Next i
Combo1.ListIndex = 0
Exit Sub
ErrorFound:
MsgBox "The level pack you specified is not valid.", vbCritical, "Error"
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Enabled = True
End Sub

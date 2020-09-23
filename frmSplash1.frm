VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1515
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3810
      Top             =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You for Playing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   1140
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   0
      Picture         =   "frmSplash1.frx":000C
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartedOn As Long
Dim SecondsToStay As Integer

Private Sub Form_Load()
Unload frmMain
StartedOn = Timer
SecondsToStay = 2

End Sub


Private Sub Timer1_Timer()
If Timer - StartedOn >= SecondsToStay Then
    Unload Me
    End
End If

End Sub

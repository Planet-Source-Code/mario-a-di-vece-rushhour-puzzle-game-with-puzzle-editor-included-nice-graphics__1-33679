VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2010
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   1050
   End
   Begin VB.Label Label3 
      Caption         =   "  Loading..."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   2805
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2001 ByteDive Entertainment"
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mario Di Vece"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   1110
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   -30
      Picture         =   "frmSplash.frx":000C
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartedOn As Long
Dim SecondsToStay As Integer

Private Sub Form_Load()
StartedOn = Timer
SecondsToStay = 3

End Sub


Private Sub Timer1_Timer()
If Timer - StartedOn >= SecondsToStay Then
    Unload Me
    frmMain.Show
End If

End Sub

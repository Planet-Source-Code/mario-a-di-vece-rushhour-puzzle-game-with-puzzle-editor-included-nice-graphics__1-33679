VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rush Hour 1.1"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   6555
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox TimeDots 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   7890
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   150
      ScaleWidth      =   60
      TabIndex        =   22
      Top             =   5580
      Width           =   60
   End
   Begin VB.PictureBox SecondLeds 
      BackColor       =   &H80000007&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7920
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   21
      Top             =   5460
      Width           =   615
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   0
      Left            =   10830
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   20
      Top             =   1890
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   1
      Left            =   11190
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   19
      Top             =   1890
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   11550
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   18
      Top             =   1890
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   3
      Left            =   11910
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Top             =   1890
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   4
      Left            =   12270
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Top             =   1890
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   5
      Left            =   10830
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Top             =   2250
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   6
      Left            =   11190
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   14
      Top             =   2250
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   7
      Left            =   11550
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   2250
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   8
      Left            =   11910
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   2250
      Width           =   300
   End
   Begin VB.PictureBox LED 
      AutoRedraw      =   -1  'True
      Height          =   300
      Index           =   9
      Left            =   12270
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   2250
      Width           =   300
   End
   Begin VB.PictureBox TimeLeds 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   5460
      Width           =   615
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   225
      Left            =   10770
      TabIndex        =   8
      Top             =   780
      Width           =   1725
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   11400
      Top             =   1230
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   10800
      TabIndex        =   5
      Top             =   390
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10860
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1545
      Left            =   10800
      TabIndex        =   0
      Top             =   2760
      Width           =   1905
      Begin VB.CommandButton Command4 
         Caption         =   "Down"
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1110
         Width           =   585
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Up"
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Top             =   300
         Width           =   585
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Right"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   690
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Left"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Top             =   690
         Width           =   765
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6750
      TabIndex        =   28
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Car: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6840
      TabIndex        =   27
      Top             =   3990
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Number: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6870
      TabIndex        =   26
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level Pack: Standard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6870
      TabIndex        =   25
      Top             =   1620
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record Time: 00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6870
      TabIndex        =   24
      Top             =   3000
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Record Moves: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6870
      TabIndex        =   23
      Top             =   2610
      Width           =   1365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moves: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6840
      TabIndex        =   9
      Top             =   3630
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   6870
      Picture         =   "frmMain.frx":0485
      Top             =   330
      Width           =   1995
   End
   Begin VB.Image CDown 
      Height          =   570
      Left            =   7770
      Picture         =   "frmMain.frx":0F4B
      ToolTipText     =   "Move Down"
      Top             =   10980
      Width           =   435
   End
   Begin VB.Image CUp 
      Height          =   570
      Left            =   6480
      Picture         =   "frmMain.frx":1015
      ToolTipText     =   "Move Up"
      Top             =   10950
      Width           =   435
   End
   Begin VB.Image CLeft 
      Height          =   435
      Left            =   6990
      Picture         =   "frmMain.frx":10DC
      ToolTipText     =   "Move Left"
      Top             =   11340
      Width           =   570
   End
   Begin VB.Image CRight 
      Height          =   435
      Left            =   7050
      Picture         =   "frmMain.frx":11A2
      ToolTipText     =   "MoveRight"
      Top             =   10770
      Width           =   570
   End
   Begin VB.Shape V2HL 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1635
      Left            =   5610
      Top             =   9210
      Width           =   825
   End
   Begin VB.Shape V3HL 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   2445
      Left            =   8220
      Top             =   8310
      Width           =   825
   End
   Begin VB.Shape H2HL 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   825
      Left            =   6510
      Top             =   9360
      Width           =   1635
   End
   Begin VB.Shape H3HL 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   825
      Left            =   5610
      Top             =   8310
      Width           =   2445
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   28
      Left            =   15690
      Picture         =   "frmMain.frx":1268
      Top             =   10860
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   27
      Left            =   14760
      Picture         =   "frmMain.frx":1BE3
      Top             =   10830
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   26
      Left            =   15510
      Picture         =   "frmMain.frx":24F0
      Top             =   8370
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   25
      Left            =   12240
      Picture         =   "frmMain.frx":2E90
      Top             =   9210
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   24
      Left            =   12240
      Picture         =   "frmMain.frx":37A2
      Top             =   10020
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   23
      Left            =   12270
      Picture         =   "frmMain.frx":4312
      Top             =   10830
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   22
      Left            =   10650
      Picture         =   "frmMain.frx":4BF5
      Top             =   9720
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   21
      Left            =   10650
      Picture         =   "frmMain.frx":51D3
      Top             =   10470
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   20
      Left            =   10650
      Picture         =   "frmMain.frx":5768
      Top             =   11220
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   19
      Left            =   12450
      Picture         =   "frmMain.frx":5D0B
      Top             =   11910
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   18
      Left            =   10710
      Picture         =   "frmMain.frx":62C8
      Top             =   12060
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   17
      Left            =   9090
      Picture         =   "frmMain.frx":68A0
      Top             =   8220
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   16
      Left            =   9090
      Picture         =   "frmMain.frx":6E09
      Top             =   8940
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   15
      Left            =   9090
      Picture         =   "frmMain.frx":7403
      Top             =   9660
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   14
      Left            =   9090
      Picture         =   "frmMain.frx":786A
      Top             =   10380
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   29
      Left            =   10680
      Picture         =   "frmMain.frx":7E19
      Top             =   8250
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   12
      Left            =   10560
      Picture         =   "frmMain.frx":83CD
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   11
      Left            =   11340
      Picture         =   "frmMain.frx":8A4E
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   10
      Left            =   12090
      Picture         =   "frmMain.frx":8F56
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   9
      Left            =   12840
      Picture         =   "frmMain.frx":961E
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   8
      Left            =   13590
      Picture         =   "frmMain.frx":9C5A
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   7
      Left            =   14340
      Picture         =   "frmMain.frx":A2E7
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   6
      Left            =   9150
      Picture         =   "frmMain.frx":A953
      Top             =   11190
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   13
      Left            =   9900
      Picture         =   "frmMain.frx":AFB7
      Top             =   11190
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   5
      Left            =   14640
      Picture         =   "frmMain.frx":B64A
      Top             =   8370
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   4
      Left            =   12210
      Picture         =   "frmMain.frx":BFFE
      Top             =   8340
      Width           =   2355
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time: 00:00"
      Height          =   255
      Left            =   10800
      TabIndex        =   7
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   3
      Left            =   9840
      Picture         =   "frmMain.frx":C925
      Top             =   6720
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Occupied Blocks (0)"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   10800
      TabIndex        =   6
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   2
      Left            =   9090
      Picture         =   "frmMain.frx":CF77
      Top             =   6720
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   1
      Left            =   10710
      Picture         =   "frmMain.frx":D5E6
      Top             =   9030
      Width           =   1485
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   20
      Left            =   1800
      Picture         =   "frmMain.frx":DB99
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   21
      Left            =   2610
      Picture         =   "frmMain.frx":DED9
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   22
      Left            =   3420
      Picture         =   "frmMain.frx":E219
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   23
      Left            =   4230
      Picture         =   "frmMain.frx":E559
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   24
      Left            =   5040
      Picture         =   "frmMain.frx":E899
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   25
      Left            =   990
      Picture         =   "frmMain.frx":EBD9
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   26
      Left            =   1800
      Picture         =   "frmMain.frx":EF19
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   27
      Left            =   2610
      Picture         =   "frmMain.frx":F259
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   28
      Left            =   3420
      Picture         =   "frmMain.frx":F599
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   29
      Left            =   4230
      Picture         =   "frmMain.frx":F8D9
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   30
      Left            =   5040
      Picture         =   "frmMain.frx":FC19
      Top             =   4170
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   31
      Left            =   990
      Picture         =   "frmMain.frx":FF59
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   32
      Left            =   1800
      Picture         =   "frmMain.frx":10299
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   33
      Left            =   2610
      Picture         =   "frmMain.frx":105D9
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   34
      Left            =   3420
      Picture         =   "frmMain.frx":10919
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   17
      Left            =   4230
      Picture         =   "frmMain.frx":10C59
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   18
      Left            =   5040
      Picture         =   "frmMain.frx":10F99
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   1
      Left            =   990
      Picture         =   "frmMain.frx":112D9
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   2
      Left            =   1800
      Picture         =   "frmMain.frx":11619
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   3
      Left            =   2610
      Picture         =   "frmMain.frx":11959
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   4
      Left            =   3420
      Picture         =   "frmMain.frx":11C99
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   5
      Left            =   4230
      Picture         =   "frmMain.frx":11FD9
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   6
      Left            =   5040
      Picture         =   "frmMain.frx":12319
      Top             =   930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   7
      Left            =   990
      Picture         =   "frmMain.frx":12659
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   8
      Left            =   1800
      Picture         =   "frmMain.frx":12999
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   9
      Left            =   2610
      Picture         =   "frmMain.frx":12CD9
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   10
      Left            =   3420
      Picture         =   "frmMain.frx":13019
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   11
      Left            =   4230
      Picture         =   "frmMain.frx":13359
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   12
      Left            =   5040
      Picture         =   "frmMain.frx":13699
      Top             =   1740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   13
      Left            =   990
      Picture         =   "frmMain.frx":139D9
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   14
      Left            =   1800
      Picture         =   "frmMain.frx":13D19
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   15
      Left            =   2610
      Picture         =   "frmMain.frx":14059
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   16
      Left            =   3420
      Picture         =   "frmMain.frx":14399
      Top             =   2550
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   0
      Left            =   15180
      Picture         =   "frmMain.frx":146D9
      Top             =   6750
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   19
      Left            =   990
      Picture         =   "frmMain.frx":14A19
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   35
      Left            =   4230
      Picture         =   "frmMain.frx":14D59
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   36
      Left            =   5040
      Picture         =   "frmMain.frx":15099
      Top             =   4980
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   270
      Picture         =   "frmMain.frx":153D9
      Top             =   240
      Width           =   6240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu smnuNewGame 
         Caption         =   "New Game"
      End
      Begin VB.Menu smnuLCS 
         Caption         =   "Load Level Pack"
      End
      Begin VB.Menu smnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu smnuExit 
         Caption         =   "Exit Game"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Not a single line of code was taken from any source
'All the code and graphics were done from scratch by me,
'MARIO DI VECE. (mariodivece@hotmail.com) (except for the LED Module)
'This project is used to demonstrate you don't need
'any strange DirectX Typelibs or extra OCX's and DLL's
'to develop a fairly complex 2D puzzle engine.
'The next challenges here are to be able to load car packs
'so the game is fully flexible, and to code a script-checking routine
'Also, to make level previews and a Visual level builder.
'A Car-Labeling routine, and a Highest scores table
'A gamestate saver and a solution saver with record moves number
'I think the hard part is alredy done... This code is
'relatively small (1000 lines) and simple to be a 2D
'game engine. Proof that VB is also suitable for game programming

'Any comments or questions, send e-mail to mariodivece@hotmail.com

'The original game is distributed by Binary Arts
'You are free to learn from this code, but not to
'use it for commercial purposes unless you ask me for a written
'permission.
'If you do not accept this, please close this project
'and try to use someone else's ideas.

'To the guys from Microsoft: Visual Basic is a great programming
'language. Congrats! But I think you should put a little more emphasis on
'the shitty executables it delivers because they are not true
'executables... It is just some true executable header with "links"
'to a DLL
'WHAT A SHAME! Even the legendary QBasic had a fairly decent compiler!

'HAVE FUN! VB'ers

Dim ActiveCar As Integer
Dim OB(36) As Boolean
Dim RightLimits(5) As Integer
Dim Starttime As Long
Dim TotalMoves As Integer
Dim TotalTime As Long


Private Sub LoadCarPositions()
On Local Error GoTo ErrorFound

Starttime = Timer
Timer2.Enabled = True

OB(1) = False
OB(2) = False
OB(3) = False
OB(4) = False
OB(5) = False
OB(6) = False
OB(7) = False
OB(8) = False
OB(9) = False
OB(10) = False
OB(11) = False
OB(12) = False
OB(13) = False
OB(14) = False
OB(15) = False
OB(16) = False
OB(17) = False
OB(18) = False
OB(19) = False
OB(20) = False
OB(21) = False
OB(22) = False
OB(23) = False
OB(24) = False
OB(25) = False
OB(26) = False
OB(27) = False
OB(28) = False
OB(29) = False
OB(30) = False
OB(31) = False
OB(32) = False
OB(33) = False
OB(34) = False
OB(35) = False
OB(36) = False

RightLimits(0) = 6
RightLimits(1) = 6 * 2
RightLimits(2) = Empty
RightLimits(3) = 6 * 4
RightLimits(4) = 6 * 5
RightLimits(5) = 6 * 6


Dim Trash As String
Dim LoadedCar As Variant
Dim LoadedBlock As Variant
Dim OccupationType As String
Dim NumberOfCars As Variant

Open App.Path & "\" & LevelPack For Input As #1

RepeatOP:
Input #1, Trash
If Trash <> "-" & LevelToLoad & "-" Then
    GoTo RepeatOP
End If

Input #1, NumberOfCars
Input #1, TargetMoves
Input #1, TargetTime

For i = 1 To Val(NumberOfCars) 'number of cars to load
Input #1, LoadedCar
Input #1, LoadedBlock
Input #1, OccupationType
YourCar(Val(LoadedCar)).Left = imgSquare(Val(LoadedBlock)).Left
YourCar(Val(LoadedCar)).Top = imgSquare(Val(LoadedBlock)).Top

    If OccupationType = "H2" Then
        OB(Val(LoadedBlock)) = True
        OB(Val(LoadedBlock) + 1) = True
    End If
    
    If OccupationType = "V2" Then
        OB(Val(LoadedBlock)) = True
        OB(Val(LoadedBlock) + 6) = True
    End If
    
    If OccupationType = "H3" Then
        OB(Val(LoadedBlock)) = True
        OB(Val(LoadedBlock) + 1) = True
        OB(Val(LoadedBlock) + 2) = True

    End If
    
    If OccupationType = "V3" Then
        OB(Val(LoadedBlock)) = True
        OB(Val(LoadedBlock) + 6) = True
        OB(Val(LoadedBlock) + 12) = True
    End If
Next i

Close #1

Call FindOccuppiedBlocks
Exit Sub
ErrorFound:
MsgBox "Errors in level script. Script Unloadable. Program will end"
End
End Sub

Private Sub GetControls(CarNumber As Integer)
FindPosition
Frame1.Caption = "Controls " & "( Car " & ActiveCar & " )"

If CarSet(ActiveCar).NextLeftBlock = 0 Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If
'---------------------------------------------------------------
check1:
If CarSet(ActiveCar).NextRightBlock = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If

For i = 0 To 5
 If CarSet(ActiveCar).Position = "Horizontal2" Or CarSet(ActiveCar).Position = "Vertical2" Then
    If CarSet(ActiveCar).NextRightBlock = RightLimits(i) Then
        Command2.Enabled = False
        GoTo check21
    End If
 End If
Next i

check21:

For i = 0 To 5
 If CarSet(ActiveCar).Position = "Horizontal3" Or CarSet(ActiveCar).Position = "Vertical3" Then
    If CarSet(ActiveCar).NextRightBlock + 1 = RightLimits(i) Then
        Command2.Enabled = False
        GoTo check2
    End If
 End If
Next i
'---------------------------------------------------------------
check2:
If CarSet(ActiveCar).NextTopBlock = 0 Then
Command3.Enabled = False
Else
Command3.Enabled = True
End If

check3:
If CarSet(ActiveCar).NextDownBlock = 0 Then
Command4.Enabled = False
Else
Command4.Enabled = True
End If

'-------------------------------------------------------------
If CarSet(ActiveCar).NextDownBlock = 31 Then Command4.Enabled = False
If CarSet(ActiveCar).NextDownBlock = 32 Then Command4.Enabled = False
If CarSet(ActiveCar).NextDownBlock = 33 Then Command4.Enabled = False
If CarSet(ActiveCar).NextDownBlock = 34 Then Command4.Enabled = False
If CarSet(ActiveCar).NextDownBlock = 35 Then Command4.Enabled = False
If CarSet(ActiveCar).NextDownBlock = 36 Then Command4.Enabled = False
'-------------------------------------------------------------------
If CarSet(ActiveCar).Position = "Vertical3" Then
    If CarSet(ActiveCar).NextDownBlock + 12 > 36 Then
        Command4.Enabled = False
    End If
End If

ShowControls

End Sub


Private Sub CDown_Click()
MoveDown
End Sub

Private Sub CLeft_Click()
MoveLeft
End Sub

Private Sub Command1_Click()
MoveLeft
End Sub

Private Sub Command2_Click()
MoveRight
End Sub

Private Sub Command3_Click()
MoveUp
End Sub

Private Sub Command4_Click()
MoveDown
End Sub

Public Function LoadNewGame()
LoadCarPositions
Label1.Caption = "Occupied Blocks (" & List1.ListCount & ")"
End Function


Private Sub CRight_Click()
MoveRight
End Sub

Private Sub CUp_Click()
MoveUp
End Sub

Private Sub ArrowMove(KeyCode As Integer)

Select Case KeyCode
    Case vbKeyUp
        If Command3.Enabled = True Then MoveUp
    Case vbKeyDown
        If Command4.Enabled = True Then MoveDown
    Case vbKeyLeft
        If Command1.Enabled = True Then MoveLeft
    Case vbKeyRight
        If Command2.Enabled = True Then MoveRight
End Select

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
ArrowMove KeyCode
End Sub

Private Sub Form_Load()
KeyPreview = True
Label9.Caption = ""
LoadGFX LED, True, frmMain.BackColor, &H404040, vbGreen
TimeLeds.BackColor = frmMain.BackColor    '&H80000007 '&H8000&
TimeLeds.BorderStyle = 0
SecondLeds.BackColor = frmMain.BackColor    '&H80000007
SecondLeds.BorderStyle = 0
TimeDots.BackColor = frmMain.BackColor    '&H80000007
TimeDots.BorderStyle = 0
DrawLED TimeLeds, Mid(Label2.Caption, 7, 2), LED, 1
DrawLED SecondLeds, Mid(Label2.Caption, 10, 2), LED, 1

LevelPack = "Standard.pak"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmSplash1.Show
End Sub

Private Sub smnuAbout_Click()
frmAbout.Show
End Sub

Private Sub smnuExit_Click()
frmSplash1.Show
End Sub

Private Sub smnuHighScores_Click()
frmHighScores.Show
End Sub

Private Sub smnuLCS_Click()
frmPackSelector.Show
End Sub

Private Sub smnuNewGame_Click()
On Local Error Resume Next
frmLevels.Show
End Sub



Private Sub Timer1_Timer()
FindOccuppiedBlocks
Label1.Caption = "Occupied Blocks (" & List1.ListCount & ")"
If YourCar(1).Left = imgSquare(18).Left And YourCar(1).Top = imgSquare(18).Top Then
RemoveHighlight
frmEndLevel.Show
Timer1.Enabled = False
Timer2.Enabled = False
Starttime = 0
End If
End Sub

Private Sub Timer2_Timer()

X = Round(Timer - Starttime, 0)
TotalTime = X
F = Format(CStr(((X \ 60) Mod 24)), "00") & ":" & Format(CStr(X Mod 60), "00")
Label2.Caption = "Time: " & F

DrawLED TimeLeds, Mid(Label2.Caption, 7, 2), LED, 1
DrawLED SecondLeds, Mid(Label2.Caption, 10, 2), LED, 1

DisplayStats
End Sub



Private Sub YourCar_Click(Index As Integer)
Timer1.Enabled = True

Select Case Index
    Case 1
    ActiveCar = 1
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 2
    ActiveCar = 2
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 3
    ActiveCar = 3
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 4
    ActiveCar = 4
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal3"
    GetControls (ActiveCar)
    Case 5
    ActiveCar = 5
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical3"
    GetControls (ActiveCar)
    Case 6
    ActiveCar = 6
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 7
    ActiveCar = 7
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 8
    ActiveCar = 8
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 9
    ActiveCar = 9
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 10
    ActiveCar = 10
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 11
    ActiveCar = 11
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 12
    ActiveCar = 12
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 13
    ActiveCar = 13
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical2"
    GetControls (ActiveCar)
    Case 14
    ActiveCar = 14
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 15
    ActiveCar = 15
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 16
    ActiveCar = 16
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 17
    ActiveCar = 17
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 18
    ActiveCar = 18
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 19
    ActiveCar = 19
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 20
    ActiveCar = 20
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 21
    ActiveCar = 21
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 22
    ActiveCar = 22
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
    Case 23
    ActiveCar = 23
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal3"
    GetControls (ActiveCar)
    Case 24
    ActiveCar = 24
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal3"
    GetControls (ActiveCar)
    Case 25
    ActiveCar = 25
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Horizontal3"
    GetControls (ActiveCar)
    Case 26
    ActiveCar = 26
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Vertical3"
    GetControls (ActiveCar)
    Case 27
    ActiveCar = 27
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical3"
    GetControls (ActiveCar)
    Case 28
    ActiveCar = 28
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "Vertical3"
    GetControls (ActiveCar)
    Case 29
    ActiveCar = 29
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "Horizontal2"
    GetControls (ActiveCar)
End Select

    HighLight


End Sub

Private Sub MoveLeft()
FindPosition

If CarSet(ActiveCar).Position = "Horizontal2" Then
If OB(CarSet(ActiveCar).NextLeftBlock) = False Then
YourCar(ActiveCar).Left = imgSquare(CarSet(ActiveCar).NextLeftBlock).Left
OB(CarSet(ActiveCar).NextLeftBlock) = True
OB(CarSet(ActiveCar).NextLeftBlock + 1) = True
OB(CarSet(ActiveCar).NextLeftBlock + 2) = False
List2.AddItem ActiveCar & " - L"
HighLight

Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If
'-------------------------------------------
If CarSet(ActiveCar).Position = "Horizontal3" Then '*
If OB(CarSet(ActiveCar).NextLeftBlock) = False Then
YourCar(ActiveCar).Left = imgSquare(CarSet(ActiveCar).NextLeftBlock).Left
OB(CarSet(ActiveCar).NextLeftBlock) = True
OB(CarSet(ActiveCar).NextLeftBlock + 1) = True
OB(CarSet(ActiveCar).NextLeftBlock + 2) = True
OB(CarSet(ActiveCar).NextLeftBlock + 3) = False
List2.AddItem ActiveCar & " - L"
HighLight

Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If

End Sub

Private Sub MoveRight()
FindPosition

If CarSet(ActiveCar).Position = "Horizontal2" Then
If OB(CarSet(ActiveCar).NextRightBlock + 1) = False Then
YourCar(ActiveCar).Left = imgSquare(CarSet(ActiveCar).NextRightBlock).Left
OB(CarSet(ActiveCar).NextRightBlock) = True
OB(CarSet(ActiveCar).NextRightBlock + 1) = True
OB(CarSet(ActiveCar).NextRightBlock - 1) = False
List2.AddItem ActiveCar & " - R"
HighLight

Else
If ActiveCar = 1 And CarSet(ActiveCar).NextRightBlock = 18 Then
YourCar(ActiveCar).Left = imgSquare(CarSet(ActiveCar).NextRightBlock).Left
Else
Label9.Caption = "Move not allowed"
End If
End If
GetControls (ActiveCar)
End If
'------------------------------------------------
If CarSet(ActiveCar).Position = "Horizontal3" Then '*
If OB(CarSet(ActiveCar).NextRightBlock + 2) = False Then
YourCar(ActiveCar).Left = imgSquare(CarSet(ActiveCar).NextRightBlock).Left
OB(CarSet(ActiveCar).NextRightBlock - 1) = False
OB(CarSet(ActiveCar).NextRightBlock) = True
OB(CarSet(ActiveCar).NextRightBlock + 1) = True
OB(CarSet(ActiveCar).NextRightBlock + 2) = True
List2.AddItem ActiveCar & " - R"
HighLight

Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If

End Sub
Private Sub MoveUp()
FindPosition

If CarSet(ActiveCar).Position = "Vertical2" Then
If OB(CarSet(ActiveCar).NextTopBlock) = False Then
YourCar(ActiveCar).Top = imgSquare(CarSet(ActiveCar).NextTopBlock).Top
OB(CarSet(ActiveCar).NextTopBlock) = True '*
OB(CarSet(ActiveCar).NextTopBlock + 6) = True
OB(CarSet(ActiveCar).NextTopBlock + 12) = False
List2.AddItem ActiveCar & " - U"
HighLight

Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If
'-------------------------------------------------------------------
If CarSet(ActiveCar).Position = "Vertical3" Then
If OB(CarSet(ActiveCar).NextTopBlock) = False Then
YourCar(ActiveCar).Top = imgSquare(CarSet(ActiveCar).NextTopBlock).Top
OB(CarSet(ActiveCar).NextTopBlock) = True '*
OB(CarSet(ActiveCar).NextTopBlock + 6) = True '*
OB(CarSet(ActiveCar).NextTopBlock + 12) = True '*
OB(CarSet(ActiveCar).NextTopBlock + 18) = False
List2.AddItem ActiveCar & " - U"
HighLight
Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If

End Sub
Private Sub MoveDown()

FindPosition

If CarSet(ActiveCar).Position = "Vertical2" Then
If OB(CarSet(ActiveCar).NextDownBlock + 6) = False Then
YourCar(ActiveCar).Top = imgSquare(CarSet(ActiveCar).NextDownBlock).Top
OB(CarSet(ActiveCar).NextDownBlock) = True
OB(CarSet(ActiveCar).NextDownBlock + 6) = True
OB(CarSet(ActiveCar).NextDownBlock - 6) = False '*
List2.AddItem ActiveCar & " - D"
HighLight
Else
Label9.Caption = "Move not allowed"
End If
GetControls (ActiveCar)
End If
'-----------------------------------------
If CarSet(ActiveCar).Position = "Vertical3" Then
    If CarSet(ActiveCar).NextDownBlock + 12 <= 36 Then
        If OB(CarSet(ActiveCar).NextDownBlock + 12) = False Then
            YourCar(ActiveCar).Top = imgSquare(CarSet(ActiveCar).NextDownBlock).Top
            OB(CarSet(ActiveCar).NextDownBlock) = True
            OB(CarSet(ActiveCar).NextDownBlock + 6) = True '*
            OB(CarSet(ActiveCar).NextDownBlock + 12) = True '*
            OB(CarSet(ActiveCar).NextDownBlock - 6) = False
            List2.AddItem ActiveCar & " - D"
            HighLight
        Else
            Label9.Caption = "Move not allowed"
        End If
    Else
        Command4.Enabled = False
    End If
GetControls (ActiveCar)
End If

End Sub
Private Sub FindPosition()

If CarSet(ActiveCar).Position = "Horizontal2" Then
Call Horizontal2Logics
End If

If CarSet(ActiveCar).Position = "Vertical2" Then
Call Vertical2Logics
End If

If CarSet(ActiveCar).Position = "Horizontal3" Then '*
Call Horizontal2Logics
End If

If CarSet(ActiveCar).Position = "Vertical3" Then '*
Call Vertical2Logics
End If

End Sub


Private Sub FindOccuppiedBlocks()
List1.Clear
    For i = 1 To 36
        If OB(i) = True Then List1.AddItem i
    Next i
End Sub

Private Sub Horizontal2Logics()
    If YourCar(ActiveCar).Left = imgSquare(1).Left And YourCar(ActiveCar).Top = imgSquare(1).Top Then
        CarSet(ActiveCar).NextRightBlock = 2
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(2).Left And YourCar(ActiveCar).Top = imgSquare(2).Top Then
        CarSet(ActiveCar).NextRightBlock = 3
        CarSet(ActiveCar).NextLeftBlock = 1
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(3).Left And YourCar(ActiveCar).Top = imgSquare(3).Top Then
        CarSet(ActiveCar).NextRightBlock = 4
        CarSet(ActiveCar).NextLeftBlock = 2
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(4).Left And YourCar(ActiveCar).Top = imgSquare(4).Top Then
        CarSet(ActiveCar).NextRightBlock = 5
        CarSet(ActiveCar).NextLeftBlock = 3
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(5).Left And YourCar(ActiveCar).Top = imgSquare(5).Top Then
        CarSet(ActiveCar).NextRightBlock = 6
        CarSet(ActiveCar).NextLeftBlock = 4
    End If
    
    'Block6 is ommitted because it is not reachable

    If YourCar(ActiveCar).Left = imgSquare(7).Left And YourCar(ActiveCar).Top = imgSquare(7).Top Then
        CarSet(ActiveCar).NextRightBlock = 8
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(8).Left And YourCar(ActiveCar).Top = imgSquare(8).Top Then
        CarSet(ActiveCar).NextRightBlock = 9
        CarSet(ActiveCar).NextLeftBlock = 7
    End If

    If YourCar(ActiveCar).Left = imgSquare(9).Left And YourCar(ActiveCar).Top = imgSquare(9).Top Then
        CarSet(ActiveCar).NextRightBlock = 10
        CarSet(ActiveCar).NextLeftBlock = 8
    End If

    If YourCar(ActiveCar).Left = imgSquare(10).Left And YourCar(ActiveCar).Top = imgSquare(10).Top Then
        CarSet(ActiveCar).NextRightBlock = 11
        CarSet(ActiveCar).NextLeftBlock = 9
    End If

    If YourCar(ActiveCar).Left = imgSquare(11).Left And YourCar(ActiveCar).Top = imgSquare(11).Top Then
        CarSet(ActiveCar).NextRightBlock = 12
        CarSet(ActiveCar).NextLeftBlock = 10
    End If

    'Block 12 is omitted because it is unreachable

    If YourCar(ActiveCar).Left = imgSquare(13).Left And YourCar(ActiveCar).Top = imgSquare(13).Top Then
        CarSet(ActiveCar).NextRightBlock = 14
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(14).Left And YourCar(ActiveCar).Top = imgSquare(14).Top Then
        CarSet(ActiveCar).NextRightBlock = 15
        CarSet(ActiveCar).NextLeftBlock = 13
    End If

    If YourCar(ActiveCar).Left = imgSquare(15).Left And YourCar(ActiveCar).Top = imgSquare(15).Top Then
        CarSet(ActiveCar).NextRightBlock = 16
        CarSet(ActiveCar).NextLeftBlock = 14
    End If

    If YourCar(ActiveCar).Left = imgSquare(16).Left And YourCar(ActiveCar).Top = imgSquare(16).Top Then
        CarSet(ActiveCar).NextRightBlock = 17
        CarSet(ActiveCar).NextLeftBlock = 15
    End If

    If YourCar(ActiveCar).Left = imgSquare(17).Left And YourCar(ActiveCar).Top = imgSquare(17).Top Then
        CarSet(ActiveCar).NextRightBlock = 18
        CarSet(ActiveCar).NextLeftBlock = 16
    End If
    
    '*The following move definition is an exeption to complete the game
    If YourCar(ActiveCar).Left = imgSquare(18).Left And YourCar(ActiveCar).Top = imgSquare(18).Top Then
        CarSet(ActiveCar).NextRightBlock = 0
        CarSet(ActiveCar).NextLeftBlock = 0
    End If
    'End of exeption

    If YourCar(ActiveCar).Left = imgSquare(19).Left And YourCar(ActiveCar).Top = imgSquare(19).Top Then
        CarSet(ActiveCar).NextRightBlock = 20
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(20).Left And YourCar(ActiveCar).Top = imgSquare(20).Top Then
        CarSet(ActiveCar).NextRightBlock = 21
        CarSet(ActiveCar).NextLeftBlock = 19
    End If

    If YourCar(ActiveCar).Left = imgSquare(21).Left And YourCar(ActiveCar).Top = imgSquare(21).Top Then
        CarSet(ActiveCar).NextRightBlock = 22
        CarSet(ActiveCar).NextLeftBlock = 20
    End If

    If YourCar(ActiveCar).Left = imgSquare(22).Left And YourCar(ActiveCar).Top = imgSquare(22).Top Then
        CarSet(ActiveCar).NextRightBlock = 23
        CarSet(ActiveCar).NextLeftBlock = 21
    End If

    If YourCar(ActiveCar).Left = imgSquare(23).Left And YourCar(ActiveCar).Top = imgSquare(23).Top Then
        CarSet(ActiveCar).NextRightBlock = 24
        CarSet(ActiveCar).NextLeftBlock = 22
    End If

    'Block 24 is omitted because it is unreachable

    If YourCar(ActiveCar).Left = imgSquare(25).Left And YourCar(ActiveCar).Top = imgSquare(25).Top Then
        CarSet(ActiveCar).NextRightBlock = 26
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(26).Left And YourCar(ActiveCar).Top = imgSquare(26).Top Then
        CarSet(ActiveCar).NextRightBlock = 27
        CarSet(ActiveCar).NextLeftBlock = 25
    End If

    If YourCar(ActiveCar).Left = imgSquare(27).Left And YourCar(ActiveCar).Top = imgSquare(27).Top Then
        CarSet(ActiveCar).NextRightBlock = 28
        CarSet(ActiveCar).NextLeftBlock = 26
    End If

    If YourCar(ActiveCar).Left = imgSquare(28).Left And YourCar(ActiveCar).Top = imgSquare(28).Top Then
        CarSet(ActiveCar).NextRightBlock = 29
        CarSet(ActiveCar).NextLeftBlock = 27
    End If

    If YourCar(ActiveCar).Left = imgSquare(29).Left And YourCar(ActiveCar).Top = imgSquare(29).Top Then
        CarSet(ActiveCar).NextRightBlock = 30
        CarSet(ActiveCar).NextLeftBlock = 28
    End If

    'Block 30 is omitted because it is unreachable

    If YourCar(ActiveCar).Left = imgSquare(31).Left And YourCar(ActiveCar).Top = imgSquare(31).Top Then
        CarSet(ActiveCar).NextRightBlock = 32
        CarSet(ActiveCar).NextLeftBlock = 0
    End If

    If YourCar(ActiveCar).Left = imgSquare(32).Left And YourCar(ActiveCar).Top = imgSquare(32).Top Then
        CarSet(ActiveCar).NextRightBlock = 33
        CarSet(ActiveCar).NextLeftBlock = 31
    End If

    If YourCar(ActiveCar).Left = imgSquare(33).Left And YourCar(ActiveCar).Top = imgSquare(33).Top Then
        CarSet(ActiveCar).NextRightBlock = 34
        CarSet(ActiveCar).NextLeftBlock = 32
    End If

    If YourCar(ActiveCar).Left = imgSquare(34).Left And YourCar(ActiveCar).Top = imgSquare(34).Top Then
        CarSet(ActiveCar).NextRightBlock = 35
        CarSet(ActiveCar).NextLeftBlock = 33
    End If

    If YourCar(ActiveCar).Left = imgSquare(35).Left And YourCar(ActiveCar).Top = imgSquare(35).Top Then
        CarSet(ActiveCar).NextRightBlock = 36
        CarSet(ActiveCar).NextLeftBlock = 34
    End If

    'Block 36 is omitted because it is unreachable
    'END OF HORIZONTAL 2 LOGICS
End Sub

Private Sub Vertical2Logics()
    If YourCar(ActiveCar).Left = imgSquare(1).Left And YourCar(ActiveCar).Top = imgSquare(1).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 7
    End If

    If YourCar(ActiveCar).Left = imgSquare(2).Left And YourCar(ActiveCar).Top = imgSquare(2).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 8
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(3).Left And YourCar(ActiveCar).Top = imgSquare(3).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 9
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(4).Left And YourCar(ActiveCar).Top = imgSquare(4).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 10
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(5).Left And YourCar(ActiveCar).Top = imgSquare(5).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 11
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(6).Left And YourCar(ActiveCar).Top = imgSquare(6).Top Then
        CarSet(ActiveCar).NextTopBlock = 0
        CarSet(ActiveCar).NextDownBlock = 12
    End If

'End of first line logics

    If YourCar(ActiveCar).Left = imgSquare(7).Left And YourCar(ActiveCar).Top = imgSquare(7).Top Then
        CarSet(ActiveCar).NextTopBlock = 1
        CarSet(ActiveCar).NextDownBlock = 13
    End If

    If YourCar(ActiveCar).Left = imgSquare(8).Left And YourCar(ActiveCar).Top = imgSquare(8).Top Then
        CarSet(ActiveCar).NextTopBlock = 2
        CarSet(ActiveCar).NextDownBlock = 14
    End If

    If YourCar(ActiveCar).Left = imgSquare(9).Left And YourCar(ActiveCar).Top = imgSquare(9).Top Then
        CarSet(ActiveCar).NextTopBlock = 3
        CarSet(ActiveCar).NextDownBlock = 15
    End If

    If YourCar(ActiveCar).Left = imgSquare(10).Left And YourCar(ActiveCar).Top = imgSquare(10).Top Then
        CarSet(ActiveCar).NextTopBlock = 4
        CarSet(ActiveCar).NextDownBlock = 16
    End If

    If YourCar(ActiveCar).Left = imgSquare(11).Left And YourCar(ActiveCar).Top = imgSquare(11).Top Then
        CarSet(ActiveCar).NextTopBlock = 5
        CarSet(ActiveCar).NextDownBlock = 17
    End If

    If YourCar(ActiveCar).Left = imgSquare(12).Left And YourCar(ActiveCar).Top = imgSquare(12).Top Then
        CarSet(ActiveCar).NextTopBlock = 6
        CarSet(ActiveCar).NextDownBlock = 18
    End If

'End of line 2 logics

    If YourCar(ActiveCar).Left = imgSquare(13).Left And YourCar(ActiveCar).Top = imgSquare(13).Top Then
        CarSet(ActiveCar).NextTopBlock = 7
        CarSet(ActiveCar).NextDownBlock = 19
    End If

    If YourCar(ActiveCar).Left = imgSquare(14).Left And YourCar(ActiveCar).Top = imgSquare(14).Top Then
        CarSet(ActiveCar).NextTopBlock = 8
        CarSet(ActiveCar).NextDownBlock = 20
    End If

    If YourCar(ActiveCar).Left = imgSquare(15).Left And YourCar(ActiveCar).Top = imgSquare(15).Top Then
        CarSet(ActiveCar).NextTopBlock = 9
        CarSet(ActiveCar).NextDownBlock = 21
    End If

    If YourCar(ActiveCar).Left = imgSquare(16).Left And YourCar(ActiveCar).Top = imgSquare(16).Top Then
        CarSet(ActiveCar).NextTopBlock = 10
        CarSet(ActiveCar).NextDownBlock = 22
    End If

    If YourCar(ActiveCar).Left = imgSquare(17).Left And YourCar(ActiveCar).Top = imgSquare(17).Top Then
        CarSet(ActiveCar).NextTopBlock = 11
        CarSet(ActiveCar).NextDownBlock = 23
    End If
    
    If YourCar(ActiveCar).Left = imgSquare(18).Left And YourCar(ActiveCar).Top = imgSquare(18).Top Then
        CarSet(ActiveCar).NextTopBlock = 12
        CarSet(ActiveCar).NextDownBlock = 24
    End If
    
'End of line 3 logics

    If YourCar(ActiveCar).Left = imgSquare(19).Left And YourCar(ActiveCar).Top = imgSquare(19).Top Then
        CarSet(ActiveCar).NextTopBlock = 13
        CarSet(ActiveCar).NextDownBlock = 25
    End If

    If YourCar(ActiveCar).Left = imgSquare(20).Left And YourCar(ActiveCar).Top = imgSquare(20).Top Then
        CarSet(ActiveCar).NextTopBlock = 14
        CarSet(ActiveCar).NextDownBlock = 26
    End If

    If YourCar(ActiveCar).Left = imgSquare(21).Left And YourCar(ActiveCar).Top = imgSquare(21).Top Then
        CarSet(ActiveCar).NextTopBlock = 15
        CarSet(ActiveCar).NextDownBlock = 27
    End If

    If YourCar(ActiveCar).Left = imgSquare(22).Left And YourCar(ActiveCar).Top = imgSquare(22).Top Then
        CarSet(ActiveCar).NextTopBlock = 16
        CarSet(ActiveCar).NextDownBlock = 28
    End If

    If YourCar(ActiveCar).Left = imgSquare(23).Left And YourCar(ActiveCar).Top = imgSquare(23).Top Then
        CarSet(ActiveCar).NextTopBlock = 17
        CarSet(ActiveCar).NextDownBlock = 29
    End If

    If YourCar(ActiveCar).Left = imgSquare(24).Left And YourCar(ActiveCar).Top = imgSquare(24).Top Then
        CarSet(ActiveCar).NextTopBlock = 18
        CarSet(ActiveCar).NextDownBlock = 30
    End If

'End of line 4 logics

    If YourCar(ActiveCar).Left = imgSquare(25).Left And YourCar(ActiveCar).Top = imgSquare(25).Top Then
        CarSet(ActiveCar).NextTopBlock = 19
        CarSet(ActiveCar).NextDownBlock = 31
    End If

    If YourCar(ActiveCar).Left = imgSquare(26).Left And YourCar(ActiveCar).Top = imgSquare(26).Top Then
        CarSet(ActiveCar).NextTopBlock = 20
        CarSet(ActiveCar).NextDownBlock = 32
    End If

    If YourCar(ActiveCar).Left = imgSquare(27).Left And YourCar(ActiveCar).Top = imgSquare(27).Top Then
        CarSet(ActiveCar).NextTopBlock = 21
        CarSet(ActiveCar).NextDownBlock = 33
    End If

    If YourCar(ActiveCar).Left = imgSquare(28).Left And YourCar(ActiveCar).Top = imgSquare(28).Top Then
        CarSet(ActiveCar).NextTopBlock = 22
        CarSet(ActiveCar).NextDownBlock = 34
    End If

    If YourCar(ActiveCar).Left = imgSquare(29).Left And YourCar(ActiveCar).Top = imgSquare(29).Top Then
        CarSet(ActiveCar).NextTopBlock = 23
        CarSet(ActiveCar).NextDownBlock = 35
    End If

    If YourCar(ActiveCar).Left = imgSquare(30).Left And YourCar(ActiveCar).Top = imgSquare(30).Top Then
        CarSet(ActiveCar).NextTopBlock = 24
        CarSet(ActiveCar).NextDownBlock = 36
    End If

    'Line 6 locgics is omitted because it is unreachable
    'END OF VERTICAL 2 LOGICS
End Sub

Private Sub HighLight()

Label9.Caption = ""


H2HL.Left = imgSquare(0).Left
H3HL.Left = imgSquare(0).Left
V2HL.Left = imgSquare(0).Left
V3HL.Left = imgSquare(0).Left

If CarSet(ActiveCar).Position = "Horizontal2" Then
H2HL.Left = YourCar(ActiveCar).Left - 30
H2HL.Top = YourCar(ActiveCar).Top - 30
End If

If CarSet(ActiveCar).Position = "Horizontal3" Then
H3HL.Left = YourCar(ActiveCar).Left - 30
H3HL.Top = YourCar(ActiveCar).Top - 30
End If

If CarSet(ActiveCar).Position = "Vertical2" Then
V2HL.Left = YourCar(ActiveCar).Left - 30
V2HL.Top = YourCar(ActiveCar).Top - 30
End If

If CarSet(ActiveCar).Position = "Vertical3" Then
V3HL.Left = YourCar(ActiveCar).Left - 30
V3HL.Top = YourCar(ActiveCar).Top - 30
End If

ShowControls

End Sub

Private Sub ShowControls()
CLeft.Left = imgSquare(0).Left
CRight.Left = imgSquare(0).Left
CDown.Left = imgSquare(0).Left
CUp.Left = imgSquare(0).Left

If CarSet(ActiveCar).Position = "Horizontal2" Then
    If Command1.Enabled = True Then
        CLeft.Left = YourCar(ActiveCar).Left + 90
        CLeft.Top = YourCar(ActiveCar).Top + 150
    End If
    
    If Command2.Enabled = True Then
        CRight.Left = YourCar(ActiveCar).Left + 870
        CRight.Top = YourCar(ActiveCar).Top + 150
    End If
End If

If CarSet(ActiveCar).Position = "Vertical2" Then
    If Command3.Enabled = True Then
        CUp.Left = YourCar(ActiveCar).Left + 180
        CUp.Top = YourCar(ActiveCar).Top + 150
    End If
    
    If Command4.Enabled = True Then
        CDown.Left = YourCar(ActiveCar).Left + 180
        CDown.Top = YourCar(ActiveCar).Top + 870
    End If
End If

If CarSet(ActiveCar).Position = "Horizontal3" Then
    If Command1.Enabled = True Then
        CLeft.Left = YourCar(ActiveCar).Left + 190
        CLeft.Top = YourCar(ActiveCar).Top + 150
    End If
    
    If Command2.Enabled = True Then
        CRight.Left = YourCar(ActiveCar).Left + 1590
        CRight.Top = YourCar(ActiveCar).Top + 150
    End If
End If

If CarSet(ActiveCar).Position = "Vertical3" Then
    If Command3.Enabled = True Then
        CUp.Left = YourCar(ActiveCar).Left + 210
        CUp.Top = YourCar(ActiveCar).Top + 150
    End If
    
    If Command4.Enabled = True Then
        CDown.Left = YourCar(ActiveCar).Left + 210
        CDown.Top = YourCar(ActiveCar).Top + 1590
    End If
End If
End Sub

Public Sub RemoveHighlight()
H2HL.Left = imgSquare(0).Left
H3HL.Left = imgSquare(0).Left
V2HL.Left = imgSquare(0).Left
V3HL.Left = imgSquare(0).Left
CLeft.Left = imgSquare(0).Left
CRight.Left = imgSquare(0).Left
CUp.Left = imgSquare(0).Left
CDown.Left = imgSquare(0).Left
End Sub

Private Sub DisplayStats()
Label6.Caption = "Level Pack: " & Left(LevelPack, Len(LevelPack) - 4)
Label3.Caption = "Moves: " & List2.ListCount
Label7.Caption = "Level Number: " & LevelToLoad
Label5.Caption = "Record Time: " & Format(CStr(((TargetTime \ 60) Mod 24)), "00") & ":" & Format(CStr(TargetTime Mod 60), "00")
Label4.Caption = "Record Moves: " & TargetMoves
Label8.Caption = "Selected Car: " & ActiveCar
End Sub

Public Sub ClearScore()
Label2.Caption = "Time: 00:00"
DrawLED TimeLeds, Mid(Label2.Caption, 7, 2), LED, 1
DrawLED SecondLeds, Mid(Label2.Caption, 10, 2), LED, 1
RemoveHighlight
For i = 1 To 29
    YourCar(i).Left = imgSquare(0).Left
Next i

Label6.Caption = "Level Pack: " & Left(LevelPack, Len(LevelPack) - 4)
Label3.Caption = "Moves: 0"
Label7.Caption = "Level Number: 0"
Label5.Caption = "Record Time: 00:00"
Label4.Caption = "Record Moves: 0"
Label8.Caption = "Selected Car: 0"

Timer1.Enabled = False
Timer2.Enabled = False

List2.Clear


End Sub

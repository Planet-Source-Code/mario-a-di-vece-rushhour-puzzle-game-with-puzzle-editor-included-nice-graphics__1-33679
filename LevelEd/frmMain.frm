VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rush Hour Level Editor"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9870
      TabIndex        =   42
      Text            =   "0"
      Top             =   5130
      Width           =   2145
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10710
      Top             =   3510
   End
   Begin VB.TextBox Text2 
      Height          =   3615
      Left            =   9870
      MultiLine       =   -1  'True
      TabIndex        =   41
      Top             =   5520
      Width           =   2145
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Default         =   -1  'True
      Height          =   465
      Left            =   10200
      TabIndex        =   40
      Top             =   2190
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   10410
      TabIndex        =   39
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Pos"
      Height          =   195
      Left            =   9810
      TabIndex        =   38
      Top             =   1770
      Width           =   270
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Block"
      Height          =   195
      Left            =   9810
      TabIndex        =   37
      Top             =   1500
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Car"
      Height          =   195
      Left            =   9810
      TabIndex        =   36
      Top             =   1230
      Width           =   240
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   1
      Left            =   7260
      Picture         =   "frmMain.frx":0000
      Top             =   150
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   2
      Left            =   6360
      Picture         =   "frmMain.frx":05B3
      Top             =   4380
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   3
      Left            =   7110
      Picture         =   "frmMain.frx":0C22
      Top             =   4380
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   4
      Left            =   300
      Picture         =   "frmMain.frx":1274
      Top             =   6270
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   5
      Left            =   3600
      Picture         =   "frmMain.frx":1B9B
      Top             =   6300
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   13
      Left            =   8820
      Picture         =   "frmMain.frx":254F
      Top             =   990
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   6
      Left            =   8070
      Picture         =   "frmMain.frx":2BE2
      Top             =   990
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   7
      Left            =   7260
      Picture         =   "frmMain.frx":3246
      Top             =   1020
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   8
      Left            =   6510
      Picture         =   "frmMain.frx":38B2
      Top             =   1020
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   9
      Left            =   8790
      Picture         =   "frmMain.frx":3F3F
      Top             =   2730
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   10
      Left            =   8040
      Picture         =   "frmMain.frx":457B
      Top             =   2730
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   11
      Left            =   7200
      Picture         =   "frmMain.frx":4C43
      Top             =   2790
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   1485
      Index           =   12
      Left            =   6420
      Picture         =   "frmMain.frx":514B
      Top             =   2790
      Width           =   675
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   29
      Left            =   7950
      Picture         =   "frmMain.frx":57CC
      Top             =   5220
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   14
      Left            =   6180
      Picture         =   "frmMain.frx":5D80
      Top             =   8490
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   15
      Left            =   6180
      Picture         =   "frmMain.frx":632F
      Top             =   7770
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   16
      Left            =   6180
      Picture         =   "frmMain.frx":6796
      Top             =   7050
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   17
      Left            =   6180
      Picture         =   "frmMain.frx":6D90
      Top             =   6330
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   18
      Left            =   7950
      Picture         =   "frmMain.frx":72F9
      Top             =   8520
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   19
      Left            =   7950
      Picture         =   "frmMain.frx":78D1
      Top             =   4410
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   20
      Left            =   7890
      Picture         =   "frmMain.frx":7E8E
      Top             =   7800
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   21
      Left            =   7890
      Picture         =   "frmMain.frx":8431
      Top             =   7050
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   660
      Index           =   22
      Left            =   7890
      Picture         =   "frmMain.frx":89C6
      Top             =   6300
      Width           =   1485
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   23
      Left            =   300
      Picture         =   "frmMain.frx":8FA4
      Top             =   8370
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   24
      Left            =   300
      Picture         =   "frmMain.frx":9887
      Top             =   7680
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   780
      Index           =   25
      Left            =   300
      Picture         =   "frmMain.frx":A3F7
      Top             =   6990
      Width           =   2355
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   26
      Left            =   4350
      Picture         =   "frmMain.frx":AD09
      Top             =   6300
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   27
      Left            =   2850
      Picture         =   "frmMain.frx":B6A9
      Top             =   6300
      Width           =   780
   End
   Begin VB.Image YourCar 
      Height          =   2355
      Index           =   28
      Left            =   5100
      Picture         =   "frmMain.frx":BFB6
      Top             =   6300
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "02"
      Height          =   195
      Index           =   35
      Left            =   1800
      TabIndex        =   35
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "03"
      Height          =   195
      Index           =   34
      Left            =   2610
      TabIndex        =   34
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "04"
      Height          =   195
      Index           =   33
      Left            =   3450
      TabIndex        =   33
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "05"
      Height          =   195
      Index           =   32
      Left            =   4230
      TabIndex        =   32
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "06"
      Height          =   195
      Index           =   31
      Left            =   5070
      TabIndex        =   31
      Top             =   960
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "07"
      Height          =   195
      Index           =   30
      Left            =   960
      TabIndex        =   30
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "08"
      Height          =   195
      Index           =   29
      Left            =   1800
      TabIndex        =   29
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      Height          =   195
      Index           =   28
      Left            =   2610
      TabIndex        =   28
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   195
      Index           =   27
      Left            =   3420
      TabIndex        =   27
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   195
      Index           =   26
      Left            =   4200
      TabIndex        =   26
      Top             =   1740
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   195
      Index           =   25
      Left            =   5070
      TabIndex        =   25
      Top             =   1770
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   195
      Index           =   24
      Left            =   990
      TabIndex        =   24
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      Height          =   195
      Index           =   23
      Left            =   1800
      TabIndex        =   23
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   195
      Index           =   22
      Left            =   2610
      TabIndex        =   22
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      Height          =   195
      Index           =   21
      Left            =   3450
      TabIndex        =   21
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      Height          =   195
      Index           =   20
      Left            =   4200
      TabIndex        =   20
      Top             =   2610
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      Height          =   195
      Index           =   19
      Left            =   5040
      TabIndex        =   19
      Top             =   2580
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      Height          =   195
      Index           =   18
      Left            =   1020
      TabIndex        =   18
      Top             =   3390
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   195
      Index           =   17
      Left            =   1830
      TabIndex        =   17
      Top             =   3390
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      Height          =   195
      Index           =   16
      Left            =   2640
      TabIndex        =   16
      Top             =   3390
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      Height          =   195
      Index           =   15
      Left            =   3420
      TabIndex        =   15
      Top             =   3360
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      Height          =   195
      Index           =   14
      Left            =   4260
      TabIndex        =   14
      Top             =   3420
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      Height          =   195
      Index           =   13
      Left            =   5010
      TabIndex        =   13
      Top             =   3360
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      Height          =   195
      Index           =   12
      Left            =   990
      TabIndex        =   12
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      Height          =   195
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      Height          =   195
      Index           =   10
      Left            =   2640
      TabIndex        =   10
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      Height          =   195
      Index           =   9
      Left            =   3420
      TabIndex        =   9
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      Height          =   195
      Index           =   8
      Left            =   4230
      TabIndex        =   8
      Top             =   4230
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      Height          =   195
      Index           =   7
      Left            =   5040
      TabIndex        =   7
      Top             =   4170
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      Height          =   195
      Index           =   6
      Left            =   990
      TabIndex        =   6
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      Height          =   195
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   5010
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      Height          =   195
      Index           =   4
      Left            =   2610
      TabIndex        =   4
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      Height          =   195
      Index           =   3
      Left            =   3450
      TabIndex        =   3
      Top             =   5010
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      Height          =   195
      Index           =   2
      Left            =   4230
      TabIndex        =   2
      Top             =   5010
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      Height          =   195
      Index           =   1
      Left            =   5010
      TabIndex        =   1
      Top             =   5010
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   195
      Index           =   0
      Left            =   990
      TabIndex        =   0
      Top             =   960
      Width           =   180
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   36
      Left            =   4770
      Picture         =   "frmMain.frx":C931
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   35
      Left            =   3960
      Picture         =   "frmMain.frx":CC71
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   19
      Left            =   720
      Picture         =   "frmMain.frx":CFB1
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   16
      Left            =   3150
      Picture         =   "frmMain.frx":D2F1
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   15
      Left            =   2340
      Picture         =   "frmMain.frx":D631
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   14
      Left            =   1530
      Picture         =   "frmMain.frx":D971
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   13
      Left            =   720
      Picture         =   "frmMain.frx":DCB1
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   12
      Left            =   4770
      Picture         =   "frmMain.frx":DFF1
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   11
      Left            =   3960
      Picture         =   "frmMain.frx":E331
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   10
      Left            =   3150
      Picture         =   "frmMain.frx":E671
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   9
      Left            =   2340
      Picture         =   "frmMain.frx":E9B1
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   8
      Left            =   1530
      Picture         =   "frmMain.frx":ECF1
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   7
      Left            =   720
      Picture         =   "frmMain.frx":F031
      Top             =   1500
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   6
      Left            =   4770
      Picture         =   "frmMain.frx":F371
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   5
      Left            =   3960
      Picture         =   "frmMain.frx":F6B1
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   4
      Left            =   3150
      Picture         =   "frmMain.frx":F9F1
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   3
      Left            =   2340
      Picture         =   "frmMain.frx":FD31
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   2
      Left            =   1530
      Picture         =   "frmMain.frx":10071
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   1
      Left            =   720
      Picture         =   "frmMain.frx":103B1
      Top             =   690
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   18
      Left            =   4770
      Picture         =   "frmMain.frx":106F1
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   17
      Left            =   3960
      Picture         =   "frmMain.frx":10A31
      Top             =   2310
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   34
      Left            =   3150
      Picture         =   "frmMain.frx":10D71
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   33
      Left            =   2340
      Picture         =   "frmMain.frx":110B1
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   32
      Left            =   1530
      Picture         =   "frmMain.frx":113F1
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   31
      Left            =   720
      Picture         =   "frmMain.frx":11731
      Top             =   4740
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   30
      Left            =   4770
      Picture         =   "frmMain.frx":11A71
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   29
      Left            =   3960
      Picture         =   "frmMain.frx":11DB1
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   28
      Left            =   3150
      Picture         =   "frmMain.frx":120F1
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   27
      Left            =   2340
      Picture         =   "frmMain.frx":12431
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   26
      Left            =   1530
      Picture         =   "frmMain.frx":12771
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   25
      Left            =   720
      Picture         =   "frmMain.frx":12AB1
      Top             =   3930
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   24
      Left            =   4770
      Picture         =   "frmMain.frx":12DF1
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   23
      Left            =   3960
      Picture         =   "frmMain.frx":13131
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   22
      Left            =   3150
      Picture         =   "frmMain.frx":13471
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   21
      Left            =   2340
      Picture         =   "frmMain.frx":137B1
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image imgSquare 
      Height          =   750
      Index           =   20
      Left            =   1530
      Picture         =   "frmMain.frx":13AF1
      Top             =   3120
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   6150
      Left            =   0
      Picture         =   "frmMain.frx":13E31
      Top             =   0
      Width           =   6240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActiveCar As Integer

Private Sub Command1_Click()
Text3.Text = Val(Text3.Text) + 1
YourCar(ActiveCar).Left = imgSquare(Val(Text1.Text)).Left
YourCar(ActiveCar).Top = imgSquare(Val(Text1.Text)).Top
Text2.Text = Text2.Text & vbCrLf & Label2.Caption & vbCrLf & Text1.Text & vbCrLf & Label4.Caption
Text1.Text = ""


End Sub

Private Sub Timer1_Timer()
Label2.Caption = ActiveCar
Label4.Caption = CarSet(ActiveCar).Position
End Sub

Private Sub YourCar_Click(Index As Integer)
Select Case Index
    Case 1
    ActiveCar = 1
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H2"

    Case 2
    ActiveCar = 2
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 3
    ActiveCar = 3
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 4
    ActiveCar = 4
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H3"

    Case 5
    ActiveCar = 5
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V3"

    Case 6
    ActiveCar = 6
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "V2"

    Case 7
    ActiveCar = 7
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 8
    ActiveCar = 8
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 9
    ActiveCar = 9
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "V2"

    Case 10
    ActiveCar = 10
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 11
    ActiveCar = 11
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "V2"

    Case 12
    ActiveCar = 12
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 13
    ActiveCar = 13
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V2"

    Case 14
    ActiveCar = 14
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H2"

    Case 15
    ActiveCar = 15
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 16
    ActiveCar = 16
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H2"

    Case 17
    ActiveCar = 17
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 18
    ActiveCar = 18
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 19
    ActiveCar = 19
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H2"

    Case 20
    ActiveCar = 20
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 21
    ActiveCar = 21
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 22
    ActiveCar = 22
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H2"

    Case 23
    ActiveCar = 23
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H3"

    Case 24
    ActiveCar = 24
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H3"

    Case 25
    ActiveCar = 25
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "H3"

    Case 26
    ActiveCar = 26
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "V3"

    Case 27
    ActiveCar = 27
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V3"

    Case 28
    ActiveCar = 28
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "UP-Down"
    CarSet(ActiveCar).Position = "V3"

    Case 29
    ActiveCar = 29
    CarSet(ActiveCar).CarID = ActiveCar
    CarSet(ActiveCar).Movement = "Left-Right"
    CarSet(ActiveCar).Position = "H2"

End Select

Text1.SetFocus

End Sub

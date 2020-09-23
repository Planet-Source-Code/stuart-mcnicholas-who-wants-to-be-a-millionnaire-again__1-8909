VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " WHO WANTS TO BE A MILLIONNAIRE"
   ClientHeight    =   8595
   ClientLeft      =   525
   ClientTop       =   330
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11160
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "TAKE THE MONEY"
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      Top             =   8160
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1920
      Top             =   4320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "50 / 50"
      Height          =   735
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PHONE A FRIEND"
      Height          =   735
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ASK THE AUDIANCE"
      Height          =   735
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7200
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00808000&
      Caption         =   "WEST HAM"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   6
      Top             =   4080
      Width           =   3015
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00808000&
      Caption         =   "MAN UTD"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   5
      Top             =   5520
      Width           =   3015
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00808000&
      Caption         =   "ARSENAL"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6960
      TabIndex        =   4
      Top             =   5520
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00808000&
      Caption         =   "LIVERPOOL"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3240
      TabIndex        =   3
      Top             =   4080
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar PROGRESS 
      Height          =   6615
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   11668
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   15
      Orientation     =   1
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      Height          =   495
      Left            =   10320
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   735
      Left            =   7560
      TabIndex        =   16
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   1215
      Left            =   6480
      TabIndex        =   15
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   1455
      Left            =   2640
      TabIndex        =   14
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.OLE OLE1 
      Class           =   "Package"
      Height          =   255
      Left            =   600
      SourceDoc       =   "C:\WINDOWS\Profiles\stuart\Desktop\enviroment[1].mp3"
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      Height          =   495
      Left            =   4080
      Top             =   8040
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   4920
      TabIndex        =   10
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   975
      Left            =   3480
      Top             =   7080
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   2775
      Left            =   3120
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WHICH FOOTBALL TEAM WON THE         PREMIER LEAGUE 98/99 ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WHO WANTS TO BE A                  MILLIONNAIRE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   9495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QUESTION

Private Sub Command1_Click()
Load ASK
ASK.Show
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
PHONE.Show
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
MsgBox "COMPUTER PLEASE TAKE TWO OF THE WRONG ANSWERS AWAY", vbInformation, "50 / 50"
Command3.Enabled = False
If QUESTION = 1 Then
Option2.Visible = False
Option1.Visible = False
End If

If QUESTION = 2 Then
Option2.Visible = False
Option4.Visible = False
End If

If QUESTION = 3 Then
Option2.Visible = False
Option1.Visible = False
End If

If QUESTION = 4 Then
Option4.Visible = False
Option1.Visible = False
End If

If QUESTION = 5 Then
Option3.Visible = False
Option1.Visible = False
End If

If QUESTION = 6 Then
Option3.Visible = False
Option4.Visible = False
End If

If QUESTION = 7 Then
Option2.Visible = False
Option1.Visible = False
End If

If QUESTION = 8 Then
Option2.Visible = False
Option3.Visible = False
End If

If QUESTION = 9 Then
Option2.Visible = False
Option3.Visible = False
End If

If QUESTION = 10 Then
Option4.Visible = False
Option1.Visible = False
End If

If QUESTION = 11 Then
Option2.Visible = False
Option1.Visible = False
End If

If QUESTION = 12 Then
Option2.Visible = False
Option3.Visible = False
End If

If QUESTION = 13 Then
Option4.Visible = False
Option1.Visible = False
End If

If QUESTION = 14 Then
Option2.Visible = False
Option1.Visible = False
End If

If QUESTION = 15 Then
Option2.Visible = False
Option4.Visible = False
End If


End Sub

Private Sub Command4_Click()
MsgBox "CONGRATULATIONS YOU LEAVE HERE TONIGHT WITH '" & Label3.Caption & "' POUNDS", vbInformation, "WHO WANTS TO BE A MILLIONNAIRE"
End
End Sub

Private Sub Form_Load()
QUESTION = 1
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
PROGRESS.Value = 0
End Sub



Private Sub Label4_Click(Index As Integer)
OLE1.DoVerb
End Sub

Private Sub Label5_Click()
OLE1.DoVerb
End Sub

Private Sub Label6_Click()
OLE1.DoVerb
End Sub

Private Sub Label7_Click()
OLE1.DoVerb
End Sub

Private Sub Label8_Click()
OLE1.DoVerb
End Sub

Private Sub Label9_Click()
OLE1.DoVerb
End Sub

Private Sub Option1_Click()
Option2.Visible = True
Option1.Visible = True
Option3.Visible = True
Option4.Visible = True
'..............

If QUESTION = 2 Then
PROGRESS.Value = 2
MsgBox "CORRECT, YOU HAVE WON 200 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 3
Option1.Value = False
Label2.Caption = "WHAT DO THE COMPANY SENSODYNE PRODUCE"
Option1.Caption = "HAIR PRODUCTS"
Option2.Caption = "NAIL CLIPPERS"
Option3.Caption = "NAIL VARNISH"
Option4.Caption = "TOOTHPASTE"
Label3.Caption = "400"
Exit Sub
End If


If QUESTION = 8 Then
PROGRESS.Value = 8
MsgBox "CORRECT, YOU HAVE WON 8,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 9
Option1.Value = False
Label2.Caption = "HOW MANY HOURS ARE LONDON INFRONT OF NEW YORK"
Option1.Caption = "2"
Option2.Caption = "10"
Option3.Caption = "24"
Option4.Caption = "5"
Label3.Caption = "16,000"
Exit Sub
End If


If QUESTION = 12 Then
PROGRESS.Value = 12
MsgBox "CORRECT, YOU HAVE WON 125,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 13
Option1.Value = False
Label2.Caption = "WHAT COMPANY MAKES THE A-K 47"
Option1.Caption = "H&k"
Option2.Caption = "KLASHNOCOV"
Option3.Caption = "RESTACOV"
Option4.Caption = "COLT"
Label3.Caption = "250,000"
Exit Sub
End If

If QUESTION = 1 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 3 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 4 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 5 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 6 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 7 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 9 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 10 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 11 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 13 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 14 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 15 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
End Sub

Private Sub Option2_Click()
Option2.Visible = True
Option1.Visible = True
Option3.Visible = True
Option4.Visible = True
'......................
If QUESTION = 5 Then
PROGRESS.Value = 5
MsgBox "CORRECT, YOU HAVE WON 1,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 6
Option2.Value = False
Label2.Caption = "WHAT CITY WAS 'NOEL GALLAGHER' BORN IN"
Option1.Caption = "LIVERPOOL"
Option2.Caption = "MANCHESTER"
Option3.Caption = "NEWCASTLE"
Option4.Caption = "BOSTON"
Label3.Caption = "2,000"
Exit Sub
End If


If QUESTION = 6 Then
PROGRESS.Value = 6
MsgBox "CORRECT, YOU HAVE WON 2,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 7
Option2.Value = False
Label2.Caption = "WHICH INTERNATIONAL CRICKET TEAM PLAYS AT THE 'S.C.G' CRICKET GROUND"
Option1.Caption = "SOUTH AFRICA"
Option2.Caption = "NEW ZEALAND"
Option3.Caption = "ENGLAND"
Option4.Caption = "AUSTRALLIA"
Label3.Caption = "4,000"
Exit Sub
End If


If QUESTION = 13 Then
PROGRESS.Value = 13
MsgBox "CORRECT, YOU HAVE WON 250,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 14
Option1.Value = False
Label2.Caption = "WHAT IS A BILL HOOK USED FOR"
Option1.Caption = "LIFTING A BUCKET FROM A WELL"
Option2.Caption = "HANGING BUTCHERS MEAT"
Option3.Caption = "TORCHER"
Option4.Caption = "BLOCKING SWORDS"
Label3.Caption = "500,000"
Exit Sub
End If

If QUESTION = 1 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 2 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 3 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 4 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 7 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 8 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 9 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 10 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 11 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 12 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 14 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 15 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
End Sub

Private Sub Option3_Click()
Option2.Visible = True
Option1.Visible = True
Option3.Visible = True
Option4.Visible = True
'............................
If QUESTION = 1 Then
PROGRESS.Value = 1
MsgBox "CORRECT, YOU HAVE WON 100 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 2
Option3.Value = False
Label2.Caption = "WHO SCORED THE LAST GOAL IN THE WORLD CUP IN FRANCE"
Option1.Caption = "EMMANUEL PETTIT"
Option2.Caption = "RYAN GIGGS"
Option3.Caption = "PATRICK VEIRA"
Option4.Caption = "ALAN SHEARER"
Label3.Caption = "200"
Exit Sub
End If

If QUESTION = 4 Then
PROGRESS.Value = 4
MsgBox "CORRECT, YOU HAVE WON 500 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 5
Option3.Value = False
Label2.Caption = "IN COMPUTER TERMS WHAT DOES 'H.C.I' STAND FOR"
Label2.FontSize = 16
Option1.Caption = "HUMAN CONTROL INTERFACE"
Option2.Caption = "HUMAN COMPUTER INTERFACE"
Option3.Caption = "HUMAN CONSOLE INTERFACE"
Option4.Caption = "HACKERS CONTROL INTERFACE"
Label3.Caption = "1,000"
Exit Sub
End If


If QUESTION = 10 Then
PROGRESS.Value = 10
MsgBox "CORRECT, YOU HAVE WON 32,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 11
Option3.Value = False
Label2.Caption = "WHAT IS THE ONLY VISIBLE MAN-MADE OBJECT FROM SPACE"
Option1.Caption = "COVENTRY CATHEDRAL"
Option2.Caption = "BERLIN WALL"
Option3.Caption = "EMPIRE STATE BUILDING"
Option4.Caption = "THE GREAT WALL OF CHINA"
Label3.Caption = "64,000"
Exit Sub
End If

If QUESTION = 15 Then
PROGRESS.Value = 15
MsgBox "CORRECT, YOU HAVE WON 1,000,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 16
Option1.Value = False
Label2.Caption = ""
Option1.Caption = ""
Option2.Caption = ""
Option3.Caption = ""
Option4.Caption = ""
MsgBox "WELL DONE YOU WIN, BUT YOU DONT GET ANY MONEY FOR IT"
End
Exit Sub
End If



If QUESTION = 2 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 3 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 5 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 6 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 7 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 8 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 9 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 11 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 12 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 13 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 14 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If

End Sub

Private Sub Option4_Click()
Option2.Visible = True
Option1.Visible = True
Option3.Visible = True
Option4.Visible = True
'............................
If QUESTION = 3 Then
PROGRESS.Value = 3
MsgBox "CORRECT, YOU HAVE WON 400 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 4
Option4.Value = False
Label2.Caption = "WHICH AMERICAN SITCOM FEATURES CHARACTERS SUCH AS MONICA, RACHEL, ROSS, JOEY"
Label2.FontSize = 14
Option1.Caption = "E.R"
Option2.Caption = "ROYAL FAMILY"
Option3.Caption = "FRIENDS"
Option4.Caption = "FRASIER"
Label3.Caption = "500"
Exit Sub
End If


If QUESTION = 7 Then
PROGRESS.Value = 7
MsgBox "CORRECT, YOU HAVE WON 4,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 8
Option4.Value = False
Label2.Caption = "IN WHICH TOWN WAS THE CAPTAIN OF THE TITANIC BORN IN"
Option1.Caption = "HANLEY"
Option2.Caption = "NEWCASTLE UNDER-LYME"
Option3.Caption = "FENTON"
Option4.Caption = "STONE"
Label3.Caption = "8,000"
Exit Sub
End If


If QUESTION = 9 Then
PROGRESS.Value = 9
MsgBox "CORRECT, YOU HAVE WON 16,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 10
Option4.Value = False
Label2.Caption = "WHO WAS THE FIRST SNOOKER PLAYER TO MAKE A '147' ON T.V"
Option1.Caption = "STEVEN HENDRY"
Option2.Caption = "RONNIE O'SULLIVAN"
Option3.Caption = "STEVE DAVIS"
Option4.Caption = "LUKE YOUNG"
Label3.Caption = "32,000"
Exit Sub
End If


If QUESTION = 11 Then
PROGRESS.Value = 11
MsgBox "CORRECT, YOU HAVE WON 64,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 12
Option4.Value = False
Label2.Caption = "WHO WAS THE OTHER CO-FOUNDER OF MICROSOFT BESIDES BILL GATES "
Option1.Caption = "PAUL ALLEN"
Option2.Caption = "STEVE JOBS"
Option3.Caption = "PETER MCCONIVILLE"
Option4.Caption = "PETER STERN"
Label3.Caption = "125,000"
Exit Sub
End If


If QUESTION = 14 Then
PROGRESS.Value = 14
MsgBox "CORRECT, YOU HAVE WON 500,000 POUNDS", vbInformation, "MILIONNAIRE"
QUESTION = 15
Option1.Value = False
Label2.Caption = "WHAT WAS THE ENIGMA CODE"
Option1.Caption = "AN AMERICAN CODE"
Option2.Caption = "A EUROPEAN CODE"
Option3.Caption = "THE GERMAN CODE"
Option4.Caption = "THE RUSIAN CODE"
Label3.Caption = "1,000,000"
Exit Sub
End If


If QUESTION = 1 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 2 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 4 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 5 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 6 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 8 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 10 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 12 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 13 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
If QUESTION = 15 Then
MsgBox "UNLUCKY YOUR SHIT"
End
End If
End Sub

Private Sub Timer1_Timer()
If Label3.Visible = True Then
Label3.Visible = False
Exit Sub
End If

If Label3.Visible = False Then
Label3.Visible = True
Exit Sub
End If
End Sub

VERSION 5.00
Begin VB.Form START 
   BackColor       =   &H00000000&
   Caption         =   "WHO WANTS TO BE A MILLIONNAIRE"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7050
   Icon            =   "START.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5010
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10200
      Picture         =   "START.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   8040
   End
   Begin VB.Timer Timer2 
      Left            =   6000
      Top             =   7080
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   4080
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   "CLICK ME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   2
      Left            =   7440
      Picture         =   "START.frx":0884
      Stretch         =   -1  'True
      Top             =   -200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   1
      Left            =   4440
      Picture         =   "START.frx":FE32
      Stretch         =   -1  'True
      Top             =   -200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Index           =   0
      Left            =   1440
      Picture         =   "START.frx":106C3
      Stretch         =   -1  'True
      Top             =   -200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "TRY TO WIN IF YOU CAN"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   24
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   615
      Left            =   4320
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "PMingLiU"
         Size            =   399.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   4440
      TabIndex        =   14
      Top             =   720
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   11760
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1335
      Index           =   12
      Left            =   -720
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   11
      Left            =   -1680
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   10
      Left            =   -2160
      TabIndex        =   11
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   9
      Left            =   -3000
      TabIndex        =   10
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   8
      Left            =   -3840
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   7
      Left            =   -4680
      TabIndex        =   8
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   6
      Left            =   15720
      TabIndex        =   7
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   5
      Left            =   15240
      TabIndex        =   6
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   4
      Left            =   14400
      TabIndex        =   5
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   3
      Left            =   13560
      TabIndex        =   4
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Index           =   2
      Left            =   13080
      TabIndex        =   3
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1335
      Index           =   1
      Left            =   11880
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   2055
      Left            =   6120
      Top             =   9000
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   4
      Height          =   2055
      Left            =   1080
      Top             =   9000
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MCPOWER COMPUTERS           PRESENTS"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2055
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   10455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "WHO WANTS TO BE  A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1080
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   9855
   End
End
Attribute VB_Name = "START"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPVAL, PROCED, WIDTHVAL
Dim LEFTVAL
Dim X As Integer
Dim PROCED2

Private Sub Command1_Click()
Load Form1
Form1.Visible = True
START.Visible = False
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Load Form1
Form1.Visible = True
START.Visible = False
End If
End Sub

Private Sub Form_Load()
TOPVAL = 150
PROCED = 0
WIDTHVAL = 80
PROCED2 = 1
End Sub

Private Sub Label6_Click()
Label7.Visible = True
Label7.BackColor = &H8000000F
End Sub

Private Sub Label7_Click()
For X = 0 To 2
Image1(X).Visible = True
Image1(X).Height = Image1(X).Height + 100
Next X
End Sub

Private Sub Timer1_Timer()

Shape1.Top = Shape1.Top - TOPVAL
Shape2.Top = Shape2.Top - TOPVAL
If Shape1.Top < 2480 Then
TOPVAL = 0
PROCED = 1
End If

If PROCED = 1 Then
Label2.Visible = True
Shape1.Width = Shape1.Width - WIDTHVAL
Shape2.Width = Shape2.Width - WIDTHVAL
Shape2.Left = Shape2.Left + 80
If Shape1.Width < 60 Then
Shape1.Visible = False
Shape2.Visible = False
WIDTHVAL = 0
PROCED = 2
End If
End If

If PROCED = 2 Then
For X = 1 To 6
Label3(X).Left = Label3(X).Left - 130
Next X

For X = 7 To 12
Label3(X).Left = Label3(X).Left + 130
Next X
End If

If Label3(11).Left > 9000 Then
Timer1.Interval = 0
Timer2.Interval = 1
End If

End Sub

Private Sub Timer2_Timer()
If PROCED2 = 1 Then
Label3(12).Visible = True
Label3(1).Visible = True
Label4.Height = Label4.Height + 100
If Label4.Height > 7335 Then
Timer3.Interval = 700
Timer2.Interval = 0
End If
End If



End Sub

Private Sub Timer3_Timer()
Command1.Visible = True
If Label5.Visible = False Then
Label5.Visible = True
Exit Sub
End If

If Label5.Visible = True Then
Label5.Visible = False
Exit Sub
End If
End Sub


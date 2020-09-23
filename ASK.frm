VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ASK 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "ASK THE AUDIENCE"
   ClientHeight    =   6000
   ClientLeft      =   3675
   ClientTop       =   1575
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      Height          =   315
      Left            =   600
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ASK THE AUDIENCE"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5640
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   4335
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7646
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   4335
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7646
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   4335
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7646
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   4335
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   7646
      _Version        =   393216
      Appearance      =   1
      Max             =   200
      Orientation     =   1
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   6
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   600
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "  B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "ASK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Visible = True
Randomize
pb(Index).Max = 100
pb(0).Value = Int((1 * 100) * Rnd)
pb(1).Value = Int((1 * 100) * Rnd)
pb(2).Value = Int((1 * 100) * Rnd)
pb(3).Value = Int((1 * 100) * Rnd)
If pb(0) + pb(1) + pb(2) + pb(3) > 100 Then
Call Command1_Click
End If
If pb(0) + pb(1) + pb(2) + pb(3) < 100 Then
Call Command1_Click
End If
Label2.Caption = pb(0).Value
Label3.Caption = pb(1).Value
Label4.Caption = pb(2).Value
Label5.Caption = pb(3).Value
End Sub

Private Sub Timer1_Timer()
Call Command1_Click
End Sub


Private Sub Command2_Click()
ASK.Visible = False
Load Form1
Form1.Visible = True
End Sub

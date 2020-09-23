VERSION 5.00
Begin VB.Form PHONE 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "PHONE A FRIEND"
   ClientHeight    =   4680
   ClientLeft      =   2445
   ClientTop       =   1935
   ClientWidth     =   7440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Left            =   6840
      Top             =   1440
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      Caption         =   "BACK"
      Enabled         =   0   'False
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
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RING"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..................................................................ITS OPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "POBABLY OPTION 2 BUT IM NOT SURE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I HAVE NO IDEA,.............. SORRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEFINATLY OPTION 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ERM I THINK ITS OPTION1 OR OPTION2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2055
      Left            =   240
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   TEL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "PHONE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyValue

Private Sub Command1_Click()
Command1.Enabled = False
Timer1.Interval = 0
Timer2.Interval = 1
End Sub

Private Sub Command2_Click()
PHONE.Visible = False
End Sub

Private Sub Text2_Change()
Command1.Visible = True
End Sub

Private Sub Timer1_Timer()
Randomize

MyValue = Int((4 * Rnd) + 1)
Text3.Text = MyValue
End Sub

Private Sub Timer2_Timer()
If Text3.Text = 1 Then
Label3.Visible = True
Label3.Width = Label3.Width + 90
End If

    If Text3.Text = 2 Then
    Label4.Visible = True
    Label4.Width = Label4.Width + 90
    End If

        If Text3.Text = 3 Then
        Label5.Visible = True
        Label5.Width = Label5.Width + 90
        End If

            If Text3.Text = 4 Then
            Label6.Visible = True
            Label6.Width = Label6.Width + 90
            End If

                If Text3.Text = 5 Then
                Label7.Visible = True
                Label7.Width = Label7.Width + 90
                End If
                
If Label3.Width > 6735 Then
Timer2.Interval = 0
Command2.Enabled = True
End If
If Label4.Width > 6735 Then
Timer2.Interval = 0
Command2.Enabled = True
End If
If Label5.Width > 6735 Then
Timer2.Interval = 0
Command2.Enabled = True
End If
If Label6.Width > 6735 Then
Timer2.Interval = 0
Command2.Enabled = True
End If
If Label7.Width > 6735 Then
Timer2.Interval = 0
Command2.Enabled = True
End If


End Sub

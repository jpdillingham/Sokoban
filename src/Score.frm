VERSION 5.00
Begin VB.Form Score 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sokoban - Completed level"
   ClientHeight    =   4140
   ClientLeft      =   5655
   ClientTop       =   3435
   ClientWidth     =   4905
   Icon            =   "Score.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Dont keep score."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   375
      TabIndex        =   19
      Top             =   3885
      Width           =   4440
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Always record my score."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   375
      TabIndex        =   18
      Top             =   3630
      Width           =   4440
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Always record my score (if better than previous best)."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   17
      Top             =   3375
      Value           =   -1  'True
      Width           =   4440
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stop asking me, and"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   16
      Top             =   3075
      Width           =   4740
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2550
      TabIndex        =   7
      Top             =   2625
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Yes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1425
      TabIndex        =   6
      Top             =   2625
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   1725
      TabIndex        =   15
      Top             =   1875
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   1725
      TabIndex        =   14
      Top             =   1650
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Previous best:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   225
      TabIndex        =   13
      Top             =   1875
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Total pushes:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   12
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1725
      TabIndex        =   11
      Top             =   1275
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1725
      TabIndex        =   10
      Top             =   1050
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Previous best:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   225
      TabIndex        =   9
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Total moves:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   225
      TabIndex        =   8
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Would you like to record your score?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   2250
      Width           =   4740
   End
   Begin VB.Label Label2 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1725
      TabIndex        =   4
      Top             =   675
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1725
      TabIndex        =   3
      Top             =   450
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Previous best:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   675
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "Your time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Congratulations!  You have successfully completed this level."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4740
   End
End
Attribute VB_Name = "Score"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intAction As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Option1(0).Enabled = True
        Option1(1).Enabled = True
        Option1(2).Enabled = True
    Else
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Option1(2).Enabled = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    LoadLevelIni strLevelFileName, intLevel + 1
End Sub

Private Sub cmdOK_Click()
    Dim strTemp As String
    If Check1.Value = 1 Then
        strTemp = Val(intAction)
        WriteIni "Options", "ScorePrompt", "0", App.Path & "\Sokoban.ini"
        WriteIni "Options", "ScoreAction", strTemp, App.Path & "\Sokoban.ini"
    End If
    SaveScore Val(intLevel)
    Unload Me
    LoadLevelIni strLevelFileName, intLevel + 1
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Label2(2) = strFormattedScoreTime
    Label2(3) = strBestTime
    Label2(6) = intMoves
    If intBestMoves = 0 Then
        Label2(7) = "N\A"
    Else
        Label2(7) = intBestMoves
    End If
    Label2(10) = intPushes
    If intBestPushes = 0 Then
        Label2(11) = "N\A"
    Else
        Label2(11) = intBestPushes
    End If
    intAction = 0
End Sub

Private Sub Option1_Click(Index As Integer)
    intAction = Index
End Sub

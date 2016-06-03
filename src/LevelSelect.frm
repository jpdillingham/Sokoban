VERSION 5.00
Begin VB.Form LevelSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sokoban - Select level"
   ClientHeight    =   900
   ClientLeft      =   4680
   ClientTop       =   3345
   ClientWidth     =   2310
   Icon            =   "LevelSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   1350
      TabIndex        =   4
      Top             =   525
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
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
      Left            =   375
      TabIndex        =   3
      Top             =   525
      Width           =   915
   End
   Begin VB.ComboBox cmbLevel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1050
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   75
      Width           =   1215
   End
   Begin VB.ComboBox cmbSeries 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2325
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "Go to level:"
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
      TabIndex        =   5
      Top             =   150
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Series:"
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
      Left            =   150
      TabIndex        =   1
      Top             =   2400
      Width           =   765
   End
End
Attribute VB_Name = "LevelSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    Dim strTemp As String
    strTemp = cmbLevel.Text
    Unload Me
    LoadLevelIni strLevelFileName, strTemp
End Sub

Private Sub Form_Load()
    Dim i As Integer, strTemp As String
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Caption = "Sokoban - Select level"
    i = 1
    Do Until i > intNumLevels
        cmbLevel.AddItem Val(i)
        i = i + 1
    Loop
    cmbLevel.Text = intLevel
End Sub

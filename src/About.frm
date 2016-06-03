VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sokoban - About"
   ClientHeight    =   1635
   ClientLeft      =   6165
   ClientTop       =   5100
   ClientWidth     =   3690
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
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
      Left            =   2700
      TabIndex        =   2
      Top             =   1275
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Sokoban 1.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   975
      TabIndex        =   1
      Top             =   75
      Width           =   2640
   End
   Begin VB.Label lblAbout 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   975
      TabIndex        =   0
      Top             =   300
      Width           =   2640
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   375
      Picture         =   "About.frx":08CA
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   75
      Picture         =   "About.frx":1194
      Top             =   75
      Width           =   810
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    lblAbout = "A Puzzle\Strategy Game" & vbCrLf & vbCrLf & "Written in Microsoft Visual Basic 6.0" & vbCrLf & "By JP Dillingham (praetorian)"
End Sub

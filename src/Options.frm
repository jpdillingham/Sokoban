VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sokoban - Options"
   ClientHeight    =   2775
   ClientLeft      =   3825
   ClientTop       =   3075
   ClientWidth     =   5115
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "When starting,"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   175
      TabIndex        =   7
      Top             =   450
      Width           =   4740
      Begin VB.OptionButton optStart 
         Caption         =   "Load a specified series and level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   9
         Top             =   750
         Width           =   3390
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
         IntegralHeight  =   0   'False
         Left            =   375
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   975
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.OptionButton optStart 
         Caption         =   "Load the first level found"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   525
         Width           =   3390
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
         IntegralHeight  =   0   'False
         Left            =   2775
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   975
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton optStart 
         Caption         =   "Load the last saved game"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Value           =   -1  'True
         Width           =   3390
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "When I complete a level,"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   175
      TabIndex        =   4
      Top             =   450
      Width           =   4740
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   375
         ScaleHeight     =   915
         ScaleWidth      =   4290
         TabIndex        =   13
         Top             =   750
         Width           =   4290
         Begin VB.OptionButton optScoreOpt 
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
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Value           =   -1  'True
            Width           =   4215
         End
         Begin VB.OptionButton optScoreOpt 
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
            Left            =   0
            TabIndex        =   15
            Top             =   225
            Width           =   3390
         End
         Begin VB.OptionButton optScoreOpt 
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
            Left            =   0
            TabIndex        =   14
            Top             =   450
            Width           =   3090
         End
      End
      Begin VB.OptionButton optScore 
         Caption         =   "Ask me what to do"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   3390
      End
      Begin VB.OptionButton optScore 
         Caption         =   "Never ask, and"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   525
         Width           =   3390
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   175
      ScaleHeight     =   1740
      ScaleWidth      =   4740
      TabIndex        =   17
      Top             =   450
      Width           =   4740
      Begin VB.CheckBox Check1 
         Caption         =   "Automatically save after each completed level"
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
         TabIndex        =   22
         Top             =   1125
         Width           =   4440
      End
      Begin VB.Frame Frame3 
         Caption         =   "When exiting Sokoban,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   4740
         Begin VB.OptionButton optExit 
            Caption         =   "Dont save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   21
            Top             =   750
            Width           =   3990
         End
         Begin VB.OptionButton optExit 
            Caption         =   "Automatically save my game"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   20
            Top             =   525
            Width           =   3990
         End
         Begin VB.OptionButton optExit 
            Caption         =   "Ask me if i want to save (if i havent already)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   19
            Top             =   300
            Width           =   3990
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
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
      Height          =   300
      Left            =   4125
      TabIndex        =   2
      Top             =   2400
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
      Left            =   2175
      TabIndex        =   1
      Top             =   2400
      Width           =   915
   End
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
      Left            =   3150
      TabIndex        =   0
      Top             =   2400
      Width           =   915
   End
   Begin MSComctlLib.TabStrip TabStrip 
      Height          =   2220
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   3916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Saving"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Scoring"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intOptionLoad As Integer
Dim intOptionExit As Integer
Dim intOptionPrompt As Integer
Dim intOptionScore As Integer

Private Sub Check1_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmbSeries_Click()
    Dim strTemp As String
    cmbLevel.Clear
    strTemp = String(10, " ")
    ReadIni "Info", "Levels", vbNullString, strTemp, 10, App.Path & "\levels\" & cmbSeries.Text & ".sok"
    intI = 1
    Do Until intI > Val(strTemp)
        cmbLevel.AddItem Val(intI)
        intI = intI + 1
    Loop
    cmbLevel.Text = "1"
End Sub

Private Sub cmdApply_Click()
    SaveSettings
    cmdApply.Enabled = False
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveSettings
    Unload Me
End Sub

Private Sub Form_Load()
    Dim intI As Integer, strTemp As String
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    Main.flbSeries.Path = App.Path & "\levels\"
    Do Until Main.flbSeries.List(intI) = ""
        cmbSeries.AddItem Split(Main.flbSeries.List(intI), ".")(0)
        intI = intI + 1
    Loop
    cmbSeries.Text = cmbSeries.List(0)
    strTemp = String(10, " ")
    ReadIni "Info", "Levels", vbNullString, strTemp, 10, App.Path & "\levels\" & cmbSeries.Text & ".sok"
    intI = 1
    Do Until intI > Val(strTemp)
        cmbLevel.AddItem Val(intI)
        intI = intI + 1
    Loop
    cmbLevel.Text = 1
    strTemp = String(10, " ")
    ReadIni "Options", "ScorePrompt", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    optScore(Val(strTemp)).Value = True
    strTemp = String(10, " ")
    ReadIni "Options", "ScoreAction", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    optScoreOpt(Val(strTemp)).Value = True
    strTemp = String(10, " ")
    ReadIni "Options", "ExitPrompt", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    optExit(Val(strTemp)).Value = True
    strTemp = String(10, " ")
    ReadIni "Options", "LevelSave", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    If strTemp = 0 Then
        Check1.Value = "0"
    Else
        Check1.Value = "1"
    End If
    strTemp = String(10, " ")
    ReadIni "Options", "LoadAction", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    optStart(Val(strTemp)).Value = True
    strTemp = String(50, " ")
    ReadIni "Options", "LoadActionLevel", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, Len(RTrim(strTemp)) - 1)
    cmbSeries.Text = Split(strTemp, ",")(0)
    cmbLevel.ListIndex = Split(strTemp, ",")(1) - 1
End Sub

Private Sub optExit_Click(Index As Integer)
    cmdApply.Enabled = True
    intOptionExit = Index
End Sub

Private Sub optScore_Click(Index As Integer)
    cmdApply.Enabled = True
    If Index = 0 And optScore(Index).Value = True Then
        optScoreOpt(0).Enabled = True
        optScoreOpt(1).Enabled = True
        optScoreOpt(2).Enabled = True
    Else
        optScoreOpt(0).Enabled = False
        optScoreOpt(1).Enabled = False
        optScoreOpt(2).Enabled = False
    End If
    intOptionPrompt = Index
End Sub

Private Sub optScoreOpt_Click(Index As Integer)
    cmdApply.Enabled = True
    intOptionScore = Index
End Sub

Private Sub optStart_Click(Index As Integer)
    cmdApply.Enabled = True
    Select Case Index
        Case 0
            HideCombos
        Case 1
            HideCombos
        Case 2
            cmbSeries.Visible = True
            cmbLevel.Visible = True
    End Select
    intOptionLoad = Index
End Sub

Private Sub HideCombos()
    cmbSeries.Visible = False
    cmbLevel.Visible = False
End Sub

Private Sub TabStrip_Click()
    If TabStrip.SelectedItem = "Scoring" Then
        Frame1.Visible = False
        Frame2.Visible = True
        Picture2.Visible = False
    ElseIf TabStrip.SelectedItem = "Saving" Then
        Frame1.Visible = False
        Frame2.Visible = False
        Picture2.Visible = True
    Else
        Frame2.Visible = False
        Frame1.Visible = True

    End If
End Sub

Private Sub SaveSettings()
    Dim strTemp As String
    
    'load options
    strTemp = Val(intOptionLoad)
    WriteIni "Options", "LoadAction", strTemp, App.Path & "\Sokoban.ini"
    strTemp = cmbSeries.Text & "," & cmbLevel.Text
    WriteIni "Options", "LoadActionLevel", strTemp, App.Path & "\Sokoban.ini"
    
    'saving options
    strTemp = Val(intOptionExit)
    WriteIni "Options", "ExitPrompt", strTemp, App.Path & "\Sokoban.ini"
    If Check1.Value = "0" Then
        strTemp = "0"
    Else
        strTemp = "1"
    End If
    WriteIni "Options", "LevelSave", strTemp, App.Path & "\Sokoban.ini"
    
    'scoring options
    strTemp = Val(intOptionPrompt)
    WriteIni "Options", "ScorePrompt", strTemp, App.Path & "\Sokoban.ini"
    strTemp = Val(intOptionScore)
    WriteIni "Options", "ScoreAction", strTemp, App.Path & "\Sokoban.ini"
End Sub

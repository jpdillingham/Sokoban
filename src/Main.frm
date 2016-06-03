VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sokoban"
   ClientHeight    =   5595
   ClientLeft      =   4440
   ClientTop       =   3270
   ClientWidth     =   6660
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6660
   Begin VB.PictureBox Picture1 
      Height          =   5340
      Left            =   0
      ScaleHeight     =   5280
      ScaleWidth      =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   6660
      Begin VB.Timer tmrLevelPause 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   6150
         Top             =   2925
      End
      Begin VB.Timer tmrTask 
         Interval        =   1
         Left            =   6150
         Top             =   4275
      End
      Begin VB.Timer tmrScore 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6150
         Top             =   3825
      End
      Begin VB.FileListBox flbSeries 
         Height          =   285
         Left            =   4425
         TabIndex        =   5
         Top             =   4500
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Timer tmrLoad 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6150
         Top             =   3375
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   6075
         Top             =   4725
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Image imgDark 
         Height          =   330
         Left            =   3450
         Picture         =   "Main.frx":08CA
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgBorder 
         Height          =   330
         Left            =   3825
         Picture         =   "Main.frx":0EE4
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image cell 
         Height          =   330
         Index           =   1
         Left            =   0
         Picture         =   "Main.frx":14FE
         Top             =   0
         Width           =   330
      End
      Begin VB.Image imgPusher 
         Height          =   330
         Left            =   5325
         Picture         =   "Main.frx":1B78
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgPusherOverGrid 
         Height          =   330
         Left            =   5700
         Picture         =   "Main.frx":2192
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgLight 
         Height          =   330
         Left            =   4950
         Picture         =   "Main.frx":27AC
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgBox 
         Height          =   330
         Left            =   4575
         Picture         =   "Main.frx":2DC6
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image imgGrid 
         Height          =   330
         Left            =   4200
         Picture         =   "Main.frx":33E0
         Top             =   4875
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   2175
         X2              =   3900
         Y1              =   5325
         Y2              =   5325
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   3075
      X2              =   3075
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   4890
      X2              =   6660
      Y1              =   5355
      Y2              =   5355
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   4890
      X2              =   4890
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   6645
      X2              =   6645
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Label lblPushes 
      Alignment       =   2  'Center
      Caption         =   "0 Pushes (Best N\A)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4875
      TabIndex        =   3
      Top             =   5370
      Width           =   1755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   4890
      X2              =   6660
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   3075
      X2              =   4845
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   4830
      X2              =   4830
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   3075
      X2              =   4845
      Y1              =   5355
      Y2              =   5355
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   960
      X2              =   960
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   915
      Y1              =   5355
      Y2              =   5355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   960
      X2              =   3030
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   960
      X2              =   3030
      Y1              =   5355
      Y2              =   5355
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   3015
      X2              =   3015
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   5370
      Y2              =   5600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   900
      X2              =   900
      Y1              =   5585
      Y2              =   5355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   900
      Y1              =   5580
      Y2              =   5580
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   5370
      Width           =   960
   End
   Begin VB.Label lblMoves 
      Alignment       =   2  'Center
      Caption         =   "0 Moves (Best N\A)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3075
      TabIndex        =   2
      Top             =   5370
      Width           =   1755
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "00:00:00 (Best N\A)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   1
      Top             =   5370
      Width           =   2055
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load Game"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu hyphen1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Options"
         Shortcut        =   ^O
      End
      Begin VB.Menu hyphen9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSeries 
      Caption         =   "&Series"
      Begin VB.Menu mnuSeriesList 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu hyphen4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeriesOther 
         Caption         =   "&Other..."
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "&Level"
      Begin VB.Menu mnuLevelRestart 
         Caption         =   "&Restart"
         Shortcut        =   ^R
      End
      Begin VB.Menu hyphen5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLevelNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLevelPrevious 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu hyphen2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLevelSelect 
         Caption         =   "&Select..."
      End
   End
   Begin VB.Menu mnuMove 
      Caption         =   "&Move"
      Begin VB.Menu mnuMoveUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuMoveRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
   End
   Begin VB.Menu mnuApp 
      Caption         =   "&Appearance"
      Begin VB.Menu mnuAppGround 
         Caption         =   "&Ground"
      End
      Begin VB.Menu mnuAppBox 
         Caption         =   "&Box"
      End
      Begin VB.Menu mnuAppWall 
         Caption         =   "&Wall"
      End
      Begin VB.Menu mnuAppBack 
         Caption         =   "&Background..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpPlay 
         Caption         =   "&How To Play"
      End
      Begin VB.Menu hyphen3242 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cell_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim intL As Integer, intU As Integer, intR As Integer, intD As Integer, intPos As Integer, intI As Integer, intT As Integer
    
    intPos = intPusherPos - Index
    intL = intPusherPos Mod 20 - Index Mod 20
    intU = intPos / 20
    intR = (Index Mod 20) - (intPusherPos Mod 20)
    intD = Index / 20 - intPusherPos / 20
    'calculate total moves needed
    If intL < 0 Then
        intL = 0
    End If
    If intU < 0 Then
        intU = 0
    End If
    If intR < 0 Then
        intR = 0
    End If
    If intD < 0 Then
        intD = 0
    End If
    
    intT = intL + intU + intD + intR
    intI = intT
    Do While intI > 0
        If intL > 0 Then
            If Not cell(intPusherPos - 1).Picture = imgBorder.Picture Then
                MovePusher -1
                intL = intL - 1
            End If
        End If
        If intU > 0 Then
            If Not cell(intPusherPos - 20).Picture = imgBorder.Picture Then
                MovePusher -20
                intU = intU - 1
            End If
        End If
        If intD > 0 Then
            If Not cell(intPusherPos + 20).Picture = imgBorder.Picture Then
                MovePusher 20
                intD = intD - 1
            End If
        End If
        If intR > 0 Then
            If Not cell(intPusherPos + 1).Picture = imgBorder.Picture Then
                MovePusher 1
                intR = intR - 1
            End If
        End If
        
        intI = intI - 1
    Loop

End Sub

Private Sub cell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        DoUndo
    End If
End Sub

Private Sub cell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And Index = intPusherPos Then
        cell(intPusherPos).DragIcon = imgPusher.DragIcon
        cell(intPusherPos).Drag
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer, intI As Integer
    Dim X As Integer
    Dim strTemp As String, strTempS As String, strTempL As String, strTempN As String
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    X = 1
    i = 2
    Y = 2
    ' load the background
    Do Until i > 320
        Load cell(i)
        cell(i).Left = cell(1).Width * (Y - 1)
        cell(i).Top = cell(1).Height * (X - 1)
        cell(i).Visible = True
        i = i + 1
        Y = Y + 1
        ' if 20 cells have been loaded to the right, proceed to the next row
        If Y > 20 Then
            X = X + 1
            Y = 1
        End If
    Loop
    ' set the "best" variables, just for looks
    strBestMoves = "N\A"
    strBestPushes = "N\A"
    flbSeries.Path = App.Path & "\levels\"
    i = 0
    Do Until flbSeries.List(i) = ""
        If i = 0 Then
            mnuSeriesList(i).Caption = Split(flbSeries.List(i), ".")(0)
            mnuSeriesList(i).Checked = True
        Else
            Load mnuSeriesList(i)
            mnuSeriesList(i).Caption = Split(flbSeries.List(i), ".")(0)
            mnuSeriesList(i).Checked = False
        End If
        i = i + 1
    Loop
    ReDim strUndo(0)
    strTemp = String(10, " ")
    ReadIni "Options", "LoadAction", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    flbSeries.Path = App.Path & "\levels\"
    Select Case strTemp
        Case 0
            LoadSavedLevel "LastLevel"
        Case 1
            LoadLevelIni App.Path & "\levels\" & flbSeries.List(0), "1"
        Case 2
            strTemp = String(50, " ")
            ReadIni "Options", "LoadActionLevel", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
            strTempS = App.Path & "\levels\" & Split(strTemp, ",")(0) & ".sok"
            strTempL = Split(strTemp, ",")(1)
            strTempL = Left(strTempL, Len(RTrim(strTempL)))
            LoadLevelIni strTempS, strTempL
    End Select
    strTemp = String(30, " ")
    ReadIni "LastLevel", "1", vbNullString, strTemp, 30, App.Path & "\Sokoban.ini"
    If Len(RTrim(strTemp)) < 16 Then
        mnuFileLoad.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intResult As Integer, strTemp As String
    strTemp = String(10, " ")
    ReadIni "Options", "ExitPrompt", vbNullString, strTemp, 10, App.Path & "\Sokoban.ini"
    strTemp = Left(strTemp, 1)
    Select Case strTemp
        Case 0
            If blnSave = False Then
                intResult = MsgBox("Would you like to save your game?", vbApplicationModal + vbQuestion + vbYesNo + vbDefaultButton1, "Sokoban")
                If intResult = 6 Then
                    GameSave "LastLevel"
                End If
            End If
        Case 1
            If blnSave = False Then
                GameSave "LastLeveL", True
            End If
        Case 2
    End Select
End Sub

Private Sub mnuAppBack_Click()
    cd.DialogTitle = "Select picture"
    cd.ShowOpen
    'imgDark.Picture = LoadPicture(cd.FileName)
    Dim i As Integer
    i = 1
    Do While i < 321
        If cell(i).Picture = imgDark.Picture Then
            cell(i).Picture = LoadPicture(cd.FileName)
        End If
        i = i + 1
    Loop
    imgDark.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub mnuAppBox_Click()
    cd.DialogTitle = "Select picture"
    cd.ShowOpen
    'imgDark.Picture = LoadPicture(cd.FileName)
    Dim i As Integer
    i = 1
    Do While i < 321
        If cell(i).Picture = imgBox.Picture Then
            cell(i).Picture = LoadPicture(cd.FileName)
        End If
        i = i + 1
    Loop
    imgBox.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub mnuAppGround_Click()
    cd.DialogTitle = "Select picture"
    cd.ShowOpen
    'imgDark.Picture = LoadPicture(cd.FileName)
    Dim i As Integer
    i = 1
    Do While i < 321
        If cell(i).Picture = imgLight.Picture Then
            cell(i).Picture = LoadPicture(cd.FileName)
        End If
        i = i + 1
    Loop
    imgLight.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub mnuAppWall_Click()
    cd.DialogTitle = "Select picture"
    cd.ShowOpen
    'imgDark.Picture = LoadPicture(cd.FileName)
    Dim i As Integer
    i = 1
    Do While i < 321
        If cell(i).Picture = imgBorder.Picture Then
            cell(i).Picture = LoadPicture(cd.FileName)
        End If
        i = i + 1
    Loop
    imgBorder.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub mnuFileLoad_Click()
    LoadSavedLevel "LastLevel"
End Sub

Private Sub MovePusher(intD As Integer)
    Dim intDD As Integer, intDDD As Integer, intBoxMoved As Integer, intBoxLast As Integer, strTemp As String
    intDD = intD * 2
    intDDD = intD * 3
    
    If cell(intPusherPos).Tag = "grid" Then
        cell(intPusherPos).Picture = imgGrid.Picture
    Else
        cell(intPusherPos).Picture = imgLight.Picture
    End If
    
    If cell(intPusherPos + intD) = imgBorder.Picture Then
        If cell(intPusherPos).Tag = "grid" Then
            cell(intPusherPos).Picture = imgPusherOverGrid.Picture
        Else
            cell(intPusherPos).Picture = imgPusher.Picture
        End If
        Exit Sub
    End If
    ' if it is being pushed into a box
    If cell(intPusherPos + intD) = imgBox.Picture Then
        ' if the box is up against a wall or another box, stop
        If cell(intPusherPos + intDD).Picture = imgBorder.Picture Or cell(intPusherPos + intDD).Picture = imgBox.Picture Then
            If cell(intPusherPos).Tag = "grid" Then
                cell(intPusherPos).Picture = imgPusherOverGrid.Picture
            Else
                cell(intPusherPos).Picture = imgPusher.Picture
            End If
            Exit Sub
        ' move the box
        Else
            If cell(intPusherPos + intDD).Tag = "grid" And Not cell(intPusherPos + intD).Tag = "grid" Then
                intGrid = intGrid - 1
            ElseIf Not cell(intPusherPos + intDD).Tag = "grid" And cell(intPusherPos + intD).Tag = "grid" Then
                intGrid = intGrid + 1
            End If
            cell(intPusherPos + intDD).Picture = imgBox.Picture
            intPushes = intPushes + 1
            lblPushes = intPushes & " Pushes (Best " & strBestPushes & ")"
            intBoxMoved = intPusherPos + intDD
            intBoxLast = intPusherPos + intD
        End If

    End If
    ' move the pusher
        If cell(intPusherPos + intD).Tag = "grid" Then
            cell(intPusherPos + intD).Picture = imgPusherOverGrid.Picture
        Else
            cell(intPusherPos + intD).Picture = imgPusher.Picture
        End If
    If intBoxMoved = 0 Then
        AddUndo intPusherPos + intD & " " & intPusherPos & " p"
    Else
        AddUndo intBoxMoved & " " & intBoxLast & " b " & intPusherPos + intD & " " & intPusherPos
    End If
    intPusherPos = intPusherPos + intD

    intMoves = intMoves + 1
    lblMoves = intMoves & " Moves (Best " & strBestMoves & ")"
    If intGrid <= 0 Then
        strTemp = String(4, " ")
        ReadIni "Options", "ScorePrompt", vbNullString, strTemp, 4, App.Path & "\Sokoban.ini"
        If Left(strTemp, 1) = "1" Then
            tmrLoad.Enabled = True
        Else
            strTemp = String(2, " ")
            ReadIni "Options", "ScoreAction", vbNullString, strTemp, 2, App.Path & "\Sokoban.ini"
            Select Case Left(strTemp, 1)
                Case 0
                    If lngScoreTime < lngBestTime Then
                        SaveScore Val(intLevel)
                    End If
                Case 1
                    SaveScore Val(intLevel)
                Case Else
            End Select
            strTemp = String(2, " ")
            ReadIni "Options", "LevelSave", vbNullString, strTemp, 2, App.Path & "\Sokoban.ini"
            strTemp = Left(strTemp, 1)
            If strTemp = "1" Then
                blnSaveAfterLoad = True
            End If
            tmrLevelPause.Enabled = True
        End If
    End If
End Sub

Private Sub mnuFileUndo_Click()
    DoUndo
End Sub

Private Sub mnuFileOptions_Click()
    Load Options
    Options.Show
End Sub

Private Sub mnuFileSave_Click()
    GameSave "LastLevel"
    blnSave = True
End Sub

Public Sub GameSave(strSaveName As String, Optional Quiet As Boolean)
    Dim strTemp As String, i As Integer, X As Integer, Y As Integer, strT As String
    i = 1
    X = 1
    Y = 1
    Do Until i > 320
        If cell(i).Tag = "grid" Then
            If cell(i).Picture = imgBox.Picture Then
                strTemp = strTemp & "7"
            ElseIf cell(i).Picture = imgPusherOverGrid.Picture Then
                strTemp = strTemp & "8"
            Else
                strTemp = strTemp & "4"
            End If
        Else
            Select Case cell(i).Picture
            Case imgLight.Picture
                strTemp = strTemp & "1"
            Case imgDark.Picture
                strTemp = strTemp & "2"
            Case imgPusher.Picture
                strTemp = strTemp & "3"
            Case imgBox.Picture
                strTemp = strTemp & "5"
            Case imgBorder.Picture
                strTemp = strTemp & "6"
            Case Else
                strTemp = strTemp & "2"
            End Select
        End If
        If X = 20 Then
            strT = Val(Y)
            WriteIni strSaveName, strT, strTemp, App.Path & "\Sokoban.ini"
            Y = Y + 1
            X = 0
            strTemp = ""
        End If
        i = i + 1
        X = X + 1
    Loop
    strTemp = Val(intNumLevels)
    WriteIni strSaveName, "Levels", strTemp, App.Path & "\Sokoban.ini"
    WriteIni strSaveName, "Series", strSeries, App.Path & "\Sokoban.ini"
    strTemp = Val(intLevel)
    WriteIni strSaveName, "Level", strTemp, App.Path & "\Sokoban.ini"
    strTemp = Val(intMoves)
    WriteIni strSaveName, "Moves", strTemp, App.Path & "\Sokoban.ini"
    strTemp = Val(intPushes)
    WriteIni strSaveName, "Pushes", strTemp, App.Path & "\Sokoban.ini"
    strTemp = Val(lngScoreTime)
    WriteIni strSaveName, "Time", strTemp, App.Path & "\Sokoban.ini"
    mnuFileLoad.Enabled = True
    If Quiet = False Then
        MsgBox "Your game has been saved." & vbCr & "It will be loaded next time you start Sokoban.", vbInformation + vbOKOnly + vbDefaultButton1
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    Load About
    About.Show
End Sub

Private Sub mnuHelpContents_Click()
    MsgBox "Im too lazy to make a .hlp, sorry."
End Sub

Private Sub mnuHelpPlay_Click()
    MsgBox "Im too lazy to make a .hlp, sorry."
End Sub

Private Sub mnuLevelNext_Click()
    'go to the next level
    LoadLevelIni strLevelFileName, intLevel + 1
End Sub

Private Sub mnuLevelPrevious_Click()
    ' go to the last level
    LoadLevelIni strLevelFileName, intLevel - 1
End Sub

Private Sub mnuLevelRestart_Click()
    LoadLevelIni strLevelFileName, intLevel + 0
End Sub

Private Sub mnuLevelSelect_Click()
    ' open the level select dialog
    Load LevelSelect
    LevelSelect.Show
End Sub

Private Sub mnuMoveRedo_Click()
    Redo
End Sub

Private Sub mnuMoveUndo_Click()
    DoUndo
End Sub

Private Sub mnuSeriesList_Click(Index As Integer)
    Dim i As Integer
    ' load the series
    LoadLevelIni App.Path & "\levels\" & mnuSeriesList(Index).Caption & ".sok", 1
    i = 0
    ' uncheck all the other series
    Do Until i > mnuSeriesList.UBound
        mnuSeriesList(i).Checked = False
        i = i + 1
    Loop
    ' check the series selected
    mnuSeriesList(Index).Checked = True
End Sub

Private Sub mnuSeriesOther_Click()
    On Error GoTo CancelSelect
    Dim i As Integer
    cd.InitDir = App.Path
    cd.Filter = "Sokoban Levels (*.sok)|*.sok"
    cd.CancelError = True
    cd.ShowOpen
    ' load the level selected
    LoadLevelIni cd.FileName, 1
    ' copy the selected level to the levels directory to add it to the list
    FileCopy cd.FileName, App.Path & "\levels\" & strSeries & ".sok"
    ' unload the series list
    i = 1
    Do Until i > mnuSeriesList.UBound
        Unload mnuSeriesList(i)
        i = i + 1
    Loop
    ' refresh the file list box so it shows the copied level
    flbSeries.Refresh
    ' add all of the series back to the list
    i = 0
    Do Until flbSeries.List(i) = ""
        ' index 0 doesnt need to be loaded
        If Not i = 0 Then
            Load mnuSeriesList(i)
        End If
        ' set the caption and checked properties
        mnuSeriesList(i).Caption = Split(flbSeries.List(i), ".")(0)
        mnuSeriesList(i).Checked = False
        ' if the series just added is the series that is being displayed, check it
        If Split(flbSeries.List(i), ".")(0) = strSeries Then
            mnuSeriesList(i).Checked = True
        End If
        i = i + 1
    Loop

CancelSelect:

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        ' backspace key
        Case 8
            DoUndo
        ' page up
        Case 13
            Redo
        Case 33
            LoadLevelIni strLevelFileName, intLevel + 1
        ' page down
        Case 34
            LoadLevelIni strLevelFileName, intLevel - 1
        ' left arrow
        Case 37
            Call MovePusher(-1)
        ' up arrow
        Case 38
            Call MovePusher(-20)
        ' right arrow
        Case 39
            Call MovePusher(1)
        ' down arrow
        Case 40
            Call MovePusher(20)
    End Select
    blnSave = False
End Sub

Private Sub tmrLevelPause_Timer()
    LoadLevelIni strLevelFileName, intLevel + 1
    tmrLevelPause.Enabled = False
End Sub

Private Sub tmrLoad_Timer()
    Load Score
    Score.Show
    RedrawCells
    tmrLoad.Enabled = False
End Sub

Private Sub tmrTask_Timer()
    ' this sorta corrects an annoying bug that hides the form when minimized
    ' by clicking on the taskbar button twice.  it may only be a problem with litestep.
    
    ' if the form isnt visible, set the state to minimized and show it again
    If Me.Visible = False Then
        Me.Hide
        Me.WindowState = vbMinimized
        Me.Show
    End If
    lblLevel = "Level " & intLevel
End Sub

Private Sub tmrScore_Timer()
    ' calculate the time in hhmmss
    hhmmss "regular", lngScoreTime
    ' increase the time
    lngScoreTime = lngScoreTime + 1
    ' display it
    lblScore = strFormattedScoreTime & " (Best " & strBestTime & ")"
End Sub


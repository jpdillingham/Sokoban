Attribute VB_Name = "Procedures"
Public Declare Function WriteIni Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function ReadIni Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public intNumLevels As Integer
Public strSeries As String
Public intLevel As Integer
Public strLevelFileName As String
Public intPusherPos As Integer
Public intGrid As Integer
Public lngScoreTime As Long
Public strFormattedScoreTime As String
Public intPushes As Integer
Public intMoves As Integer
Public strFileName As String
Public intBestMoves As Integer
Public strBestMoves As String
Public intBestPushes As Integer
Public strBestPushes As String
Public lngBestTime As Long
Public strBestTime As String
Public strUndo() As String
Public strResult As String
Public strDrag As String
Public intLastDrag As Integer
Public intUndoPos As Integer
Public blnSave As Boolean
Public blnSaveAfterLoad As Boolean


Public Sub LoadSavedLevel(strSavedLevel As String)
    ' loads the last configuration
    ' i could have modified the loadlevelini procedure, but this was easier.
    
    Dim strTemp As String, intI As Integer, intL As Integer, strT As String
    Dim intImage As Integer, intCell As Integer, strBleh As String
    'On Error GoTo LoadErr
    Open App.Path & "\Sokoban.ini" For Input As #1
    Close #1
    
    intGrid = 0
    intI = 1
    Do While intI < 17
        strTemp = String(30, " ")
        strBleh = Val(intI)
        ReadIni strSavedLevel, strBleh, vbNullString, strTemp, 30, App.Path & "\Sokoban.ini"
        If Len(RTrim(strTemp)) < 16 Then GoTo LoadErr
        intL = 1
        Do While intL < 21
            intImage = Val(Mid(strTemp, intL, 1))
            intCell = intCell + 1
            Main.cell(intCell).Tag = ""
            Select Case intImage
                Case 1
                    Main.cell(intCell).Picture = Main.imgLight.Picture
                Case 2
                    Main.cell(intCell).Picture = Main.imgDark.Picture
                Case 3
                    Main.cell(intCell).Picture = Main.imgPusher.Picture
                    intPusherPos = intCell
                Case 4
                    Main.cell(intCell).Picture = Main.imgGrid.Picture
                    Main.cell(intCell).Tag = "grid"
                    intGrid = intGrid + 1
                Case 5
                    Main.cell(intCell).Picture = Main.imgBox.Picture
                Case 6
                    Main.cell(intCell).Picture = Main.imgBorder.Picture
                ' boxes on top of grid cells.
                Case 7
                    Main.cell(intCell).Tag = "grid"
                    Main.cell(intCell).Picture = Main.imgBox.Picture
                ' when the pusher is in a grid cell.
                Case 8
                    Main.cell(intCell).Tag = "grid"
                    intGrid = intGrid + 1
                    Main.cell(intCell).Picture = Main.imgPusherOverGrid.Picture
                    intPusherPos = intCell
            End Select
            intL = intL + 1
        Loop
        intI = intI + 1
    Loop
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Level", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intLevel = Val(strTemp)
    Main.lblLevel = "Level " & intLevel
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Series", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    strSeries = Left(Trim(strTemp), Len(Trim(strTemp)) - 1)
    strLevelFileName = App.Path & "\levels\" & strSeries & ".sok"
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Levels", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intNumLevels = Val(strTemp)
    Main.Caption = "Sokoban - " & strSeries & " - Level " & intLevel & " of " & intNumLevels
    
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Time", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    lngScoreTime = Val(strTemp)
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "BestTime", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    lngBestTime = Val(strTemp)
    hhmmss "best", lngBestTime
    hhmmss "blah", lngScoreTime
    Main.lblScore = strFormattedScoreTime & " (Best " & strBestTime & ")"
    
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Moves", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intMoves = Val(strTemp)
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "BestMoves", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intBestMoves = Val(strTemp)
    If intBestMoves = 0 Then
        Main.lblMoves = intMoves & " Moves (Best N\A)"
    Else
        Main.lblMoves = intMoves & " Moves (Best " & intBestMoves & ")"
    End If
    
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "Pushes", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intPushes = Val(strTemp)
    strTemp = String(50, " ")
    ReadIni strSavedLevel, "BestPushes", vbNullString, strTemp, 50, App.Path & "\Sokoban.ini"
    intBestPushes = Val(strTemp)
    If intBestPushes = 0 Then
        Main.lblPushes = intPushes & " Moves (Best N\A)"
    Else
        Main.lblPushes = intPushes & " Moves (Best " & intBestMoves & ")"
    End If
    Main.tmrScore.Enabled = True
    intI = 0
    Do Until intI > Main.flbSeries.ListCount - 1
        If Main.mnuSeriesList(intI).Caption = strSeries Then
            Main.mnuSeriesList(intI).Checked = True
        Else
            Main.mnuSeriesList(intI).Checked = False
        End If
        intI = intI + 1
    Loop
    ReDim strUndo(0)
    Main.mnuMoveRedo.Enabled = False
    Main.mnuMoveUndo.Enabled = False
    If intLevel = intNumLevels Then
        Main.mnuLevelNext.Enabled = False
        Main.mnuLevelPrevious.Enabled = True
    ' if its level 1, disable the previous menu choice
    ElseIf intLevel = 1 Then
        Main.mnuLevelPrevious.Enabled = False
    ' enable them both if not
    Else
        Main.mnuLevelNext.Enabled = True
        Main.mnuLevelPrevious.Enabled = True
    End If
    GetScore strSeries, Val(intLevel)
    Main.lblScore = strFormattedScoreTime & " (Best " & strBestTime & ")"
    blnSave = True
    GoTo EndOfSub

LoadErr:
    Main.flbSeries.Path = App.Path & "\levels\"
    If Main.flbSeries.List(0) = "" Then
        MsgBox "The level files are missing.  If the levels are stored in a folder other than " \ levels \ ", please relocate them to the levels directory.  If not, please reinstall Sokoban.", vbCritical + vbOKOnly, "Sokoban"
        End
    End If
    LoadLevelIni App.Path & "\levels\" & Main.flbSeries.List(0), 1
    
EndOfSub:
End Sub

Public Sub LoadLevelIni(strFileName As String, strLevel As String)
    Dim strTemp As String, intI As Integer, intL As Integer
    Dim intImage As Integer, intCell As Integer, strBleh As String
    strLevelFileName = strFileName
    'On Error GoTo LoadErr
    ' get the number of levels contained in the file
    strTemp = String(30, " ")
    ReadIni "info", "levels", vbNullString, strTemp, 30, strFileName
    intNumLevels = Val(strTemp)

    ' if the level number provided is higher than the total number of levels, send a msg
    If Val(strLevel) > intNumLevels Then
        MsgBox "You have reached the last level.", vbInformation + vbOKOnly, "Sokoban"
        Exit Sub
    ' if the level is the last one, disable the next menu choice
    ElseIf Val(strLevel) = intNumLevels Then
        Main.mnuLevelNext.Enabled = False
        Main.mnuLevelPrevious.Enabled = True
    ElseIf Val(strLevel) < 1 Then
        Beep
        Exit Sub
    ' if its level 1, disable the previous menu choice
    ElseIf Val(strLevel) = 1 Then
        Main.mnuLevelPrevious.Enabled = False
    ' enable them both if not
    Else
        Main.mnuLevelNext.Enabled = True
        Main.mnuLevelPrevious.Enabled = True
    End If
    
    RedrawCells
    intGrid = 0
    intI = 1
    Do While intI < 17
        strTemp = String(30, " ")
        strBleh = Val(intI)
        ReadIni strLevel, strBleh, vbNullString, strTemp, 30, strFileName
        intL = 1
        Do While intL < 21
            intImage = Val(Mid(strTemp, intL, 1))
            intCell = intCell + 1
            Main.cell(intCell).Tag = ""
            Select Case intImage
                Case 1
                    Main.cell(intCell).Picture = Main.imgLight.Picture
                Case 2
                    Main.cell(intCell).Picture = Main.imgDark.Picture
                Case 3
                    Main.cell(intCell).Picture = Main.imgPusher.Picture
                    intPusherPos = intCell
                Case 4
                    Main.cell(intCell).Picture = Main.imgGrid.Picture
                    Main.cell(intCell).Tag = "grid"
                    intGrid = intGrid + 1
                Case 5
                    Main.cell(intCell).Picture = Main.imgBox.Picture
                Case 6
                    Main.cell(intCell).Picture = Main.imgBorder.Picture
                ' boxes on top of grid cells.
                Case 7
                    Main.cell(intCell).Tag = "grid"
                    Main.cell(intCell).Picture = Main.imgBox.Picture
                ' when the pusher is in a grid cell.
                Case 8
                    Main.cell(intCell).Tag = "grid"
                    intGrid = intGrid + 1
                    Main.cell(intCell).Picture = Main.imgPusherOverGrid.Picture
                    intPusherPos = intCell
            End Select
            intL = intL + 1
        Loop
        intI = intI + 1
    Loop
    intLevel = Val(strLevel)
    SetSeriesAndLevel strFileName
    strFormattedScoreTime = "00:00:00"
    lngScoreTime = 0
    intMoves = 0
    intPushes = 0
    Main.tmrScore.Enabled = True
    ReDim strUndo(0)
    Main.mnuMoveRedo.Enabled = False
    Main.mnuMoveUndo.Enabled = False
    If blnSaveAfterLoad = True Then
        Main.GameSave "LastLevel", True
        blnSaveAfterLoad = False
    End If
    intI = 0
    Do Until intI > Main.flbSeries.ListCount - 1
        If Main.mnuSeriesList(intI).Caption = strSeries Then
            Main.mnuSeriesList(intI).Checked = True
        Else
            Main.mnuSeriesList(intI).Checked = False
        End If
        intI = intI + 1
    Loop
    GoTo EndOfSub

LoadErr:
    MsgBox "Error opening '" & strFileName & "'." & vbLf & " File is either missing or is not a valid sokoban level.", vbOKOnly + vbInformation + vbApplicationModalmodal, "Sokoban"

EndOfSub:
    blnSave = True
End Sub

Public Sub SaveScore(strLevel As String)
    Dim strTemp As String
    strTemp = Replace(strSeries, " ", "_")
    WriteIni strTemp, strLevel, intMoves & "," & intPushes & "," & lngScoreTime, App.Path & "\Sokoban.ini"
End Sub

Public Sub GetScore(strSeriesName As String, strLevel As String)
    Dim strTemp As String
    On Error GoTo NoScore
    strTemp = String(70, " ")
    strSeriesName = Replace(strSeriesName, " ", "_")
    ReadIni strSeriesName, strLevel, vbNullString, strTemp, 70, App.Path & "\Sokoban.ini"
    If strTemp = vbLf Then GoTo NoScore
    intBestMoves = Val(Split(strTemp, ",")(0))
    strBestMoves = intBestMoves
    intBestPushes = Val(Split(strTemp, ",")(1))
    strBestPushes = intBestPushes
    lngBestTime = Val(Split(strTemp, ",")(2))
    Main.lblMoves = "0 Moves (Best " & intBestMoves & ")"
    Main.lblPushes = "0 Pushes (Best " & intBestPushes & ")"
    hhmmss "best", lngBestTime
    Main.lblScore = "00:00:00 (Best " & strBestTime & ")"
    GoTo NormalEnd
    
NoScore:
    intBestPushes = 0
    intBestMoves = 0
    lngBestTime = 99999999
    strBestPushes = "N\A"
    strBestMoves = "N\A"
    strBestTime = "N\A"
    Main.lblMoves = "0 Moves (Best " & strBestMoves & ")"
    Main.lblPushes = "0 Pushes (Best " & strBestPushes & ")"
    Main.lblScore = "00:00:00 (Best " & strBestTime & ")"

NormalEnd:
End Sub

Public Sub AddUndo(strNewUndo As String)
    If intUndoPos > 0 Then
        ReDim Preserve strUndo(UBound(strUndo) - intUndoPos)
    End If
    If intUndoPos = UBound(strUndo) Then
        ReDim strUndo(0)
    End If
    Dim intI As Integer
    ReDim Preserve strUndo(UBound(strUndo) + 1)
    strUndo(UBound(strUndo)) = strNewUndo
    intUndoPos = 0
    Main.mnuMoveUndo.Enabled = True
    blnSave = False
End Sub

Public Sub Redo()
    Dim intN As Integer, intO As Integer, intPN As Integer, intPO As Integer
    ' if there is nothing to undo
    intUndoPos = intUndoPos - 1
    If UBound(strUndo) <= 0 Or intUndoPos = UBound(strUndo) Or intUndoPos < 0 Then
        Beep
        intUndoPos = 0
        Main.mnuMoveRedo.Enabled = False
        Exit Sub
    End If
    
    ' if a box was moved
    If Split(strUndo(UBound(strUndo) - intUndoPos), " ")(2) = "b" Then
        ' get the coords of the box and pusher
        intN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(1)
        intO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(0)
        intPN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(4)
        intPO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(3)
        ' if the pusher is over a grid square
        If Main.cell(intPO).Tag = "grid" Then
            Main.cell(intPO).Picture = Main.imgPusherOverGrid.Picture
        Else
            Main.cell(intPO).Picture = Main.imgPusher.Picture
        End If
        If Main.cell(intPN).Tag = "grid" Then
            Main.cell(intPN).Picture = Main.imgGrid.Picture
        Else
            Main.cell(intPN).Picture = Main.imgLight.Picture
        End If
        
        intPusherPos = intPO
        ' if something was over a grid square, replace the image
        If Main.cell(intN).Tag = "grid" Then
            If intN = intPusherPos Then
                Main.cell(intN).Picture = Main.imgPusherOverGrid.Picture
            Else
                Main.cell(intN).Picture = Main.imgGrid.Picture
            End If
            intGrid = intGrid + 1
        Else
            ' over plain space
            If intN = intPusherPos Then
                Main.cell(intN).Picture = Main.imgPusher.Picture
            Else
                Main.cell(intN).Picture = Main.imgLight.Picture
            End If
        End If
        ' replace the box, if it is going onto a grid, decrease the grid count
        Main.cell(intO).Picture = Main.imgBox.Picture
        If Main.cell(intO).Tag = "grid" Then
            intGrid = intGrid - 1
        End If
    Else
        ' just the pusher was moved
        intN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(1)
        intO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(0)
        ' if it was over a grid
        If Main.cell(intN).Tag = "grid" Then
            Main.cell(intN).Picture = Main.imgGrid.Picture
        Else
            Main.cell(intN).Picture = Main.imgLight.Picture
        End If
        ' if it is going over a grid
        If Main.cell(intO).Tag = "grid" Then
            Main.cell(intO).Picture = Main.imgPusherOverGrid.Picture
        Else
            Main.cell(intO).Picture = Main.imgPusher.Picture
        End If
        ' reset the pusher position
        intPusherPos = intO
    End If
    blnSave = False
End Sub

Public Sub DoUndo()
    Dim intN As Integer, intO As Integer, intPN As Integer, intPO As Integer
    ' if there is nothing to undo
    If UBound(strUndo) <= 0 Or intUndoPos >= UBound(strUndo) Or intUndoPos < 0 Then
        Beep
        Main.mnuMoveUndo.Enabled = False
        Exit Sub
    End If
    
    ' if a box was moved
    If Split(strUndo(UBound(strUndo) - intUndoPos), " ")(2) = "b" Then
        ' get the coords of the box and pusher
        intN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(0)
        intO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(1)
        intPN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(3)
        intPO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(4)
        ' if the pusher is over a grid square
        If Main.cell(intPO).Tag = "grid" Then
            Main.cell(intPO).Picture = Main.imgPusherOverGrid.Picture
        Else
            Main.cell(intPO).Picture = Main.imgPusher.Picture
        End If
        intPusherPos = intPO
        ' if something was over a grid square, replace the image
        If Main.cell(intN).Tag = "grid" Then
            Main.cell(intN).Picture = Main.imgGrid.Picture
            intGrid = intGrid + 1
        Else
            ' over plain space
            Main.cell(intN).Picture = Main.imgLight.Picture
        End If
        ' replace the box, if it is going onto a grid, decrease the grid count
        Main.cell(intO).Picture = Main.imgBox.Picture
        If Main.cell(intO).Tag = "grid" Then
            intGrid = intGrid - 1
        End If
    Else
        ' just the pusher was moved
        intN = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(0)
        intO = Split(strUndo(UBound(strUndo) - intUndoPos), " ")(1)
        ' if it was over a grid
        If Main.cell(intN).Tag = "grid" Then
            Main.cell(intN).Picture = Main.imgGrid.Picture
        Else
            Main.cell(intN).Picture = Main.imgLight.Picture
        End If
        ' if it is going over a grid
        If Main.cell(intO).Tag = "grid" Then
            Main.cell(intO).Picture = Main.imgPusherOverGrid.Picture
        Else
            Main.cell(intO).Picture = Main.imgPusher.Picture
        End If
        ' reset the pusher position
        intPusherPos = intO
    End If
    Main.mnuMoveRedo.Enabled = True
    intUndoPos = intUndoPos + 1
    blnSave = False
End Sub

Public Sub hhmmss(strWhat As String, lngTime As Long)
    ' since VB documentation is so poor, i couldnt find a way to turn an integer into a HH:MM:SS string
    ' so i made my own function to do it.
    Dim intH As Integer, intM As Integer, intS As Integer
    Dim strH As String, strM As String, strS As String
    intH = 0
    intS = 0
    intM = 0
    ' if its zero, return n\a or all zeros
    If lngTime < 1 Then
        If strWhat = "best" Then
            strBestTime = "N\A"
        Else
            strFormattedScoreTime = "00:00:00"
        End If
    Else
        ' calculate the time
        intS = lngTime Mod 60
        intH = Int(lngTime / 60)
        intM = intH Mod 60
        intH = Int(intH / 60)
        ' if the number is only one digit, add another zero
        If intH < 10 Then
            strH = "0" & intH
        Else
            strH = intH
        End If
        If intM < 10 Then
            strM = "0" & intM
        Else
            strM = Val(intM)
        End If
        If intS < 10 Then
            strS = "0" & intS
        Else
            strS = intS
        End If
        ' if the function is being used to calculate "best" time
        If strWhat = "best" Then
            strBestTime = strH & ":" & strM & ":" & strS
        ' if it is being used for the current time
        Else
            strFormattedScoreTime = strH & ":" & strM & ":" & strS
        End If
    End If
End Sub

Public Sub RedrawCells()
    Dim i As Integer
    i = 1
    ' replace all cells with the default image
    Do Until i > 320
        Main.cell(i).Picture = Main.imgDark.Picture
        i = i + 1
    Loop
End Sub

Public Sub SetSeriesAndLevel(strFileName As String)
    Dim strName As String, intI As Integer, strLevel As String
    ' remove the extension
    strName = Split(strFileName, ".")(0)
    ' remove the path
    intI = Len(strName)
    Do Until Mid(strName, intI, 1) = "\"
        intI = intI - 1
    Loop
    ' set the series name
    strName = Right(strName, Len(strName) - intI)
    ' set the form caption
    Main.Caption = "Sokoban - " & strName & " - Level " & intLevel & " of " & intNumLevels
    ' load scores
    strLevel = Val(intLevel)
    strSeries = strName
    GetScore strName, strLevel
End Sub

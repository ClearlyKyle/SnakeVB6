Attribute VB_Name = "Module1"
Public ScoreRecord(2, 4) As SCOREDATA
Public PlayerScore As Integer
Public PlayerName As String

Public TimeAttackMode As Boolean, SurvivalMode As Boolean

Public Type SCOREDATA
    ScoreName As String
    ScoreMark As Integer
End Type

Sub LoadHighScore()
    Dim tempName As String, fileNames As String, fileScore As String
    Dim tempMark As Integer
    
    fileNames = App.Path + "\HighscoreNames.dat"
    fileScore = App.Path + "\HighscoreScore.dat"
    
    Open fileNames For Input As #1
    Open fileScore For Input As #2
        For x = 1 To 2
            For i = 0 To 4
                    Input #1, tempName
                    Input #2, tempMark
                ScoreRecord(x, i).ScoreName = tempName
                ScoreRecord(x, i).ScoreMark = tempMark
            Next i
        Next x
    Close #1
    Close #2
End Sub

Private Sub SaveHighScore()
    Dim tempName As String, fileNames As String, fileScore As String
    Dim tempMark As Integer
    
    fileNames = App.Path + "\HighscoreNames.dat"
    fileScore = App.Path + "\HighscoreScore.dat"
    
    Open fileNames For Output As #1
    Open fileScore For Output As #2
        For x = 1 To 2
            For i = 0 To 4
                tempName = ScoreRecord(x, i).ScoreName
                tempMark = ScoreRecord(x, i).ScoreMark
                Print #1, tempName
                Print #2, tempMark
            Next i
        Next x
    Close #1
    Close #2
End Sub

Sub ProcessScore(ByVal PlayerScore As String, PlayerName As String, TimeAttackMode As Boolean, SurvivalMode As Boolean)
    Dim x As Integer
    
    If TimeAttackMode = True Then
        x = 1
    ElseIf SurvivalMode = True Then
        x = 2
    End If
    
    Call LoadHighScore
    
    If PlayerScore < ScoreRecord(x, 0).ScoreMark Then
        MsgBox "Oops sorry.. Game Over !" & vbCrLf & "Your score is :" & PlayerScore, vbInformation, ""
        Exit Sub
    End If
    
    If PlayerName = "" Then
        PlayerName = "Unknown"
    End If
    
    Select Case PlayerScore
        Case Is >= ScoreRecord(x, 4).ScoreMark
            ScoreRecord(x, 0).ScoreName = ScoreRecord(x, 1).ScoreName
            ScoreRecord(x, 0).ScoreMark = ScoreRecord(x, 1).ScoreMark
            ScoreRecord(x, 1).ScoreName = ScoreRecord(x, 2).ScoreName
            ScoreRecord(x, 1).ScoreMark = ScoreRecord(x, 2).ScoreMark
            ScoreRecord(x, 2).ScoreName = ScoreRecord(x, 3).ScoreName
            ScoreRecord(x, 2).ScoreMark = ScoreRecord(x, 3).ScoreMark
            ScoreRecord(x, 3).ScoreName = ScoreRecord(x, 4).ScoreName
            ScoreRecord(x, 3).ScoreMark = ScoreRecord(x, 4).ScoreMark
            ScoreRecord(x, 4).ScoreName = PlayerName
            ScoreRecord(x, 4).ScoreMark = PlayerScore
            
            SaveHighScore
            Exit Sub
            
        Case Is >= ScoreRecord(x, 3).ScoreMark
            ScoreRecord(x, 0).ScoreName = ScoreRecord(x, 1).ScoreName
            ScoreRecord(x, 0).ScoreMark = ScoreRecord(x, 1).ScoreMark
            ScoreRecord(x, 1).ScoreName = ScoreRecord(x, 2).ScoreName
            ScoreRecord(x, 1).ScoreMark = ScoreRecord(x, 2).ScoreMark
            ScoreRecord(x, 2).ScoreName = ScoreRecord(x, 3).ScoreName
            ScoreRecord(x, 2).ScoreMark = ScoreRecord(x, 3).ScoreMark
            ScoreRecord(x, 3).ScoreName = PlayerName
            ScoreRecord(x, 3).ScoreMark = PlayerScore
            
            SaveHighScore
            Exit Sub

        Case Is >= ScoreRecord(x, 2).ScoreMark
            ScoreRecord(x, 0).ScoreName = ScoreRecord(x, 1).ScoreName
            ScoreRecord(x, 0).ScoreMark = ScoreRecord(x, 1).ScoreMark
            ScoreRecord(x, 1).ScoreName = ScoreRecord(x, 2).ScoreName
            ScoreRecord(x, 1).ScoreMark = ScoreRecord(x, 2).ScoreMark
            ScoreRecord(x, 2).ScoreName = PlayerName
            ScoreRecord(x, 2).ScoreMark = PlayerScore
            
            SaveHighScore
            Exit Sub
            
        Case Is >= ScoreRecord(x, 1).ScoreMark
            ScoreRecord(x, 0).ScoreName = ScoreRecord(x, 1).ScoreName
            ScoreRecord(x, 0).ScoreMark = ScoreRecord(x, 1).ScoreMark
            ScoreRecord(x, 1).ScoreName = PlayerName
            ScoreRecord(x, 1).ScoreMark = PlayerScore
            
            SaveHighScore
            Exit Sub
            
        Case Is >= ScoreRecord(x, 0).ScoreMark
            ScoreRecord(x, 0).ScoreName = PlayerName
            ScoreRecord(x, 0).ScoreMark = PlayerScore
            
            SaveHighScore
            Exit Sub
    
    End Select
    
End Sub


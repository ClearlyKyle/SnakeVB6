VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Snake"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox HelpScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5265
      ScaleWidth      =   5265
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox txtScoreName 
      Appearance      =   0  'Flat
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer TimeLeft 
      Interval        =   1000
      Left            =   5520
      Top             =   2520
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   5520
      Top             =   2040
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Strobe"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox ScoreBlock 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   720
      Width           =   255
   End
   Begin VB.Timer StartTimer 
      Interval        =   1000
      Left            =   5520
      Top             =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label lblHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label lblGameOver 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   600
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblEnterName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblScores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Highscores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblFinalScore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label LabelName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label StartTA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Attack"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label StartSurvival 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Survival"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label TitleSnake 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Snake"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   4215
   End
   Begin VB.Shape Snake 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   1200
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Snake 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   1440
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Snake 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   1680
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Snake 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   1920
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Snake 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   2160
      Top             =   720
      Width           =   255
   End
   Begin VB.Label TimeControl 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LabelCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2040
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempT As Integer, TempL As Integer
Dim currtop As Integer, currleft As Integer

Dim PauseGame As Boolean
Dim lastMove As String

Private mainGameLoop As Boolean
Private facingUp As Boolean, facingDown As Boolean, facingLeft As Boolean, facingRight As Boolean
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const vbGameSpeed As Long = 80

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft
            If facingRight Then Exit Sub
            If StartTimer = True Then Exit Sub
                If PauseGame = False Then
                    facingLeft = True: facingRight = False: facingUp = False: facingDown = False
                End If
        Case vbKeyRight
            If facingLeft Then Exit Sub
                If PauseGame = False Then
                    facingRight = True: facingLeft = False: facingUp = False: facingDown = False
                End If
        Case vbKeyUp
            If facingDown Then Exit Sub
                If PauseGame = False Then
                    facingUp = True: facingDown = False: facingLeft = False: facingRight = False
                End If
        Case vbKeyDown
            If facingUp = True Or PauseGame = True Then Exit Sub
                If PauseGame = False Then
                    facingDown = True: facingUp = False: facingLeft = False: facingRight = False
                End If
        Case vbKeyEscape
            If PauseGame = False Then
                        If TimeAttackMode = True Then
                            TimeLeft = False
                        End If
                Select Case True
                    Case facingLeft:
                        lastMove = "Left"
                        facingLeft = False: PauseGame = True
                        Exit Sub
                    Case facingRight:
                        lastMove = "Right"
                        facingRight = False: PauseGame = True
                        Exit Sub
                    Case facingUp:
                        lastMove = "Up"
                        facingUp = False: PauseGame = True
                        Exit Sub
                    Case facingDown:
                        lastMove = "Down"
                        facingDown = False: PauseGame = True
                        Exit Sub
                End Select
            ElseIf PauseGame = True Then
                        If TimeAttackMode = True Then
                            TimeLeft = True
                        End If
                Select Case lastMove
                    Case "Left"
                        facingLeft = True: PauseGame = False
                    Case "Right"
                        facingRight = True: PauseGame = False
                    Case "Up"
                        facingUp = True: PauseGame = False
                    Case "Down"
                        facingDown = True: PauseGame = False
                End Select
            End If
        End Select

End Sub

Private Sub GameLoop()

PauseGame = False

 Do While mainGameLoop
    DoEvents
            currtop = Snake(0).Top
            currleft = Snake(0).Left
        If (GetTickCount - lastTickCount) >= vbGameSpeed Then
            lastTickCount = GetTickCount
                    Select Case True
                        Case facingLeft:    'LEFT
                                Snake(0).Left = (Snake(0).Left - 240)
                            Call TailMove(currleft, currtop)
                            Call SnakeCollisions
                                If Snake(0).Top = ScoreBlock.Top And Snake(0).Left = ScoreBlock.Left Then
                                    PlayerScore = PlayerScore + 1
                                        If TimeAttackMode = True Then Form1.Caption = "Snake - Time Left: " & TimeControl.Caption & " - Score: " & PlayerScore
                                        If SurvivalMode = True Then Form1.Caption = "Snake - Score: " & PlayerScore
                                            Call GetNewFood
                                End If
                        Case facingRight:   'RIGHT
                                Snake(0).Left = (Snake(0).Left + 240)
                            Call TailMove(currleft, currtop)
                            Call SnakeCollisions
                                If Snake(0).Top = ScoreBlock.Top And Snake(0).Left = ScoreBlock.Left Then
                                    PlayerScore = PlayerScore + 1
                                    If TimeAttackMode = True Then Form1.Caption = "Snake - Time Left: " & TimeControl.Caption & " - Score: " & PlayerScore
                                    If SurvivalMode = True Then Form1.Caption = "Snake - Score: " & PlayerScore
                                        Call GetNewFood
                                End If
                        Case facingUp:      'UP
                                Snake(0).Top = (Snake(0).Top - 240)
                            Call TailMove(currleft, currtop)
                            Call SnakeCollisions
                                If Snake(0).Top = ScoreBlock.Top And Snake(0).Left = ScoreBlock.Left Then
                                    PlayerScore = PlayerScore + 1
                                        If TimeAttackMode = True Then Form1.Caption = "Snake - Time Left: " & TimeControl.Caption & " - Score: " & PlayerScore
                                        If SurvivalMode = True Then Form1.Caption = "Snake - Score: " & PlayerScore
                                            Call GetNewFood
                                End If
                        Case facingDown:    'DOWN
                                Snake(0).Top = (Snake(0).Top + 240)
                            Call TailMove(currleft, currtop)
                            Call SnakeCollisions
                                If Snake(0).Top = ScoreBlock.Top And Snake(0).Left = ScoreBlock.Left Then
                                    PlayerScore = PlayerScore + 1
                                        If TimeAttackMode = True Then Form1.Caption = "Snake - Time Left: " & TimeControl.Caption & " - Score: " & PlayerScore
                                        If SurvivalMode = True Then Form1.Caption = "Snake - Score: " & PlayerScore
                                            Call GetNewFood
                                End If
                        End Select
                Snake(Snake.Count - 1).FillColor = &H808080     'Updates the colour of the tail end
        End If
 Loop
End Sub

Private Sub Form_Load()

    PlayerScore = 0
    TimeLeft.Enabled = False
    StartTimer.Enabled = False
    LabelCount.Visible = False

    For i = 0 To 4
        Snake(i).Visible = False
    Next i
End Sub

Sub SnakeCollisions()
Dim ctr As Integer

    If Snake(0).Top < 0 Then Call GameOver
    If Snake(0).Top > 5160 Then Call GameOver
    If Snake(0).Left < 0 Then Call GameOver
    If Snake(0).Left > 5160 Then Call GameOver

    ctr = 1
    
    While ctr < Snake.Count - 1
        If Snake(ctr).Left = Snake(0).Left And Snake(ctr).Top = Snake(0).Top Then
                MsgBox ("Dead!")
                mainGameLoop = False
            Call GameOver
            Exit Sub
        End If
        ctr = ctr + 1
    Wend
End Sub

Private Sub HelpScreen_Click()
    HelpScreen.Visible = False
    
    TitleSnake.Visible = True
    StartTA.Visible = True
    StartSurvival.Visible = True
    lblScores.Visible = True
End Sub

Private Sub lblEnterName_Click()
        PlayerName = txtScoreName.Text
    Call ProcessScore(PlayerScore, PlayerName, TimeAttackMode, SurvivalMode)
    
    lblFinalScore.Visible = False
    lblEnterName.Visible = False
    txtScoreName.Visible = False
    LabelName.Visible = False
    lblGameOver.Visible = False
    
    SurvivalMode = False
    TimeAttackMode = False
    
    TitleSnake.Visible = True
    StartTA.Visible = True
    StartSurvival.Visible = True
    lblScores.Visible = True
    lblHelp.Visible = True
    
    txtScoreName.Text = ""
    Form1.Caption = "Snake"
    PlayerScore = 0
End Sub

Private Sub lblHelp_Click()
    TitleSnake.Visible = False
    StartTA.Visible = False
    StartSurvival.Visible = False
    lblScores.Visible = False
    
    HelpScreen.Visible = True
End Sub

Private Sub lblScores_Click()
    frmSplash.Show              'Shows the Highscores
End Sub

Private Sub StartSurvival_Click()

    Form1.Caption = "Snake - Score: 0"
    
    SurvivalMode = True
    TimeAttackMode = False
    
    LabelCount.Visible = True
    StartTimer.Enabled = True
    
    StartTA.Visible = False
    StartSurvival.Visible = False
    lblScores.Visible = False
    lblHelp.Visible = False
    TitleSnake.Visible = False
    
    For i = 0 To 4
        Snake(i).Visible = True
    Next i
End Sub

Private Sub StartTA_Click()

    Form1.Caption = "Snake - Time Left: 60 - Score: " & PlayerScore
    
    TimeAttackMode = True
    SurvivalMode = False
    
    LabelCount.Visible = True
    StartTimer.Enabled = True
    
    StartTA.Visible = False
    StartSurvival.Visible = False
    lblScores.Visible = False
    lblHelp.Visible = False
    TitleSnake.Visible = False
    
    For i = 0 To 4
        Snake(i).Visible = True
    Next i
End Sub

Private Sub TimeLeft_Timer()

TimeControl = TimeControl.Caption - 1

Form1.Caption = "Snake - Time Left: " & TimeControl.Caption & " - Score: " & PlayerScore
    If TimeControl.Caption = 0 Then
            TimeControl.Caption = 60
            mainGameLoop = False
        Call GetScoreResults(PlayerScore)
    End If
End Sub

Private Sub StartTimer_Timer()

    Snake(0).Visible = True
    LabelCount.Visible = True
    LabelCount.Caption = LabelCount.Caption - 1
    
    If LabelCount.Caption = 0 Then
        StartTimer.Enabled = False
        LabelCount.Caption = 3
        
            For Counter = 1 To 2
                Randomize
                    temptop = (Int(Rnd * 21) + 1) * 240
                    templeft = (Int(Rnd * 21) + 1) * 240
            Next Counter
            
          ScoreBlock.Top = temptop
          ScoreBlock.Left = templeft
          
        If SurvivalMode = True Then
            TimeLeft.Enabled = False
        ElseIf TimeAttackMode = True Then
            TimeLeft.Enabled = True
        End If
                mainGameLoop = True
                LabelCount.Visible = False
             
SetFocus: GameLoop

    End If
End Sub

Private Sub GetNewFood()
Dim max As Integer, temptop As Integer, templeft As Integer

    Randomize
    temptop = (Int(Rnd * 21) + 1) * 240
    templeft = (Int(Rnd * 21) + 1) * 240
    
    max = Snake.Count - 1
    
    For x = 0 To max
        If temptop = Snake(x).Top And templeft = Snake(x).Left Then
            Call GetNewFood
            Exit Sub
        End If
    Next x
    
    ScoreBlock.Top = temptop
    ScoreBlock.Left = templeft

    Snake(Snake.Count - 1).Top = Snake(Snake.Count - 2).Top
    Snake(Snake.Count - 1).Left = Snake(Snake.Count - 2).Left

    max = Snake.Count
    Load Snake(max)
    Snake(max).Left = templeft
    Snake(max).Top = temptop
    Snake(max).Visible = True
    
End Sub

Private Sub TailMove(ByRef currleft As Integer, ByRef currtop As Integer)
Dim ctr As Integer, templeft As Integer, temptop As Integer

ctr = 1

    While ctr < Snake.Count
        templeft = Snake(ctr).Left
        temptop = Snake(ctr).Top
        Snake(ctr).Left = currleft
        Snake(ctr).Top = currtop
        currleft = templeft
        currtop = temptop
        ctr = ctr + 1
    Wend
End Sub

Sub GetScoreResults(ByVal PlayerScore As Integer)
    lblFinalScore.Visible = True
    lblEnterName.Visible = True
    txtScoreName.Visible = True
    LabelName.Visible = True
    lblGameOver.Visible = True
    
    lblFinalScore.Caption = "Score: " & PlayerScore
End Sub

Sub GameOver()
    Form1.Caption = "Snake"
    TimeLeft.Enabled = False
    TimeControl.Caption = 60
    mainGameLoop = False
    
    facingRight = False
    facingUp = False
    facingLeft = False
    facingDown = False
    
    Snake(0).Top = 720
    Snake(0).Left = 2160
    ScoreBlock.Top = 720
    ScoreBlock.Left = 5520
    currtop = Snake(0).Top
    currleft = Snake(0).Left
    
    Dim max
    
    max = Snake.Count - 1
    
    For i = 5 To max
        Unload Snake(i)
    Next i
    
    For Counter = 1 To 4
        Snake(Counter).Top = Snake(Counter - 1).Top
        Snake(Counter).Left = Snake(Counter - 1).Left - 240
        Snake(Counter).Visible = False
    Next
    
    Snake(0).Visible = False
    
    Call GetScoreResults(PlayerScore)

End Sub

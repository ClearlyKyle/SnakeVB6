VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scores"
   ClientHeight    =   4635
   ClientLeft      =   7995
   ClientTop       =   3495
   ClientWidth     =   4635
   ClipControls    =   0   'False
   Icon            =   "frmLeaderboard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox ModeSelect 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "frmLeaderboard.frx":000C
      Left            =   1680
      List            =   "frmLeaderboard.frx":0016
      TabIndex        =   2
      Text            =   "Select Gamemode"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Scores 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Scores 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Scores 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   15
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Scores 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Scores 
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   13
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Names 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Names 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Names 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1080
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Names 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Posistion 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   480
      TabIndex        =   8
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Posistion 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Posistion 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Posistion 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Posistion 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Gamemode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Names 
      Caption         =   "Test Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Highscores"
      BeginProperty Font 
         Name            =   "Adobe Heiti Std R"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
LoadHighScore

    For i = 0 To 4
        Names(i).Caption = "--------"
        Scores(i).Caption = 0
    Next i

End Sub

Private Sub Label3_Click()
    Unload frmSplash
    Form1.Show
End Sub

Private Sub ModeSelect_Click()
Dim x As Integer

    If ModeSelect = "Time Attack" Then
    x = 1
        For i = 4 To 0 Step -1
            Names(i).Caption = ScoreRecord(x, i).ScoreName
            Scores(i).Caption = ScoreRecord(x, i).ScoreMark
        Next i
    ElseIf ModeSelect = "Survival" Then
    x = 2
        For i = 4 To 0 Step -1
            Names(i).Caption = ScoreRecord(x, i).ScoreName
            Scores(i).Caption = ScoreRecord(x, i).ScoreMark
        Next i
    End If
End Sub

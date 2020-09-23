VERSION 5.00
Begin VB.Form frmWinner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Timer step 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "&Restart"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "&Instructions"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuickRestart 
      Caption         =   "Back to The GAME!"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label lblScoreBoard 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmWinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim direction As Integer
Option Explicit

Private Sub cmdExit_Click()
    If MsgBox("Are You Sure You Dont Want To Keep Playing?", vbYesNo, "GoodBye Valued Player, :'(") = vbYes Then
        MsgBox "Bomberman, Remade using Visual basic" _
               & vbCrLf & "By Ian Gregor and Nicholas Hogan" _
               & vbCrLf & "See instructions for more details..."
        End
    End If
    'if they are dont want to keep playing, then give info and exit
End Sub

Private Sub cmdInstructions_Click()
    frmInstructions.Show    'show the instructions
    Unload Me
End Sub

Private Sub cmdRestart_Click()
    frmTitle.Show           'show the title
    Unload Me
End Sub

Private Sub cmdQuickRestart_Click()
    frmGame.Show            'back to the GAME!
    Unload Me
End Sub

Private Sub Form_Load()
Dim tempName As String
Dim pCount As Integer
    Unload frmGame          'stops the game form from running
    If winner <> -1 Then    'if someone wins then
        If PlayerNames(winner) <> "" Then
            'if the have a name then "Name Wins!"
            lblWinner.Caption = PlayerNames(winner) & " Wins!"
        Else
            'if they dont then their number is used instead
            lblWinner.Caption = "Player " & winner + 1 & " Wins!"
        End If
    Else
        'if noone wins then its definitaly a tie
        lblWinner.Caption = "Its A Tie!!"
    End If
    PlayerWins(winner) = PlayerWins(winner) + 1
    direction = 10  'starts moving the label up and down
    TotalGames = TotalGames + 1
    lblScoreBoard.Caption = "SCORE BOARD=-" & vbCrLf
    For pCount = -1 To NumOfPlayers - 1
        If pCount = 0 Then lblScoreBoard.Caption = lblScoreBoard.Caption & vbCrLf
        If PlayerNames(pCount) = "" Then
            tempName = "    Player " & (pCount + 1)
        Else
            tempName = "    " & PlayerNames(pCount)
        End If
        lblScoreBoard.Caption = lblScoreBoard.Caption & tempName & ": " & PlayerWins(pCount) & vbCrLf
    Next pCount
    lblScoreBoard.Caption = lblScoreBoard.Caption & vbCrLf & "Total Games: " & TotalGames
End Sub

Private Sub step_Timer()
    lblWinner.Top = lblWinner.Top + direction
    If lblWinner.Top > 600 Then direction = -10
    If lblWinner.Top < 0 Then direction = 10
    'moves the label up and down for added affect
    'was the easiest way to add a bit of fun to the
    'winner screen
End Sub

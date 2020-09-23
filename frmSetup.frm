VERSION 5.00
Begin VB.Form frmSetup 
   Caption         =   "Setting Up The Game!"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LstMode 
      Height          =   1620
      ItemData        =   "frmSetup.frx":0000
      Left            =   360
      List            =   "frmSetup.frx":000A
      TabIndex        =   29
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   6840
      TabIndex        =   28
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdInstr 
      Caption         =   "Instructions"
      Height          =   495
      Left            =   6840
      TabIndex        =   27
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtNameBox 
      Height          =   285
      Index           =   3
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "Name?"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtNameBox 
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   25
      Text            =   "Name?"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtNameBox 
      Height          =   285
      Index           =   1
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   24
      Text            =   "Name?"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtNameBox 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   23
      Text            =   "Name?"
      Top             =   810
      Width           =   1935
   End
   Begin VB.CommandButton cmdSetColours 
      Caption         =   "Set Colours"
      Height          =   375
      Left            =   1080
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      Height          =   495
      Index           =   8
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   7
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Height          =   495
      Index           =   6
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Height          =   495
      Index           =   5
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Height          =   495
      Index           =   4
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Height          =   495
      Index           =   0
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdColourChoice 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Height          =   495
      Index           =   1
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optPlayer 
      Caption         =   "Player 4"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Width           =   1000
   End
   Begin VB.OptionButton optPlayer 
      Caption         =   "Player 3"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1000
   End
   Begin VB.OptionButton optPlayer 
      Caption         =   "Player 2"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1000
   End
   Begin VB.OptionButton optPlayer 
      Caption         =   "Player 1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1000
   End
   Begin VB.TextBox txtKeyDefine 
      Height          =   495
      Index           =   4
      Left            =   3720
      TabIndex        =   8
      Text            =   "Bomb"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtKeyDefine 
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Text            =   "Down"
      Top             =   5880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtKeyDefine 
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   6
      Text            =   "Left"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtKeyDefine 
      Height          =   495
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Text            =   "Up"
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtKeyDefine 
      Height          =   495
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Text            =   "Right"
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSetKeys 
      Caption         =   "Set Keys"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdPlayGame 
      Caption         =   "Play The Game!"
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox cmbPlayers 
      Height          =   315
      ItemData        =   "frmSetup.frx":0027
      Left            =   1200
      List            =   "frmSetup.frx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Players"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim playerTemp As Integer

Private Sub cmdColourChoice_Click(Index As Integer)
    PlayerColours(playerTemp) = cmdColourChoice(Index).BackColor
    For i = 0 To cmdColourChoice.UBound
        cmdColourChoice(i).Visible = False
    Next i
    'sets the players colour and then hides the choosing buttons
End Sub

Private Sub cmdInstr_Click()
    Load frmInstructions
    frmInstructions.Show    'goes to the instructions form
    Unload Me
End Sub

Private Sub cmdSetColours_Click()
    For i = 0 To (cmdColourChoice.UBound)
        cmdColourChoice(i).Visible = True
    Next i
    'shows all the available colours
End Sub

Private Sub cmdSetKeys_Click()
    
    txtKeyDefine(0).Text = "RIGHT"
    txtKeyDefine(1).Text = "UP"
    txtKeyDefine(2).Text = "LEFT"
    txtKeyDefine(3).Text = "DOWN"
    txtKeyDefine(4).Text = "BOMB"
    
    For i = 0 To (txtKeyDefine.UBound)
        txtKeyDefine(i).Visible = True
    Next i
    'shows all the definable key boxes
End Sub

Private Sub cmdPlayGame_Click()
SetPlayerKeys
SetPlayerColours
NumOfPlayers = cmbPlayers.List(cmbPlayers.ListIndex)
SetNames
'sets the keys/colours/names and the number of players
frmGame.Show
Unload frmSetup
'shows the game!
End Sub

Private Sub Command1_Click()
    End ':( Limited Use Only...
    'ends the Game
End Sub

Private Sub Form_Load()
Dim pCount As Integer
cmbPlayers.ListIndex = 0    'selects the top choice in the combo box (4)
optPlayer(0).Value = True   'and the top option button
For pCount = 0 To 3
    If PlayerNames(pCount) <> "" Then txtNameBox(pCount) = PlayerNames(pCount)
Next pCount
LstMode.Selected(Mode) = True
End Sub

Private Sub SetNames()
For i = 0 To 3
    If PlayerNames(pCount) <> "" And PlayerNames(pCount) <> txtNameBox(pCount).Text Then PlayerWins(pCount) = 0
    PlayerNames(i) = UCase(Left(txtNameBox(i).Text, 1)) & LCase(Mid(txtNameBox(i).Text, 2))
Next i
'converts WiErD_cAsE to Sentence_case and then sets the name
checkNames  'checks the names in the module
End Sub

Private Sub LstMode_Click()
    Mode = LstMode.ItemData(LstMode.ListIndex)
End Sub

Private Sub optPlayer_Click(Index As Integer)
    playerTemp = Index
End Sub

Private Sub txtKeyDefine_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    txtKeyDefine(Index).Visible = False
    PlayerKeys(playerTemp, Index) = KeyCode
    'if a key is pressed then move to the next box...
    'and add that key to the players controls
End Sub


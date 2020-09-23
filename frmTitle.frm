VERSION 5.00
Begin VB.Form frmTitle 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   Picture         =   "frmTitle.frx":0000
   ScaleHeight     =   7170
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInstr 
      BackColor       =   &H000080FF&
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdSetup 
      BackColor       =   &H000080FF&
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H000080FF&
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInstr_Click()
    Load frmInstructions
    frmInstructions.Show
    Unload Me
    'shows the instructions form
End Sub

Private Sub cmdPlay_Click()
    SetPlayerKeys    'sets the keys to be used
    SetPlayerColours 'sets the players colours
    frmGame.Show    'shows the game form
    Unload Me
End Sub

Private Sub cmdSetup_Click()
    frmSetup.Show   'shows the setup form
    Unload Me
End Sub

Private Sub Form_Load()
    If NumOfPlayers = 0 Then NumOfPlayers = 4
    'if the number of players is previously undefined, then set it to 4..
    PlayerNames(-1) = "Draws"
End Sub

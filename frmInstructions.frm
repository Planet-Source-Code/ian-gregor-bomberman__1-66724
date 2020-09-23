VERSION 5.00
Begin VB.Form frmInstructions 
   Caption         =   "Form1"
   ClientHeight    =   8160
   ClientLeft      =   5055
   ClientTop       =   1845
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtInstr 
      Height          =   1215
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmInstructions.frx":0000
      Top             =   6840
      Width           =   6615
   End
   Begin VB.CommandButton cmdMoreInstr 
      Caption         =   ">> Next >>"
      Height          =   375
      Index           =   1
      Left            =   7200
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton cmdMoreInstr 
      Caption         =   "<< Previous <<"
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "How To Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgInstr 
      BorderStyle     =   1  'Fixed Single
      Height          =   6615
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6630
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim maxLesson As Integer        'the maximum amount of lessons
Dim LessonContent() As String   'an array that holds all the lessons
Dim Lesson As Integer           'the current lesson

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdMoreInstr_Click(Index As Integer)
    'since this button is next and previous, depending on which one was pressed
    'the next/prev info/pic will show...
    Lesson = (Lesson + (2 * Index) - 1) Mod maxLesson
    If Lesson < 0 Then Lesson = maxLesson - 1
    txtInstr.Text = "Lesson " & Lesson + 1 & ": " & LessonContent(Lesson)
    imgInstr.Picture = LoadPicture("Images/InstrImg/instrPic" & Lesson & ".jpg")
End Sub

Private Sub cmdPlay_Click()
    SetPlayerKeys       'sets the keys
    SetPlayerColours    'sets the colours
    Load frmGame        '
    frmGame.Show        'loads the game
    Unload Me           '
End Sub

Private Sub cmdSetup_Click()
    Load frmSetup
    frmSetup.Show   'Loads the Setup Screen
    Unload Me
End Sub

Private Sub Form_Load()
    'the help file is found at (CurDir&"\instr.txt")
    'and the images are found at (CurDir&"\Images\InstrImg\*.jpg")
    Dim Data As String  'data obtain from file
    ReDim LessonContent(0)  'sets the array to 0 length
    Open "instr.txt" For Input As #1    'opens the instruction file for reading(input)
    Do While Not EOF(1) 'loops through the file till the end is reached
        Line Input #1, Data         'gets a line of data
        maxLesson = maxLesson + 1   'increases the amount of lessons
        Data = Replace(Data, "#R", vbCrLf)  'changes all the #R in the instructions to new lines
        LessonContent(UBound(LessonContent)) = Data 'sets the info for the newest lesson
        ReDim Preserve LessonContent(UBound(LessonContent) + 1) 'increases the size of the lesson array
    Loop
    txtInstr.Text = "Lesson " & Lesson + 1 & ": " & LessonContent(Lesson)
    imgInstr.Picture = LoadPicture("Images/InstrImg/instrPic" & Lesson & ".jpg")
    Lesson = 0
    'sets the initial lesson and image
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close #1
    'closes the instruction file when the form is closed
End Sub

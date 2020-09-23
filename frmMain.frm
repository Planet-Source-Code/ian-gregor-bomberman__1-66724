VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Made By Ian Gregor. http://nkq0rs.spaces.live.com/"
   ClientHeight    =   8280
   ClientLeft      =   15045
   ClientTop       =   1830
   ClientWidth     =   11295
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Step 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdInstr 
      Caption         =   "Instructions"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End Game"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "Setup"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblStats 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4980
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   4260
   End
   Begin VB.Label lblPause 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PAUSED"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label lblDebug 
      Caption         =   "Label1"
      Height          =   1935
      Left            =   7440
      TabIndex        =   4
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label cmdPause 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************'
'DECLARATIONS'
'************'
Option Explicit                 'all variables must be declared

Private Type tGameMode
    bombInterval As Integer
    MaxPickups As Integer
    pickupChance As Integer
End Type

Dim GameMode As tGameMode

Private Type Pickup             'Pickup Type
    type As Integer             'what affect it has on the player
    age As Integer              'if it can be destroyed
    PickUpBody As Image         'the image of the pickup
    inUse As Boolean            'if the pickup is being used in the board
End Type

Dim pickupArray() As Pickup     'An array to Hold all the pickups
Dim otherPickup As Pickup       'used in collisions

'Player Declaration and extra player variables
Private Type player             'Player Type
    strName As String           'The Players Name
    nameTag As Label            'A Label that will follow the player around the board
    xCoord As Integer           'X Coordinate
    yCoord As Integer           'Y Coordinate
    bombs As Integer            'Amount of Bombs available
    power As Integer            'Power of the Bomb
    moveOk As Boolean           'if it is ok for the player to mover
    attached(1) As Integer         'the player that this player is attached to as a ghost
    direction(1) As Integer     'Direction moving
    lives As Integer            'amount of lives left...
    body As Shape               'The Body Shape
    dieOk As Integer            'If it is OK to die (so one bomb only takes one life
    lenDirX As Integer          'length in the direction of the xCoordinate
    lendirY As Integer          'length in the direction of the yCoordinate
    keysDefined(4) As Integer   'The keys that the player used, can be defined in Setup form
    skullState(1) As Integer    '0-timer, 1-type
End Type

Dim stepcoords() As Integer
Dim dist(5, 2) As Integer
Const DANGER = 1, SAFE = 0

'length in direction of the X and Y coordinates, used
'to reduce "if statements" by using sin/cos instead of
'checking the direction and adding to them specifically
Dim lenDirX As Integer
Dim lendirY As Integer

Dim ID As Integer           'The ID of the selected player
Dim playerArray() As player 'an Array to contain the players
Dim otherPlayer As player   'the other player in collisions

'bomb Declaration and extra player variables
Private Type bomb           'Bomb Type
    ownerID As Integer      'The Bombs owner
    ID As Integer           'The bombs Id
    inUse As Boolean        'If the Bomb is in use
    age As Integer          'countdown timer to destruction
    power As Integer        'area of affect
    direction As Integer    'if it has been kicked
    moveOk As Boolean       'if it is ok for the bomb to move
    type As Integer         'what type of bomb is it, (Standard is Default)
    BombBody As Shape       'the body shape
    explosionX As Shape     'the explosions X shape
    explosionY As Shape     'the explosions Y shape
    xCoord As Integer       'X Coordinate
    yCoord As Integer       'Y Coordinate
    lenDirX As Integer      'length in the direction of the xCoordinate
    lendirY As Integer      'length in the direction of the yCoordinate
End Type

Dim bombCount As Integer    'Amount of Bombs,
Dim bombArray() As bomb     'an Array to contain the bombs
Dim otherBomb As bomb       'the other bomb in collisions

'destructable brick peices
Private Type brick          'Brick Type
    ID As Integer           'The brick's ID#
    age As Integer          'Age to develop
    inUse As Boolean        'if the brick is in use
    BrickBody As Shape      'the body shape
End Type

Dim brickArray() As brick       'an Array to contain the bricks
Dim otherBrick As brick         'the other brick in collisions
Dim brickInterval As Integer    'the variable interval at which bricks are created. (proportional to bricks)
Dim GhostRemInterval As Integer

'board peice declarations
Dim square() As Shape       'the background of the board
Dim squareCount As Integer  'the amount of Squares on the board
Dim boardSquare As Shape    'the square shape

'extra information
Dim gameRunning As Boolean  'if the game is running or not... unused atm
Dim temp As Variant         'Temporary Variable

Dim pCount As Integer       'player Count, used in "For" loops

'mmm pi...
Const Pi = 3.14159265358979

'direction Constants (basically degrees / 90)
'these are used for moving shapes and types
Const LEFT_ = 2
Const UP_ = 1
Const RIGHT_ = 0
Const DOWN_ = 3
Const STILL_ = -1

'Game Constants
Const WIDTH_ = 6 * 100      'the width of each piece [players,squares,bricks] Etc
Const BOARDWIDTH_ = 10      'the width of the board
Const MAXITEMS_ = 8         'the max power/bombs that a player can aquire

'BackStyle Constants
Const BACKTRANS_ = 0        'the vbConstants for solid and...
Const BACKOPAQUE_ = 1       '...clear are wrong for back style

'Collision Constants
'The object causing the collision will react different to each returned value
Const NOTHING_ = 0          'No Reaction
Const BOUNDARY_ = 1         'Stops the object, no questions asked
Const WALL_ = 1             'As above
Const BRICK_ = 2            'Stops the object, and is destroyed if the object is fire
Const PLAYER_ = 3           'Stops the object, unless it is fire
Const BOMB_ = 4             'Kicks the bomb, if fire, sets the bomb's age to death age
Const PICKUP_ = 5           'Stops nothing, and is destroyed

'pickup Constants
'pickup Types
Const puBOMBUP = 0          'is used to create a bomb pickup
Const puFIREUP = 1          'is used to create a power pickup
Const puLIFEUP = 2          'is used to create a Life pickup
Const puSkull = 3
    'Skull types
    Const skBombs = 1
    Const skShortBomb = 2
    Const skSlow = 3
    Const skDirection = 4
    Const skInvisible = 5
    Const skSwap = 6
    Const skMagnet = 7
    Const skFast = 8
    Const skPower = 9
    Const skLife = 10
    Const skDragon = 11
    'ratio (-:~:+) (5:2:4)
    Const skMax = 11
'Const puKickUP = 3         'unused at the moment

'Collision Action Constants
'Const NOTHING_ = 0         'already declared
 Const DESTROY_ = 1         'on collision, destroy that object
 
'Brick Constants
Const ageStep_ = 50             'intervals between age Steps (visibility and solidity)
Const ageColorMin_ = vbWhite    'what colour it begins at
Const ageColorMax_ = vbBlack    'what colour it ends at

'Bomb Type Constants
Const NORMAL_ = 0           'Normal type Bombs
Const FIREBALL = 1
'bomb life length Constants
Const lifeLONG_ = 500       'Long Life Length
Const lifeMEDIUM_ = 250     'Medium life Length
Const lifeSHORT_ = 100      'Short life length
Const bombDEATH_ = 50       'The age at which the bomb starts to explode

'Some Colours, Extra undefined Colours
Const vbLightBlue = &HFF8080
Const vbLightGrey = &HC0C0C0
Const vbGrey = &H808080

Private Sub cmdEnd_Click()
    'creates a msgbox telling that the game will quit if they continue
    If MsgBox("This Will End Your Current Game", vbOKCancel) = vbOK Then
        winner = -1
        Load frmWinner  'loads the winner form
        frmWinner.Show  'shows the winner form
        Unload Me       'unloads this one so it stops running
    End If
End Sub

Private Sub cmdInstr_Click()
    'creates a msgbox telling that the game will quit if they continue
    If MsgBox("This Will End Your Current Game", vbOKCancel) = vbOK Then
        Load frmInstructions    'loads the instruction form
        frmInstructions.Show    'shows the instruction form
        Unload Me               'unloads this form so the game stops running
    End If
End Sub

Private Sub cmdSetup_Click()
    'creates a msgbox telling that the game will quit if they continue
    If MsgBox("This Will End Your Current Game", vbOKCancel) = vbOK Then
        
        Load frmSetup   'loads the setup form
        frmSetup.Show   'shows the setup form
        Unload Me       'unloads this form so the game stops running
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim bombTest As Byte        'is used to determine if a bomb exists where your are standing
Dim gameTest As Byte        'they arent going to go near 255...
Dim iCount As Integer, dCount As Byte, dirCount As Byte

For ID = 0 To UBound(playerArray)   'starts a loop going through the players
    With playerArray(ID)                'Opens a with Block for the Player Selected
        For dirCount = RIGHT_ To DOWN_
            If .keysDefined(dirCount) = KeyCode Then
                For dCount = 0 To 1
                    .direction(dCount) = dirCount
                    .attached(1) = 1
                    If .attached(0) <> -1 Then
                        playerArray(.attached(0)).direction(dCount) = dirCount
                        playerArray(.attached(0)).attached(1) = 0
                    End If
                Next dCount
            End If
        Next dirCount
        Select Case KeyCode
            Case .keysDefined(BOMB_):   'if the key for Bomb is pressed...
                bombTest = 0            'Sets Test to 0 as default
                gameTest = 0            'Sets Test to 0 as default
                If .bombs > 0 And .lives <> 0 And (.skullState(1) <> skMagnet) Then
                    For iCount = 0 To UBound(bombArray)
                        If bombArray(iCount).xCoord = .xCoord And bombArray(iCount).yCoord = .yCoord Then
                            bombTest = bombTest + 1
                        End If
                        'if a collision occurs, add one to test
                    Next iCount             'loops through itteration count
                    For pCount = 0 To UBound(playerArray)
                        If playerArray(pCount).lives >= 1 Then gameTest = gameTest + 1
                    Next pCount
                    If bombTest = 0 And gameTest > 1 Then  'if no collisions were found, and at least 2p are alive
                        makeBomb ID         'make a bomb with the ownerID as the player ID
                        .bombs = .bombs - 1 'decrease the players bombs
                    End If
                ElseIf .lives <= 0 And .attached(0) <> -1 Then
                    If playerArray(.attached(0)).bombs > 0 Then
                        For iCount = 0 To UBound(bombArray)
                            If bombArray(iCount).xCoord = .xCoord And bombArray(iCount).yCoord = .yCoord Then
                                bombTest = bombTest + 1
                            End If
                            'if a collision occurs, add one to test
                        Next iCount             'loops through itteration count
                        For pCount = 0 To UBound(playerArray)
                            If playerArray(pCount).lives >= 1 Then gameTest = gameTest + 1
                        Next pCount
                        If bombTest = 0 And gameTest > 1 Then  'if no collisions were found, and at least 2p are alive
                            makeBomb .attached(0)          'make a bomb with the ownerID as the player ID
                            playerArray(.attached(0)).bombs = playerArray(.attached(0)).bombs - 1   'decrease the players bombs
                        End If
                    End If
                End If
        End Select
    End With
Next ID
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim test As Integer     'is used to determine if a bomb exists where your are standing
    Dim kCount As Integer   'Key Count
    For pCount = 0 To UBound(playerArray) 'loops through player count
        For kCount = RIGHT_ To DOWN_ 'anticlockwise
            If playerArray(pCount).keysDefined(kCount) = KeyCode Then
                playerArray(pCount).direction(0) = STILL_
                If playerArray(pCount).attached(0) <> -1 Then
                    playerArray(playerArray(pCount).attached(0)).direction(0) = STILL_
                End If
            End If
            'if the key released is one of the defined keys then set the players direction to still
        Next kCount
    Next pCount
End Sub

Private Sub Form_Load()
    
Dim sCount As Integer           'declares the square count
Dim kCount As Integer   'declares key loop variable
    'Determines the amount of squares by rounding the
    'boardwidth down to the nearest even number, then
    'it adds 2, halves it then squares it.
    ' Example: boardwidth = 5
    '          5-1+2 = 6
    '          6/2 = 3
    '          3^2 = 9 squares
    squareCount = (Int(BOARDWIDTH_ - (BOARDWIDTH_ Mod 2) + 2) / 2) ^ 2
    ReDim bombArray(0)      'sets the bomb array to 0 length
    ReDim brickArray(0)     'sets the brick array to 0 length
    ReDim pickupArray(0)    'sets the pickup array to 0 length
    
    PickupDir(puBOMBUP) = "pickup_bombup.bmp"   'set the names for the images...
    PickupDir(puFIREUP) = "pickup_Fireup.bmp"   'of the pickups, which are...
    PickupDir(puLIFEUP) = "pickup_lifeup.bmp"   'in the images folder...
    PickupDir(puSkull) = "pickup_skull.bmp"

    ReDim Preserve square(squareCount)  'Sets the square array to hold the designated number
    For sCount = 0 To squareCount            'loops through all the Squares
        Set square(sCount) = Me.Controls.Add("VB.Shape", "square" & sCount)
        'declares the shape in the form controls
        With square(sCount)
            .Shape = vbShapeRoundedSquare   'Rounded Square for ascetics
            .BackColor = vbBlue             'Blue is a nice colour
            .BackStyle = BACKOPAQUE_        'opaque back style
            .Height = WIDTH_                'determines Height based on constant
            .Width = WIDTH_                 'determines Width  based on constant
            .Visible = True                 'Makes the shape Visible
        End With
    Next sCount
    
    SetBoard                'Put the squares in the appropriate places
    SetMode
    
    ReDim Preserve playerArray(NumOfPlayers - 1) As player
    '^^Redimensions the Array to hold all the players, can be expanded beyond 4 for future development
    
    
    For pCount = 0 To UBound(playerArray)    'loops through the players
        With playerArray(pCount)
            'determines their position
            .xCoord = WIDTH_ * ((BOARDWIDTH_ - 1) * (pCount Mod 2)) + (WIDTH_ * ((pCount + 1) Mod 2))
            .yCoord = (BOARDWIDTH_ - 2) * (WIDTH_ * (Int(pCount / 2))) + WIDTH_
            .bombs = 1                  'sets the initial bomb count
            .power = 1                  'set the initial power level
            .lives = 1                  'set the initial lives to 1
            .direction(0) = STILL_         'init direction is Still
            .moveOk = True              'it ok to move
            .strName = PlayerNames(pCount)   'test name
            .skullState(0) = 0
            .skullState(1) = 0
            
            For kCount = 0 To 4         'each of the keys: right up left down, bomb keys
                .keysDefined(kCount) = PlayerKeys(pCount, kCount)
                'defines the keys from the setup
            Next kCount
            
            Set .nameTag = Me.Controls.Add("VB.Label", "NameTag" & pCount)
            'adds the name tag
            .nameTag.Caption = .strName 'sets the name of the nametag
            .nameTag.Top = .yCoord           'aligns the name with the top of the player
            .nameTag.Left = .xCoord + WIDTH_ 'Sets the position of the nametag to the right of the player
            .nameTag.Font = "Times New Roman"
            With .nameTag
                .AutoSize = True        'makes the tag the size of the name
                .BackStyle = 0          'Transparent
                .Visible = True         'shows the nametag
            End With
            Set .body = Me.Controls.Add("VB.Shape", "body" & pCount)
            '^^adds the bodyshape to the form controls
            With .body
                .BackColor = PlayerColours(pCount)
                'Sets the backcolour to the previously defined colour
                .BackStyle = BACKOPAQUE_    'sets the back style to be Solid/Opaque
                .Top = playerArray(pCount).yCoord     '\ sets the position
                .Left = playerArray(pCount).xCoord    '/
                .Visible = True             'Makes the player visible
                .Width = WIDTH_             'Sets height and width based on
                .Height = WIDTH_            ' based on the constant
                .Shape = vbShapeCircle      'Sets the shape to be a circle
            End With 'Body
        End With 'player
    Next pCount 'next player
    
    Set boardSquare = Me.Controls.Add("VB.Shape", "boardSquare")
    'adds the background Square to the form controls
    With boardSquare
        .Width = (BOARDWIDTH_ + 1) * WIDTH_     'Sets the width and height...
        .Height = (BOARDWIDTH_ + 1) * WIDTH_    'to fit behind the squares.
        .BackStyle = BACKOPAQUE_                'sets the back style to solid
        .BackColor = vbLightBlue                'and the colour to light blue
        .ZOrder (1)                             'sends the shape to the back
        .Visible = True                         'makes it visible
    End With
    step.Interval = 1
End Sub

Private Sub SetBoard()
'This subroutine loops through all the possible X and Y locations of
'the squares on the board. it then determines which square should be
'where and sets the positions.
Dim xCoord As Integer
Dim yCoord As Integer
Dim squareID As Integer
    For xCoord = 0 To BOARDWIDTH_ Step 2
        For yCoord = 0 To BOARDWIDTH_ Step 2
            squareID = (yCoord / 2) + (xCoord / 2 * Int((BOARDWIDTH_ + 2) / 2))
            With square(squareID)
                .Left = xCoord * WIDTH_
                .Top = yCoord * WIDTH_
            End With
        Next yCoord
    Next xCoord
End Sub

Private Sub cmdPause_Click()
'Because everything runs off the step timer, stopping the timer will
'effectively stop everything. The command buttons are hidden when the
'game is running because they interfere with the keypresses, the pause
'button is a clickable label, which doesnt interfere.
    If step.Enabled = True Then
        step.Enabled = False
        lblPause.Visible = True
        lblPause.ZOrder 0
        cmdInstr.Visible = True
        cmdEnd.Visible = True
        cmdSetup.Visible = True
    Else
        step.Enabled = True
        lblPause.Visible = False
        cmdInstr.Visible = False
        cmdEnd.Visible = False
        cmdSetup.Visible = False
    End If
    
End Sub

'##########
'#/      \#
'#| STEP |#
'#\______/#
'##########
Private Sub step_Timer() 'A timer with a 1ms interval
Dim bombCheck
Dim brickCount As Integer
Dim bCount
    Randomize
    SkullEffects
    PlayerMovement      'Moves the players if they need moving
    bombMovement        'moves the bombs that have been kicked
    NameTagMovement
    brickAge            'ages the bricks, so that they develop
    SetMode
    GhostAbilities
    
    'runs the Tag integer to 1000 then drops it to 0 and loops
    'the tag is used as a delay on events, such as brick creation
    brickInterval = brickInterval + 1

    If brickInterval Mod GameMode.bombInterval = 0 Then
        addBrick            'creates a brick
        brickInterval = 0   'resets brick interval to 0
    End If
    
    For pCount = 0 To UBound(playerArray)
        With playerArray(pCount)
            If .dieOk > 0 Then .dieOk = .dieOk - 1
            'since the bomb destroys everything in its range _every step_ then a limit between dieing is needed
            If .bombs > MAXITEMS_ Then .bombs = MAXITEMS_
        End With
    Next pCount

    For bCount = 0 To UBound(bombArray)
        With bombArray(bCount)
            'if age is above zero then lower it
            If .age >= 0 Then .age = .age - 1
            'if age is below deathAge and the bomb is inuse, then explode it
            If .age <= bombDEATH_ And .inUse = True Then bombExplode .ID
        End With
    Next bCount
    'this removes redundant bombs from the end of the array.
    If bombArray(UBound(bombArray)).inUse = False And UBound(bombArray) > 0 Then
        ReDim Preserve bombArray(UBound(bombArray) - 1)
    End If
    'isnt nessecary for 4 players!

    'this checks that all the bombs have exploded, if they are, then check to see if the game has been won.
    'if there is only one player, then it is impossible to plant bombs
    bombCheck = 0
    For bCount = 0 To UBound(bombArray)
        If bombArray(bCount).inUse = True Then bombCheck = bombCheck + 1
    Next bCount
    If bombCheck = 0 Then WinCheck
End Sub

Private Sub PlayerMovement()
    Dim tempColl As Integer
    Dim tempMove As Integer
    Dim pCount As Integer
    For pCount = 0 To UBound(playerArray) 'starts a loop for all players
        With playerArray(pCount)
            If .moveOk = True Then      'checks if player is ok to move
                .lendirY = Int(-Sin(.direction(0) * 90 * Pi / 180) * WIDTH_)    'uses Sin/Cos to convert direction...
                .lenDirX = Int(Cos(.direction(0) * 90 * Pi / 180) * WIDTH_)     '...to x and y in the directon.
                .lendirY = .lendirY - (.lendirY Mod WIDTH_)    'because pi isnt accurate, we need...
                .lenDirX = .lenDirX - (.lenDirX Mod WIDTH_)    '...to take off the trailing fraction
                If .direction(0) = STILL_ Then
                    .lenDirX = 0     'if lenDir is calculated using STILL_ (-1)...    | cos (-90^) = 0
                    .lendirY = 0     '...it becomes downwards, and we dont want that  | sin (-90^) = -1
                End If
                If .lives = 0 And .attached(0) <> -1 And .attached(1) = 0 Then
                    temp = 1
                Else
                    temp = .lives
                End If
                tempColl = Collisions(pCount, .direction(0), .xCoord, .yCoord, , Int(temp), PLAYER_)
                'since the value is used more than once, it is nessecary to store it
                If tempColl = NOTHING_ Or tempColl = Left(PICKUP_, 1) Then
                    .xCoord = .xCoord + .lenDirX   'updates the .xCoord and .yCoord variables to
                    .yCoord = .yCoord + .lendirY   'where they should be after a button press
                End If
                If Left(tempColl, 1) = PICKUP_ Then 'if the player runs into a pickup
                    With playerArray(pCount)
                    Select Case Mid(tempColl, 2)
                        Case puBOMBUP 'increase the bombs if less than maxitems
                            If .bombs < GameMode.MaxPickups Then .bombs = .bombs + 1
                            
                        Case puFIREUP 'increase the power if less than maxitems
                            If .power < GameMode.MaxPickups Then .power = .power + 1
                        Case puLIFEUP 'increase the lives if less than 2
                            If .lives < 2 Then .lives = .lives + 1
                            playerDestroy pCount, False
                        Case puSkull
                            .skullState(0) = 1
                            SkullEffects
                            .skullState(0) = 1500
                            .skullState(1) = Int(Rnd * skMax) + 1
                            'MsgBox .skullState(1)
                    End Select
                    End With
                End If
            End If
            'checks if x or y is more or less than the body's left and top
            'then moves the body 1/10 of width to slowly creep it across
            'the board instead of it jumping. When the x and y are equal
            'it is ok to move again, otherwise it isnt
            tempMove = (10 + (Abs(Int(.skullState(1) = skSlow)) * 20) - (Abs(Int(.skullState(1) = skFast)) * 8))
            If .xCoord <> .body.Left Then
                .moveOk = False
                .body.Left = .body.Left + .lenDirX / tempMove
            End If
            If .yCoord <> .body.Top Then
                .body.Top = .body.Top + .lendirY / tempMove
                .moveOk = False
            End If
            If .xCoord = .body.Left And .yCoord = .body.Top And .moveOk = False Then
                .moveOk = True
                UpdateStats
            End If
            
        End With
    Next pCount
End Sub

Private Function Collisions( _
        intID As Integer, _
        direction As Integer, _
        collX As Integer, _
        collY As Integer, _
        Optional action As Integer = 0, _
        Optional lives As Integer = -1, _
        Optional myType As Integer = -1)

Dim lenDirX As Integer
Dim lendirY As Integer
Dim tempColl As Integer, tempSkull(1) As Integer
Dim bCount As Integer, pCount As Integer

    lendirY = Int(-Sin(direction * 90 * Pi / 180) * WIDTH_) 'Detemines the lenght in the
    lenDirX = Int(Cos(direction * 90 * Pi / 180) * WIDTH_)  'direction stated by the function
    lendirY = lendirY - (lendirY Mod WIDTH_)    'removes the trailing
    lenDirX = lenDirX - (lenDirX Mod WIDTH_)    'innacurate numbers...
    If direction = STILL_ Then
        lendirY = 0     'sets the length to 0 if there is no direction
        lenDirX = 0     '^^above
    End If
    'block collisions
    If lives > 0 Or lives < 0 Then 'if the player is not dead or no lives are stated...
        If ((collX + lenDirX) / WIDTH_) Mod 2 = 0 And ((collY + lendirY) / WIDTH_) Mod 2 = 0 Then
        'check if there should be a block in the way of the object
            Collisions = WALL_ 'returns WALL_ from the function
            Exit Function   'exits the function so no more collisions are tested
        End If
        For bCount = 0 To UBound(brickArray)
            otherBrick = brickArray(bCount) 'sets OtherBrick as a possible collision brick
            If (otherBrick.inUse = True) Then
                If (otherBrick.BrickBody.FillColor = vbBlack Or lives = -2) Or action = DESTROY_ Then
                'if otherbrick if fully made or is going to be destroyed
                    If collX + lenDirX = otherBrick.BrickBody.Left And collY + lendirY = otherBrick.BrickBody.Top Then
                        'if both xCoords and yCoords are the same
                        Collisions = BRICK_     'returns  BRICK_ from the function
                        If action = DESTROY_ Then brickDestroy otherBrick.ID
                        'if the brick is going to be destroyed, the call BrickDestroy
                        Exit Function   'exits the function so no more collisions are tested
                    End If
                End If
            End If
        Next bCount
        
        For pCount = 0 To UBound(playerArray)
            otherPlayer = playerArray(pCount)   'sets otherplayer as a possible collision player
            If (myType = PLAYER_ And pCount <> ID) Or myType = -1 Then
            If collX + lenDirX = otherPlayer.xCoord And collY + lendirY = otherPlayer.yCoord And otherPlayer.lives > 0 Then
            'if xCoords and yCoords are the same and the player is not a ghost
                If otherPlayer.lives > 0 Then
                'if the other player has more than 0 lives
                    Collisions = PLAYER_ & pCount 'returns Player and the player id
                    If myType = PLAYER_ And intID <> pCount Then
                        With playerArray(pCount)
                            If playerArray(intID).skullState(1) <> 0 And (.dieOk = 0) And (playerArray(intID).dieOk = 0) Then
                                tempSkull(0) = .skullState(0)
                                tempSkull(1) = .skullState(1)
                                
                                .skullState(0) = 1
                                SkullEffects
                                
                                .skullState(0) = playerArray(intID).skullState(0)
                                .skullState(1) = playerArray(intID).skullState(1)
                                
                                playerArray(intID).skullState(0) = 1
                                SkullEffects
                                
                                playerArray(intID).skullState(0) = tempSkull(0)
                                playerArray(intID).skullState(1) = tempSkull(1)
                                
                                playerArray(intID).dieOk = 10
                                .dieOk = 10
                            End If
                        End With
                    End If
                Else
                    Collisions = NOTHING_ 'returns no collisions
                End If
                If action = DESTROY_ Then playerDestroy pCount
                'destroys the player if action is destroy
                Exit Function 'exits the function so no more collisions are tested
            End If
            End If
        Next pCount
    End If
    
    If (collX + lenDirX < 0) Or (collY + lendirY < 0) _
    Or (collX + lenDirX > BOARDWIDTH_ * WIDTH_) Or (collY + lendirY > BOARDWIDTH_ * WIDTH_) Then
    'if the collider is inside the board boundaries
        Collisions = BOUNDARY_ 'returns BOUNDARY_ from the function
        Exit Function   'exits so no more collisions are tested
    End If
    
    If lives > 0 Or lives < 0 Then 'if lives are more than 0 or not counted
        For bCount = 0 To UBound(bombArray)
            otherBomb = bombArray(bCount) 'sets the otherBomb for possible bomb collisions
            If collX + lenDirX = otherBomb.xCoord And collY + lendirY = otherBomb.yCoord Then
            'if the xCoords and yCoords of the bomb and collider are the same
                Collisions = BOMB_ 'returns BOMB_ from the function
                If ((action = DESTROY_) And (otherBomb.type <> FIREBALL)) And (otherBomb.age > bombDEATH_) Then bombArray(bCount).age = bombDEATH_ + 1
                'if the otherBomb is to be destroyed and its age is more than deathAge, then age = deathAge
                If direction <> STILL_ Then tempColl = Collisions(Int(bCount), direction, otherBomb.xCoord, otherBomb.yCoord)
                'if collider is moving then check for another bomb in the current direction
                If tempColl = BOMB_ Then
                'if the collision is a bomb then
                    bombArray(bCount).direction = STILL_
                Else
                    If lives <> 0 Then bombArray(bCount).direction = direction
                    'checks if lives is not 0 then moves the otherbomb
                End If
                Exit Function 'exits to function so no more collisions are checked
            End If
        Next bCount
    End If
    For pCount = 0 To UBound(pickupArray)
        otherPickup = pickupArray(pCount)
        'sets the otherbomb for possible bomb collisions
        If otherPickup.inUse = True Then
            If collX + lenDirX = otherPickup.PickUpBody.Left _
            And collY + lendirY = otherPickup.PickUpBody.Top Then
            'if the xCoords and yCoords are the same
                If action = DESTROY_ Then
                    'bomb explosions action are destroy
                    If otherPickup.age > 50 Then pickupDestroy pCount
                    'destroys the pickup if the age is more than 50
                ElseIf lives > 0 Or lives = -1 Then
                    pickupDestroy Int(pCount)
                    'destroys the pickup if is being run into
                End If
                If lives > 0 Or lives = -1 Or otherPickup.type = puLIFEUP Then
                    Collisions = PICKUP_ & otherPickup.type
                End If
                'returns Pickup_ and the type from the function
                Exit Function
            End If
        End If
    Next pCount
            
    Collisions = NOTHING_
    'if there are no collisions after all that.... return Nothing!
End Function

Public Sub makeBomb(dropperID As Integer)
    Dim bCount As Integer
    temp = -1
    'temp is used for testing which bombs are available
    For bCount = 0 To UBound(bombArray) 'for all the bombs
        If bombArray(bCount).inUse = False Then
            'if the bomb isnt inUse then set temp to the bomb number
            temp = bCount
            bCount = UBound(bombArray) + 1
            'break the loop
        End If
    Next bCount
    'if the loop above didnt find a bomb then temp will still be -1
    If temp = -1 Then
        temp = UBound(bombArray) + 1
        ReDim Preserve bombArray(UBound(bombArray) + 1)
        'sets temp to an undefined bomb and expands the array
    End If
    
    With bombArray(temp)
        .ownerID = dropperID    'sets the ownerID to the DropperID
        .ID = temp              'ID is temp
        .age = lifeMEDIUM_ + (100 * (playerArray(dropperID).skullState(1) = skShortBomb))
        .direction = STILL_     'stops the bomb (0 is default, 0 is right)
        .inUse = True
        .type = NORMAL_         'sets type to normal (only one type at the moment
        .xCoord = playerArray(dropperID).xCoord 'moves the bomb to the
        .yCoord = playerArray(dropperID).yCoord 'droppers coordinates
        .moveOk = True          'the bomb is able to move
        Set .BombBody = Me.Controls.Add("VB.Shape", "bombBody" & .ID)
        Set .explosionX = Me.Controls.Add("VB.Shape", "explosionX" & .ID)
        Set .explosionY = Me.Controls.Add("VB.Shape", "explosionY" & .ID)
        'adds the bombs objects to the form controls
        If playerArray(.ownerID).skullState(1) = skPower Then
            .BombBody.BackColor = vbRed
            .power = BOARDWIDTH_ + 1
        ElseIf playerArray(.ownerID).skullState(1) = skDragon Then
            .direction = playerArray(dropperID).direction(1)
            .BombBody.BackColor = vbRed
            .BombBody.FillStyle = vbDiagonalCross
            .BombBody.FillColor = vbYellow
            .type = FIREBALL
            .power = 0
        Else
            .BombBody.BackColor = vbLightGrey
            .power = playerArray(.ownerID).power
        End If
        With .BombBody
            .BackStyle = BACKOPAQUE_    'styles of the bomb shape
            .BorderColor = vbGrey       '
            .Left = playerArray(dropperID).xCoord + WIDTH_ / 4  'sets the body position
            .Top = playerArray(dropperID).yCoord + WIDTH_ / 4   'to the middle of x/yCoord
            .Visible = True             'makes the body visible
            .Width = WIDTH_ / 2         'sets the height and width
            .Height = WIDTH_ / 2        'to half of default width
            .Shape = vbShapeCircle      'sets the shape to a circle
            .ZOrder (0)                 'brings the bomb to the front
        End With
        
        With .explosionX
            .BackColor = vbRed                  'makes it red
            .BackStyle = BACKOPAQUE_            'solid backstyle
            .Shape = vbShapeRoundedRectangle    'shape is a rounded rectangle
            .Height = WIDTH_                    '
            .BorderStyle = 0                    'No border style
            .ZOrder (0)                         'brings to front
        End With
        
        With .explosionY
            .BackColor = vbRed                  'makes it red
            .BackStyle = BACKOPAQUE_            'solid backstyle
            .Shape = vbShapeRoundedRectangle    'shape is a rounded rectangle
            .Width = WIDTH_                    '
            .BorderStyle = 0                    'No border style
            .ZOrder (0)                         'brings to front
        End With
        
    End With
End Sub

Private Sub SetMode()
Dim brickCount As Integer
Dim bCount
    Select Case Mode
        Case 0:
            brickCount = 0
            For bCount = 0 To UBound(brickArray)
                If brickArray(bCount).inUse = True Then brickCount = brickCount + 1
            Next bCount
            GameMode.bombInterval = (50 + Int(brickCount / 10) * 150)
            
            GameMode.MaxPickups = 8
            
            GameMode.pickupChance = Int(Rnd * 2)
        Case 1:
            GameMode.bombInterval = 25
            
            GameMode.MaxPickups = 100
            
            GameMode.pickupChance = 1
    End Select
End Sub

Private Sub bombExplode(intID As Integer)
    Dim fireCollision As Integer
    Dim range As Integer
    Dim iCount As Integer
    Dim powerCheck As Boolean
    Dim tempxCoord As Integer, tempyCoord As Integer
    
    With bombArray(intID)
        .BombBody.Visible = False       'makes the bomb invisible
        
        .explosionX.Visible = True      'makes the explosions visible
        .explosionY.Visible = True      '^^
        If .age = bombDEATH_ Then       'if age is deathAge
            .explosionX.Left = .xCoord  'explosionX left = xCoord
            .explosionX.Top = .yCoord   'explosionX Top = yCoord
            .explosionX.Width = WIDTH_  '
            
            .explosionY.Left = .xCoord  'explosionY left = xCoord
            .explosionY.Top = .yCoord   'explosionY Top = yCoord
            .explosionY.Height = WIDTH_ '
            powerCheck = (playerArray(.ownerID).skullState(1) <> skPower)
            For iCount = -1 To 1 Step 2
                For range = 1 To .power
                    fireCollision = Collisions(.ID, STILL_, .xCoord + range * iCount * WIDTH_, .yCoord, DESTROY_)
                    'destroys anything explosionX hits
                    If iCount = -1 Then
                        .explosionX.Left = (range * iCount * WIDTH_) + .xCoord
                        'moves the explosion leftwards
                    Else
                        .explosionX.Width = (.xCoord - .explosionX.Left) + (range * iCount * WIDTH_) + WIDTH_
                        'strecthes the explosionX to the right
                    End If
                    If (fireCollision = WALL_ Or fireCollision = BRICK_ Or fireCollision = BOMB_) And powerCheck Then
                        range = 1000
                        'if the explosion hits a wall or a brick, break the loop
                    End If
                Next range
                For range = 1 To .power
                    fireCollision = Collisions(.ID, STILL_, .xCoord, .yCoord + range * iCount * WIDTH_, DESTROY_)
                    'destroys anything explosionY hits
                    If iCount = -1 Then
                        .explosionY.Top = (range * iCount * WIDTH_) + .yCoord
                        'moves the explosionY upwards
                    Else
                        .explosionY.Height = (.yCoord - .explosionY.Top) + (range * iCount * WIDTH_) + WIDTH_
                        'stretches the explosionY downwards
                    End If
                    If (fireCollision = WALL_ Or fireCollision = BRICK_ Or fireCollision = BOMB_) And powerCheck Then
                        range = 1000
                        'if the explosion hits a wall or a brick, break the loop
                    End If
                Next range
                fireCollision = Collisions(.ID, STILL_, .xCoord, .yCoord, DESTROY_)
                'destroys anything under the bomb
            Next iCount
        End If
        If (0 < .age < bombDEATH_) Then
        'Between deathAge And 0
            For range = .explosionX.Left To (.explosionX.Left + .explosionX.Width - WIDTH_) Step WIDTH_
                Collisions .ID, STILL_, range, .yCoord, DESTROY_
                'for every spot in explosionX, destroy the objects there
            Next range
            For range = .explosionY.Top To (.explosionY.Top + .explosionY.Height - WIDTH_) Step WIDTH_
                Collisions .ID, STILL_, .xCoord, range, DESTROY_
                'for every spot in explosionY, destroy the objects there
            Next range
            If .type = FIREBALL Then .age = 0
        End If
        .direction = STILL_     'Stop moving
        If .age <= 0 Then
            .inUse = False      'Stops the bomb from being inUse
            .xCoord = -1        'moves it so no collisions
            .yCoord = -1        '^^
            Me.Controls.Remove "bombBody" & .ID
            Me.Controls.Remove "explosionX" & .ID
            Me.Controls.Remove "explosionY" & .ID
            'removes the bomb's objects from the form controls
            playerArray(.ownerID).bombs = playerArray(.ownerID).bombs + 1
            'adds a bomb back to the owner's pack
            UpdateStats
        End If
    End With
End Sub

Private Sub bombMovement()
Dim bCount As Integer
Dim speedVar As Integer
    For bCount = 0 To UBound(bombArray)
        If bombArray(bCount).inUse = True Then
        With bombArray(bCount)
            If .moveOk = True Then      'checks if player is ok to move
                .lendirY = Int(-Sin(.direction * 90 * Pi / 180) * WIDTH_)    'uses Sin/Cos to convert direction...
                .lenDirX = Int(Cos(.direction * 90 * Pi / 180) * WIDTH_)     '...to x and y in the directon.
                .lendirY = .lendirY - (.lendirY Mod WIDTH_)    'because pi isnt accurate, we need...
                .lenDirX = .lenDirX - (.lenDirX Mod WIDTH_)    '...to take off the trailing fraction
                If .direction = STILL_ Then
                    .lenDirX = 0     'if lenDir is calculated using STILL_ (-1)...    | cos (-90^) = 0
                    .lendirY = 0     '...it becomes downwards, and we dont want that  | sin (-90^) = -1
                End If
                If Collisions(bCount, .direction, .xCoord, .yCoord) = NOTHING_ Then
                    .xCoord = .xCoord + .lenDirX   'updates the .xCoord and .yCoord variables to
                    .yCoord = .yCoord + .lendirY   'where they should be after a button press
                Else
                    If .type = FIREBALL Then
                        .xCoord = .xCoord + .lenDirX
                        .yCoord = .yCoord + .lendirY
                        .BombBody.Top = .BombBody.Top + .lendirY
                        .BombBody.Left = .BombBody.Left + .lenDirX
                        If .age > bombDEATH_ Then .age = bombDEATH_ + 1
                    End If
                    .direction = STILL_
                    'stops the bomb
                End If
            End If
            If playerArray(.ownerID).skullState(1) = skDragon Then
                speedVar = 2
            Else
                speedVar = 1
            End If
            If .xCoord + (WIDTH_ / 4) <> .BombBody.Left Then
            'if the bomb's xCoord isnt equal to the body's left
                .moveOk = False
                .BombBody.Left = .BombBody.Left + .lenDirX / (10 / speedVar)
                'makes it not ok to move in another direction, and moves the bomb
            End If
            If .yCoord + (WIDTH_ / 4) <> .BombBody.Top Then
            'if the bomb's xCoord isnt equal to the body's left
                .BombBody.Top = .BombBody.Top + .lendirY / (10 / speedVar)
                .moveOk = False
                'makes it not ok to move in another direction, and moves the bomb
            End If
            If .xCoord + (WIDTH_ / 4) = .BombBody.Left And _
            .yCoord + (WIDTH_ / 4) = .BombBody.Top And .moveOk = False Then
            'if the xCoord/yCoord = the body's left/top,
            'having the moveOk in there stops it from continually checking
                .moveOk = True
                'makes it ok to move another direction
            End If
            
        End With
        End If
        
    Next bCount
End Sub

Public Sub addBrick()
    Dim brickCount As Integer
    Dim bCount As Integer
    temp = -1
    'temp is used for testing which bricks are available
    For bCount = 0 To UBound(brickArray)
        If brickArray(bCount).inUse = False Then
            temp = bCount
            bCount = 10000
            'if there is a brick that isnt in use, set that to temp, and break the loop
        End If
    Next bCount
    
    If temp = -1 Then
        temp = UBound(brickArray) + 1
        ReDim Preserve brickArray(temp)
        'if no spare brick was found then set temp to a new brick and make the array larger to fit it
    End If
    
    For bCount = 0 To UBound(brickArray)
        If brickArray(bCount).inUse = True Then brickCount = brickCount + 1
        'find the amount of bricks in use
    Next bCount
    
    If brickCount < ((BOARDWIDTH_ + 1) ^ 2 - (UBound(square) + 1) + 1) / 3 Then
    'if the amount of bricks is less than half the available space
        With brickArray(temp)
            .ID = temp      'sets the ID to temp
            .age = 0        'sets the age to 0
            .inUse = True   '
            Set .BrickBody = Me.Controls.Add("VB.Shape", "brickBody" & .ID)
            'adds the brick to the forms controls
                brickPlace .ID  'place the brick in a spare spot
                With .BrickBody
                    .FillStyle = vbDiagonalCross    'sets the body shape properties
                    .FillColor = ageColorMin_       'for fillcolour,style,border,
                    .BorderStyle = vbTransparent    'width,height and shape,
                    .Width = WIDTH_                 'then makes the shape visible.
                    .Height = WIDTH_                'and brings it to the front
                    .Visible = True
                    .Shape = vbShapeRoundedSquare
                    .ZOrder (0)
                End With
        End With
    End If
End Sub

Public Sub brickPlace(ByVal ID As Integer)
    Dim leftTemp As Integer 'is used to hold a temporary position
    Dim topTemp As Integer  'for the bricks left and top
    With brickArray(ID).BrickBody
    temp = -1
        While temp <> NOTHING_
            'keeps looping while there is a collision
            leftTemp = Int(Rnd * (BOARDWIDTH_ + 1)) * WIDTH_ 'moves the tempVariables to
            topTemp = Int(Rnd * BOARDWIDTH_) * WIDTH_   'a random square on the board
            temp = Collisions(ID, STILL_, leftTemp, topTemp, , -2) 'then checks for collisions
        Wend
        .Left = leftTemp    'finally, when there is no collision
        .Top = topTemp      'the brick is moved to the spare space
    End With
End Sub

Public Sub brickAge()
Dim brCount As Integer  'brick Looping Variable
Dim boCount As Integer  'Bomb Looping Variable
    For brCount = 0 To UBound(brickArray)
        With brickArray(brCount)
            If .inUse = True Then   'checks if the brick is being used on the board
                If .BrickBody.FillColor <> vbBlack Then
                'when the fillcolour is vbBlack, the brick does not have to develop any further
                    .age = .age + 1 'increases the age
                    Select Case Int(.age / ageStep_)
                        Case 2:
                            .BrickBody.FillColor = &HE0E0E0
                            'very light grey
                        Case 3:
                            .BrickBody.FillColor = &HC0C0C0
                            'light grey
                        Case 4:
                            .BrickBody.FillColor = &H808080
                            'medium grey
                        Case 5:
                            .BrickBody.FillColor = &H404040
                            'dark grey
                        Case 6:
                            temp = Collisions(.ID, STILL_, .BrickBody.Left, .BrickBody.Top)
                            'sets temp to any collisions at the bricks spot
                            If Int(Left(temp, 1)) = PLAYER_ Then playerDestroy Mid(temp, 2)
                            'if there is a player, destroy it
                            For boCount = 0 To UBound(bombArray)
                                If bombArray(boCount).inUse = True Then
                                'if the bomb is being used
                                    If .BrickBody.Left = bombArray(boCount).xCoord And .BrickBody.Top = bombArray(boCount).yCoord Then
                                    'if there is a bomb under the brick
                                        If bombArray(boCount).age > bombDEATH_ + 1 Then bombArray(boCount).age = bombDEATH_ + 1
                                        'if bomb hasnt exploded yet... explode it
                                    End If
                                End If
                            Next boCount
                            .BrickBody.FillColor = vbBlack      'sets the format to the final appearence
                            .BrickBody.BorderStyle = 1          'of the brick
                            .BrickBody.BackStyle = BACKOPAQUE_  '
                            .BrickBody.BackColor = vbGrey       'grey background, with black cross fillstyle
                    End Select
                End If
            End If
        End With
    Next brCount
    
    For pCount = 0 To UBound(pickupArray)
        With pickupArray(pCount)
            If .age < 100 Then .age = .age + 1
            'also ages pickups
        End With
    Next pCount
End Sub

Private Sub brickDestroy(ByVal ID As Integer)
    With brickArray(ID)
        .age = 0        'sets the age to 0
        .inUse = False
        makePickup .BrickBody.Left, .BrickBody.Top
        Me.Controls.Remove "brickBody" & ID
        'takes the brick off the board and creates a pickup
    End With
End Sub

Private Sub playerDestroy(ByVal ID As Integer, Optional Death As Boolean = True)
    Dim dCount As Integer   'direction loop variable
    With playerArray(ID)
        If .dieOk = 0 And Death And .lives > 0 Then .lives = .lives - 1
        'checks if it is ok for player to die then lowers the lives by one to zero
        If .lives = 0 Then  'if the player is dead
            '.bombs = -MAXITEMS_     'removes all of the players possible bombs
            .direction(0) = STILL_     'stops the player from movin
            .body.BackStyle = BACKTRANS_
            .body.FillStyle = vbDiagonalCross
            .body.ZOrder 0
            .skullState(0) = 1
        Else 'if the player isnt dead
'            For dCount = RIGHT_ To DOWN_ 'for each direction
'                If Collisions(ID, dCount, .xCoord, .yCoord) = NOTHING_ Then
'                'if there isnt a collision in the direction
'                    .DIRECTION(0) = dCount 'then move in that direction
'                    dCount = 10000      'and break the loop
'                End If
'            Next dCount
            .body.BackColor = PlayerColours(ID)
            .body.BackStyle = BACKOPAQUE_
            .body.FillStyle = 1
        End If

        .dieOk = 50 'stops the player from instantly dying again
    End With
End Sub

Private Sub makePickup(x As Integer, y As Integer)
Dim sparePickup As Integer   'used to check for unused pickups
Dim puCount As Integer     'pickUp looping variable
If GameMode.pickupChance = 1 Then     '50/50 chance that it will be 1
    sparePickup = -1       'will remain the same if
    For puCount = 0 To UBound(pickupArray)
        If pickupArray(puCount).inUse = False Then
        'if an unused pickup is found,
            sparePickup = puCount               'set spare pickup to it and
            pickupArray(puCount).inUse = True   'set it to be inUse
            puCount = 10000                     'then break the Loop
        End If
    Next puCount
    If sparePickup = -1 Then
    'if no pickup was found
        sparePickup = UBound(pickupArray) + 1   'make a new space in
        ReDim Preserve pickupArray(sparePickup) 'the array to hold a pickup
    End If
    With pickupArray(sparePickup)
        .age = 0
        .inUse = True
        .type = Int((Rnd * 100))
        If .type <= 44 Then
            .type = puBOMBUP
        ElseIf .type > 44 And .type < 90 Then
            .type = puFIREUP
        ElseIf .type >= 90 And .type <= 97 Then
            .type = puSkull
        Else
            .type = puLIFEUP
        End If
        Set .PickUpBody = Me.Controls.Add("VB.Image", "pickup" & sparePickup)
        'adds the image to the forms controls
        .PickUpBody.Picture = LoadPicture("Images/" & PickupDir(.type))
        'loads the picture that is to be used
        With .PickUpBody
            .Visible = True   'sets the image properties
            .Left = x         'and position, also brings
            .Top = y          'the image to the front of
            .Width = WIDTH_   'the form
            .Height = WIDTH_  '
            .ZOrder 0         '
        End With
    End With
End If
End Sub

Private Sub pickupDestroy(ID As Integer)
    Me.Controls.Remove ("pickup" & ID)
    pickupArray(ID).inUse = False
    'removes the pickup from the form controls and stops it being inUse
End Sub

Private Sub SkullEffects()
Dim bCount As Integer
    
    For pCount = 0 To UBound(playerArray)
    With playerArray(pCount)
        If .skullState(1) <> 0 And .skullState(0) >= 2 Then
            .skullState(0) = .skullState(0) - 1
            .body.BorderColor = vbRed
            .body.BorderWidth = 2
            Select Case .skullState(1)
                Case skBombs
                    Form_KeyDown .keysDefined(BOMB_), 0
                Case skDirection
                    If .keysDefined(LEFT_) <> PlayerKeys(pCount, RIGHT_) Then
                        .keysDefined(LEFT_) = PlayerKeys(pCount, RIGHT_)
                        .keysDefined(RIGHT_) = PlayerKeys(pCount, LEFT_)
                        .keysDefined(UP_) = PlayerKeys(pCount, DOWN_)
                        .keysDefined(DOWN_) = PlayerKeys(pCount, UP_)
                    End If
                Case skInvisible
                    If .skullState(0) Mod 100 < 10 Then
                        .body.BorderStyle = 1
                        .body.Visible = False
                        .nameTag.Visible = True
                    Else
                        .nameTag.Visible = False
                    End If
                Case skSwap
                    Dim SwapTo As Integer
                    Do
                        temp = Int(Rnd * NumOfPlayers)
                    Loop While temp = pCount And playerArray(temp).lives <= 0
                    SwapTo = .xCoord
                    .body.Left = playerArray(temp).xCoord
                    playerArray(temp).body.Left = SwapTo
                    SwapTo = .yCoord
                    .body.Top = playerArray(temp).yCoord
                    playerArray(temp).body.Top = SwapTo
                    SwapTo = .xCoord
                    .xCoord = playerArray(temp).xCoord
                    playerArray(temp).xCoord = SwapTo
                    SwapTo = .yCoord
                    .yCoord = playerArray(temp).yCoord
                    playerArray(temp).yCoord = SwapTo
                    .skullState(0) = 1
                Case skMagnet
                    For bCount = 0 To UBound(bombArray)
                        If bombArray(bCount).xCoord = .xCoord Then
                            If bombArray(bCount).yCoord > .yCoord Then
                                bombArray(bCount).direction = UP_
                            Else
                                bombArray(bCount).direction = DOWN_
                            End If
                        ElseIf bombArray(bCount).yCoord = .yCoord Then
                            If bombArray(bCount).xCoord < .xCoord Then
                                bombArray(bCount).direction = RIGHT_
                            Else
                                bombArray(bCount).direction = LEFT_
                            End If
                        End If
                    Next bCount
                Case skLife
                    .dieOk = .skullState(0) + 50
                    .body.BorderColor = vbYellow
            End Select
        End If
        If .skullState(0) = 1 Then
            'set everything back to normal
            Select Case .skullState(1)
                Case skDirection
                    .keysDefined(LEFT_) = PlayerKeys(pCount, LEFT_)
                    .keysDefined(RIGHT_) = PlayerKeys(pCount, RIGHT_)
                    .keysDefined(UP_) = PlayerKeys(pCount, UP_)
                    .keysDefined(DOWN_) = PlayerKeys(pCount, DOWN_)
                Case skInvisible
                    .body.BorderStyle = 1
                    .body.Visible = True
                    .nameTag.Visible = True
                Case skSlow
                    .body.Left = .xCoord
                    .body.Top = .yCoord
                Case skFast
                    .body.Left = .xCoord
                    .body.Top = .yCoord
            End Select
            .skullState(1) = 0
            .skullState(0) = 0
            .body.BorderWidth = 1
            .body.BorderColor = vbBlack
        End If
    End With
    Next pCount
End Sub

Private Sub NameTagMovement()
    For pCount = 0 To UBound(playerArray)   'for each player
        With (playerArray(pCount))
            If .nameTag.Top <> .body.Top _
            Or .nameTag.Left <> .body.Left + WIDTH_ Then
                .nameTag.Top = .body.Top
                .nameTag.Left = .body.Left + WIDTH_
            End If
        End With
        'if the tag is not near the body,
        'then move the tag to next to the body
    Next pCount
End Sub

Private Sub WinCheck()
    Dim winnerCheck As Integer
    winnerCheck = 0
    'this loop checks for living players, if there are more than 2 then nothing...
    'happens if there is 1 then the winner is declared and the form changes, if
    'there are no winners, it is declared a tie.
    For pCount = 0 To UBound(playerArray)
        If playerArray(pCount).lives > 0 Then
            winnerCheck = winnerCheck + 1   'adds one for every living player
            winner = pCount                 'the winner, if there is only 1
        End If
    Next pCount
    Select Case winnerCheck
        Case 0
            winner = -1     'no winner
            Load frmWinner    'changes the form to
            frmWinner.Show    'the next one and stops
            Unload Me         'this form from running
            Exit Sub        'ends this subprocedure
            
        Case 1
            If playerArray(winner).nameTag.Caption <> "" Then
                MsgBox playerArray(winner).nameTag.Caption & " Wins"
            Else
                MsgBox "Player " & winner + 1 & " Wins"
            End If
            'if the player has a name then msgbox it, otherwise, msgbox the player number
            Load frmWinner  'changes to the next
            frmWinner.Show  'form and stops this
            Unload Me       'one running continuously
    End Select
End Sub

Private Sub UpdateStats()
    Dim pCount As Byte
    Dim skullName As String, tempName As String
    lblStats.Caption = ""
    For pCount = 0 To UBound(playerArray)
        With playerArray(pCount)
            Select Case .skullState(1)
                Case NOTHING_
                    skullName = "None"
                Case skBombs
                    skullName = "Auto-Bomb"
                Case skShortBomb
                    skullName = "Short Fuse"
                Case skSlow
                    skullName = "Slow-motion"
                Case skDirection
                    skullName = "Reverse Walking"
                Case skInvisible
                    skullName = "Invisible"
                Case skSwap
                    skullName = "Swap"
                Case skMagnet
                    skullName = "Magnetic Bomber"
                Case skFast
                    skullName = "Fast-motion"
                Case skPower
                    skullName = "Power Bomb"
                Case skLife
                    skullName = "Invulnerability"
                Case skDragon
                    skullName = "Fireball/Dragon"
            End Select
            If .strName = "" Then
                tempName = "Player " & pCount + 1
            Else
                tempName = .strName
            End If
            lblStats.Caption = lblStats.Caption & tempName & vbCrLf & "    Power: " & .power _
            & vbCrLf & "    Bombs Left: " & .bombs & vbCrLf & "    Skull Type: " & skullName & vbCrLf _
            & "    Lives: " & .lives & vbCrLf
        End With
    Next pCount
End Sub

Private Sub GhostAbilities()
    Dim gCount As Byte, pCount As Byte
    
    For pCount = 0 To UBound(playerArray)
        playerArray(pCount).attached(0) = -1
    Next pCount
    
    For gCount = 0 To UBound(playerArray)
        With playerArray(gCount)
            For pCount = 0 To UBound(playerArray)
                If .lives = 0 And playerArray(pCount).lives > 0 Then
                    If .xCoord = playerArray(pCount).xCoord And .yCoord = playerArray(pCount).yCoord Then
                        .attached(0) = pCount
                        playerArray(pCount).attached(0) = gCount
                    End If
                End If
            Next pCount
        End With
    Next gCount
End Sub

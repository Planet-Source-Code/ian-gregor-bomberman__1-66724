Attribute VB_Name = "modMisc"
'the use of a module allows global vaiables to be set,
'also it allows global sub procedures and functions to be set,
'
'the main use of globals in this game is passing information between
'forms, like the winner of a battle, and the keys/colours/names defined
'in the setup, also the directory of the pickups is in there...

Global PlayerKeys(3, 4) As Integer  'each of the 4 players
Global NumOfPlayers As Integer      'the number of players playing
Global PlayerColours(3) As Long     'the colour of each player
Global PlayerNames(-1 To 3) As String     'The names of each player
Global PlayerWins(-1 To 3) As Integer
Global TotalGames As Integer
Global winner As Integer            'the player that won
Global PickupDir(3) As String       'the directory of the pickup images
Global Mode As Integer

'the default colour of each player
Const DefColour0 = vbBlue
Const DefColour1 = vbRed
Const DefColour2 = vbGreen
Const DefColour3 = vbYellow

'the default keys for each of the players,
    'first player
Const DefKey0Right_ = vbKeyRight
Const DefKey0Up_ = vbKeyUp
Const DefKey0Left_ = vbKeyLeft
Const DefKey0Down_ = vbKeyDown
Const DefKey0Bomb_ = vbKeyShift
    'second player
Const DefKey1Right_ = vbKeyD
Const DefKey1Up_ = vbKeyW
Const DefKey1Left_ = vbKeyA
Const DefKey1Down_ = vbKeyS
Const DefKey1Bomb_ = vbKeyTab
    'third player
Const DefKey2Right_ = vbKeyH
Const DefKey2Up_ = vbKeyT
Const DefKey2Left_ = vbKeyF
Const DefKey2Down_ = vbKeyG
Const DefKey2Bomb_ = vbKeyR
    'fourth player
Const DefKey3Right_ = vbKeyL
Const DefKey3Up_ = vbKeyI
Const DefKey3Left_ = vbKeyJ
Const DefKey3Down_ = vbKeyK
Const DefKey3Bomb_ = vbKeyU

Public Sub SetPlayerKeys()
Dim KeyCheck As Integer
For i = 0 To 3
    KeyCheck = 0
    For j = 0 To 4
        If PlayerKeys(i, j) = 0 Then KeyCheck = 1
    Next j
    'if any of the keys are undefined then define them as defualt
    If KeyCheck = 1 Then
        Select Case i
            Case 0
                PlayerKeys(0, 0) = DefKey0Right_
                PlayerKeys(0, 1) = DefKey0Up_
                PlayerKeys(0, 2) = DefKey0Left_
                PlayerKeys(0, 3) = DefKey0Down_
                PlayerKeys(0, 4) = DefKey0Bomb_
            Case 1
                PlayerKeys(1, 0) = DefKey1Right_
                PlayerKeys(1, 1) = DefKey1Up_
                PlayerKeys(1, 2) = DefKey1Left_
                PlayerKeys(1, 3) = DefKey1Down_
                PlayerKeys(1, 4) = DefKey1Bomb_
            Case 2
                PlayerKeys(2, 0) = DefKey2Right_
                PlayerKeys(2, 1) = DefKey2Up_
                PlayerKeys(2, 2) = DefKey2Left_
                PlayerKeys(2, 3) = DefKey2Down_
                PlayerKeys(2, 4) = DefKey2Bomb_
            Case 3
                PlayerKeys(3, 0) = DefKey3Right_
                PlayerKeys(3, 1) = DefKey3Up_
                PlayerKeys(3, 2) = DefKey3Left_
                PlayerKeys(3, 3) = DefKey3Down_
                PlayerKeys(3, 4) = DefKey3Bomb_
        End Select
    End If
Next i
End Sub
Public Sub SetPlayerColours()

For i = 0 To 3
'for each player, if the colour is undefined then use the default
    If PlayerColours(i) = 0 Then
        Select Case i
            Case 0
                PlayerColours(0) = DefColour0
            Case 1
                PlayerColours(1) = DefColour1
            Case 2
                PlayerColours(2) = DefColour2
            Case 3
                PlayerColours(3) = DefColour3
        End Select
    End If
Next i
    
End Sub

Public Sub checkNames()
'if the name hasnt changed then remove it so the name tag doesnt say "Name?"
    For i = 0 To 3
        If LCase(PlayerNames(i)) = "name?" Then PlayerNames(i) = ""
    Next i
End Sub




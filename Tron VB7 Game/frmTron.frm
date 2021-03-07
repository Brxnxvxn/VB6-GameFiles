VERSION 5.00
Begin VB.Form frmTron 
   BackColor       =   &H80000012&
   Caption         =   "TronGame"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   831
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraModes 
      Caption         =   "Modes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   7800
      TabIndex        =   2
      Top             =   4200
      Width           =   4455
      Begin VB.OptionButton optMultiplayer 
         Caption         =   "Multiplayer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   4215
      End
      Begin VB.OptionButton optSinglePlayHard 
         Caption         =   "SinglePlayer Hard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   4215
      End
      Begin VB.OptionButton optSinglePlay 
         Caption         =   "SinglePlayer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Timer tmrMoveUser 
      Interval        =   30
      Left            =   7200
      Top             =   7440
   End
   Begin VB.PictureBox picTron 
      BackColor       =   &H80000007&
      Height          =   6000
      Left            =   120
      ScaleHeight     =   396
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   1
      Top             =   1560
      Width           =   7500
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmTron.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   7800
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image imgTron 
      Height          =   1455
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label lblScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7560
      Width           =   7500
   End
End
Attribute VB_Name = "frmTron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Branavan Keethabaskaran
'Date : December 13, 2019
'Purpose : To recreate the game of tron with additional features

Option Explicit

'Declare Global Variables
Dim intX As Integer
Dim intY As Integer
Dim strDirection As String
Dim strDirectionTwo As String
Dim intTimer As Integer
Dim intCounterRed As Integer
Dim intCounterBlue As Integer
Dim intXSecond As Integer
Dim intYSecond As Integer
Dim intRndItemX As Integer
Dim intRndItemY As Integer
Dim intCount As Integer
Dim intCountImmortalRed As Integer
Dim intCountImmortalBlue As Integer

Private Sub Form_Load()
'Print Title
imgTron.Picture = LoadPicture(App.Path & "\Tron.jpg")

'Randomize item droppings
Randomize

'Intialize Global Variables
intX = 40
intY = 40
intXSecond = 450
intYSecond = 350
strDirection = "Right" 'Move User forward when idle
strDirectionTwo = "Left" 'Move Second Player to the left when idle
intCounterRed = 0
intCounterBlue = 0
intTimer = 0
intRndItemX = 0
intRndItemY = 0
intCount = 0
intCountImmortalRed = 0
intCountImmortalBlue = 0
'Keep the game not started until user clicks start game
tmrMoveUser.Enabled = False


End Sub

Private Sub optMultiplayer_Click()
'Clear items on grid from previous round
picTron.Cls

'Place Players on the same place of the board
intX = 40
intY = 40
intXSecond = 450
intYSecond = 350

'Create random items onto grid
For intCount = 0 To 5
    intRndItemX = Int(Rnd * 500) + 1
    intRndItemY = Int(Rnd * 400) + 1
    'Make the items the shape of a square
    picTron.PSet (intRndItemX, intRndItemY), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY - 1), vbGreen
Next intCount
End Sub

Private Sub optSinglePlay_Click()
'Clear items on grid from previous round
picTron.Cls

'Place Players on the same place of the board
intX = 40
intY = 40
intXSecond = 450
intYSecond = 350

'Create random items onto grid
For intCount = 0 To 5
    intRndItemX = Int(Rnd * 500) + 1
    intRndItemY = Int(Rnd * 400) + 1
    'Make the items the shape of a square
    picTron.PSet (intRndItemX, intRndItemY), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY - 1), vbGreen
Next intCount
End Sub

Private Sub optSinglePlayHard_Click()
'Clear items on grid from previous round
picTron.Cls

'Place Players on the same place of the board
intX = 40
intY = 40
intXSecond = 450
intYSecond = 350

'Create random items onto grid
For intCount = 0 To 5
    intRndItemX = Int(Rnd * 500) + 1
    intRndItemY = Int(Rnd * 400) + 1
    'Make the items the shape of a square
    picTron.PSet (intRndItemX, intRndItemY), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY), vbGreen
    picTron.PSet (intRndItemX, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX - 1, intRndItemY - 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY + 1), vbGreen
    picTron.PSet (intRndItemX + 1, intRndItemY - 1), vbGreen
Next intCount

End Sub

Private Sub picTron_KeyPress(KeyAscii As Integer)
'Allow the game to start when any key is pressed
tmrMoveUser.Enabled = True

'Decide which direction user wants to move in
'If user presses w and wants to go up
If KeyAscii = 119 Or KeyAscii = 87 Then
    strDirection = "Up"
    
'If user pressess s and wants to go down
ElseIf KeyAscii = 115 Or KeyAscii = 83 Then
    strDirection = "Down"

'If user presses a and wants to go left
ElseIf KeyAscii = 97 Or KeyAscii = 65 Then
    strDirection = "Left"
    
'If user presses d and wants to go right
ElseIf KeyAscii = 100 Or KeyAscii = 68 Then
    strDirection = "Right"
  
End If

'If User wants to pause the game
If KeyAscii = 112 Or KeyAscii = 80 Then
    tmrMoveUser.Enabled = False
    MsgBox ("Game Paused: Click u to resume game")
   
'If User wants to unpause
ElseIf KeyAscii = 117 Or KeyAscii = 85 Then
    tmrMoveUser.Enabled = True
    
ElseIf KeyAscii = 82 Or KeyAscii = 114 Then
    picTron.Cls
    MsgBox ("Game Reseted")
    tmrMoveUser.Enabled = False
    optSinglePlay = False
    optSinglePlayHard = False
    optMultiplayer = False
End If

'If there are two players
If optMultiplayer = True Then
    'If i is pressed
    If KeyAscii = 105 Or KeyAscii = 73 Then
        strDirectionTwo = "Up"
        
    'If k is pressed
    ElseIf KeyAscii = 107 Or KeyAscii = 75 Then
        strDirectionTwo = "Down"
    
    'If j is pressed
    ElseIf KeyAscii = 106 Or KeyAscii = 74 Then
        strDirectionTwo = "Left"
    
    'if l is Pressed
    ElseIf KeyAscii = 108 Or KeyAscii = 76 Then
        strDirectionTwo = "Right"
   End If
End If

End Sub


Private Sub tmrMoveUser_Timer()
'Subtract values of invincibilility by 1
If intCountImmortalRed <> 0 Then
    intCountImmortalRed = intCountImmortalRed - 1

ElseIf intCountImmortalBlue <> 0 Then
    intCountImmortalBlue = intCountImmortalBlue - 1
End If

'Track Score of player Red
intCounterRed = intCounterRed + 1

'Move User
If strDirection = "Up" Then
    intY = intY - 2
ElseIf strDirection = "Down" Then
    intY = intY + 2
ElseIf strDirection = "Left" Then
    intX = intX - 2
ElseIf strDirection = "Right" Then
    intX = intX + 2
End If
           
    
'If user decides to play single player hard mode
If optSinglePlayHard = True Then
    If intCounterRed Mod 500 = 0 Then
        tmrMoveUser.Interval = tmrMoveUser.Interval - 3
        If tmrMoveUser.Interval = 1 Then
            tmrMoveUser.Interval = 1
        End If
    End If
End If

'If user selects multi player option
If optMultiplayer = True Then
    intCounterBlue = intCounterBlue + 1
    If strDirectionTwo = "Up" Then
        intYSecond = intYSecond - 2
    
    ElseIf strDirectionTwo = "Down" Then
        intYSecond = intYSecond + 2
    
    ElseIf strDirectionTwo = "Right" Then
        intXSecond = intXSecond + 2
    
    ElseIf strDirectionTwo = "Left" Then
        intXSecond = intXSecond - 2
    
    End If
    'Detect when Red crashes into trails of themselves
    If CStr(picTron.Point(intX, intY)) = vbRed And intCountImmortalRed = 0 Then
        picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into itself. Blue Wins! But Red wins by points with a score of: " & intCounterRed)
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into itself. Blue Wins!")
        End If
        
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    
    'Detect when red crashed into trails of blue
    ElseIf CStr(picTron.Point(intX, intY)) = vbBlue And intCountImmortalRed = 0 Then
        picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into Blue. Blue Wins! But Red wins by points with a score of: " & intCounterRed)
            
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into Blue. Blue Wins!")
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    
    'Detect when blue crashed into Red
    ElseIf CStr(picTron.Point(intXSecond, intYSecond)) = vbRed And intCountImmortalBlue = 0 Then
        picTron.Circle (intXSecond, intYSecond), 5, vbGreen
        If intCounterRed >= intCounterBlue Then
            MsgBox ("Blue Crashed into Red. Red Wins!")
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Blue Crashed into Red. Red Wins! But Blue wins by points with a score of: " & intCounterBlue)
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    
    'Detect when blue crashes into itself
    ElseIf CStr(picTron.Point(intXSecond, intYSecond)) = vbBlue And intCountImmortalBlue = 0 Then
        picTron.Circle (intXSecond, intYSecond), 5, vbGreen
        If intCounterRed >= intCounterBlue Then
            MsgBox ("Blue Crashed into itself. Red Wins!")
        
        ElseIf intCounterBlue > intCounterRed Then
            MsgBox ("Blue Crashed into itself. Red Wins! But Blue wins with points with a score of: " & intCounterBlue)
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
        
    
    'Detect when Red hits border
    If intX = 0 Or intX = 500 Or intY = 0 Or intY = 400 Then
        picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into the border! Blue Wins! But red wins with points with a score of: " & intCounterRed)
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into the border! Blue Wins!")
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
        
    'Detect when Blue hits the border
    ElseIf intXSecond = 0 Or intXSecond = 500 Or intYSecond = 0 Or intYSecond = 400 Then
        picTron.Circle (intXSecond, intYSecond), 5, vbGreen
        If intCounterRed >= intCounterBlue Then
            MsgBox ("Blue Crashed into the border! Red Wins!")
        
        ElseIf intCounterBlue > intCounterRed Then
            MsgBox ("Blue Crashed into the border! Red Wins! But blue wins by points with a score of: " & intCounterBlue)
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
    
    'When Blue crashes into Yellow
If intCountImmortalBlue = 0 Then
    If CStr(picTron.Point(intXSecond, intYSecond)) = vbYellow Then
                picTron.Circle (intXSecond, intYSecond), 5, vbGreen
        If intCounterRed >= intCounterBlue Then
            MsgBox ("Blue Crashed into Red. Red Wins!")
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Blue Crashed into Red. Red Wins! But Blue wins by points with a score of: " & intCounterBlue)
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
End If

    'When Red crashes into Magenta
If intCountImmortalRed = 0 Then
    If CStr(picTron.Point(intX, intY)) = vbMagenta Then
                picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into Blue. Blue Wins! But Red wins by points with a score of: " & intCounterRed)
            
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into Blue. Blue Wins!")
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
 End If
 
'Make sure that red cannot go through yellow trail after invincibility
If intCountImmortalRed = 0 Then
    If CStr(picTron.Point(intX, intY)) = vbYellow Then
        picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into itself. Blue Wins! But Red wins by points with a score of: " & intCounterRed)
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into itself. Blue Wins!")
        End If
        
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
 End If
 
 'Make sure that blue cannot go through magenta trail after invincibility
 If intCountImmortalBlue = 0 Then
    If CStr(picTron.Point(intXSecond, intYSecond)) = vbMagenta Then
           picTron.Circle (intXSecond, intYSecond), 5, vbGreen
        If intCounterRed >= intCounterBlue Then
            MsgBox ("Blue Crashed into itself. Red Wins!")
        
        ElseIf intCounterBlue > intCounterRed Then
            MsgBox ("Blue Crashed into itself. Red Wins! But Blue wins with points with a score of: " & intCounterBlue)
        End If
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
End If

    'Display score for multiplayer
    lblScore.Caption = "PlayerOne Score: " & intCounterRed & "    " & "PlayerTwo Score:" & intCounterBlue
End If
    
'Display Score for single player
If optSinglePlay = True Or optSinglePlayHard = True Then
    lblScore.Caption = "Your Score = " & intCounterRed
End If

'For Single Player
'Detect when Red crashes into trails of themselves
If CStr(picTron.Point(intX, intY)) = vbRed And intCountImmortalRed = 0 Then
    picTron.Circle (intX, intY), 5, vbGreen
    MsgBox ("You Crashed! Your score is: " & intCounterRed)
    picTron.Cls
    tmrMoveUser.Enabled = False
    'Reset Position of Users
    intX = 40
    intY = 40
    intXSecond = 450
    intYSecond = 350
    'Make user pick the game type after end of game
    optSinglePlay = False
    optSinglePlayHard = False
    optMultiplayer = False
    'Reset Scores
    intCounterRed = 0
    'Make sure invincibility power does not transfer to next game
    intCountImmortalBlue = 0
    intCountImmortalRed = 0

'Detect when Red hits border
ElseIf intX = 0 Or intX = 500 Or intY = 0 Or intY = 400 Then
    picTron.Circle (intX, intY), 5, vbGreen
    MsgBox ("You Hit the Border! Your Score is: " & intCounterRed)
    picTron.Cls
    tmrMoveUser.Enabled = False
    'Reset Position of Users
    intX = 40
    intY = 40
    intXSecond = 450
    intYSecond = 350
    'Make user pick the game type after end of game
    optSinglePlay = False
    optSinglePlayHard = False
    optMultiplayer = False
    'Reset Scores
    intCounterRed = 0
    'Make sure invincibility power does not transfer to next game
    intCountImmortalBlue = 0
    intCountImmortalRed = 0
End If

'Make sure that red cannot go through yellow trail after invincibility
If intCountImmortalRed = 0 Then
    If CStr(picTron.Point(intX, intY)) = vbYellow Then
        picTron.Circle (intX, intY), 5, vbGreen
        If intCounterRed > intCounterBlue Then
            MsgBox ("Red Crashed into itself. Blue Wins! But Red wins by points with a score of: " & intCounterRed)
        
        ElseIf intCounterBlue >= intCounterRed Then
            MsgBox ("Red Crashed into itself. Blue Wins!")
        End If
        
        picTron.Cls
        tmrMoveUser.Enabled = False
        'Reset Position of Users
        intX = 40
        intY = 40
        intXSecond = 450
        intYSecond = 350
        'Make user pick the game type after end of game
        optSinglePlay = False
        optSinglePlayHard = False
        optMultiplayer = False
        'Reset Scores
        intCounterRed = 0
        intCounterBlue = 0
        'Make sure invincibility power does not transfer to next game
        intCountImmortalBlue = 0
        intCountImmortalRed = 0
    End If
 End If

'Detect when either blue or red hits a green item
If CStr(picTron.Point(intX, intY)) = vbGreen Then
    intCounterRed = intCounterRed + 1000
    intCountImmortalRed = 200
End If

If CStr(picTron.Point(intXSecond, intYSecond)) = vbGreen Then
    intCounterBlue = intCounterBlue + 1000
    intCountImmortalBlue = 200
End If


'Display users
picTron.PSet (intX, intY), vbRed
'Make the trial of red a rectangle shape
'When user is moving left or right
If strDirection = "Left" Or strDirection = "Right" Then
    picTron.PSet (intX - 1, intY), vbRed
    picTron.PSet (intX - 1, intY - 1), vbRed
    picTron.PSet (intX - 1, intY + 1), vbRed
    picTron.PSet (intX - 1, intY + 2), vbRed
    picTron.PSet (intX - 1, intY - 2), vbRed
    picTron.PSet (intX, intY - 1), vbRed
    picTron.PSet (intX, intY + 1), vbRed

'When user is moving up or down
ElseIf strDirection = "Up" Or strDirection = "Down" Then
    picTron.PSet (intX, intY - 1), vbRed
    picTron.PSet (intX - 1, intY - 1), vbRed
    picTron.PSet (intX + 1, intY - 1), vbRed
    picTron.PSet (intX - 2, intY - 1), vbRed
    picTron.PSet (intX + 2, intY - 1), vbRed
    picTron.PSet (intX + 1, intY), vbRed
    picTron.PSet (intX - 1, intY), vbRed
End If

If optMultiplayer = True Then
    picTron.PSet (intXSecond, intYSecond), vbBlue
    
    'Make blue trail more like a rectangle
    'If blue moves up or down
    If strDirectionTwo = "Up" Or strDirectionTwo = "Down" Then
        picTron.PSet (intXSecond, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond + 1, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond - 2, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond + 2, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond + 1, intYSecond), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond), vbBlue
    
    'IF blue moves right or left
    ElseIf strDirectionTwo = "Right" Or strDirectionTwo = "Left" Then
        picTron.PSet (intXSecond - 1, intYSecond), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond + 1), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond - 2), vbBlue
        picTron.PSet (intXSecond - 1, intYSecond + 2), vbBlue
        picTron.PSet (intXSecond, intYSecond - 1), vbBlue
        picTron.PSet (intXSecond, intYSecond + 1), vbBlue
    End If
End If

'Indicate if user has invincibility as a power
If intCountImmortalRed <> 0 Then
    picTron.PSet (intX, intY), vbYellow
    'Make the trial of yellow a rectangle shape
    'When user is moving left or right
    If strDirection = "Left" Or strDirection = "Right" Then
        picTron.PSet (intX - 1, intY), vbYellow
        picTron.PSet (intX - 1, intY - 1), vbYellow
        picTron.PSet (intX - 1, intY + 1), vbYellow
        picTron.PSet (intX - 1, intY + 2), vbYellow
        picTron.PSet (intX - 1, intY - 2), vbYellow
        picTron.PSet (intX, intY - 1), vbYellow
        picTron.PSet (intX, intY + 1), vbYellow
    
    'When user is moving up or down
    ElseIf strDirection = "Up" Or strDirection = "Down" Then
        picTron.PSet (intX, intY - 1), vbYellow
        picTron.PSet (intX - 1, intY - 1), vbYellow
        picTron.PSet (intX + 1, intY - 1), vbYellow
        picTron.PSet (intX - 2, intY - 1), vbYellow
        picTron.PSet (intX + 2, intY - 1), vbYellow
        picTron.PSet (intX + 1, intY), vbYellow
        picTron.PSet (intX - 1, intY), vbYellow
        
    End If
End If

If intCountImmortalBlue <> 0 Then
    picTron.PSet (intXSecond, intYSecond), vbMagenta
    'Make magenta trail more like a rectangle
    'If blue moves up or down
    If strDirectionTwo = "Up" Or strDirectionTwo = "Down" Then
        picTron.PSet (intXSecond, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond + 1, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond - 2, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond + 2, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond + 1, intYSecond), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond), vbMagenta
    
    'IF blue moves right or left
    ElseIf strDirectionTwo = "Right" Or strDirectionTwo = "Left" Then
        picTron.PSet (intXSecond - 1, intYSecond), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond + 1), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond - 2), vbMagenta
        picTron.PSet (intXSecond - 1, intYSecond + 2), vbMagenta
        picTron.PSet (intXSecond, intYSecond - 1), vbMagenta
        picTron.PSet (intXSecond, intYSecond + 1), vbMagenta
    End If
End If


End Sub

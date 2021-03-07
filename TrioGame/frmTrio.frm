VERSION 5.00
Begin VB.Form frmTrio 
   Caption         =   "Trio"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdComputer 
      Caption         =   "Computer's Turn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   7455
   End
   Begin VB.Frame fraWhoGoesFirst 
      Caption         =   "Who Goes First?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      TabIndex        =   3
      Top             =   5520
      Width           =   2415
      Begin VB.OptionButton optComputerFirst 
         Caption         =   "ComputerFirst?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton optUserFirst 
         Caption         =   "UserFirst?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6000
      TabIndex        =   2
      Top             =   5520
      Width           =   1575
   End
   Begin VB.PictureBox picTrio 
      Height          =   4500
      Left            =   120
      ScaleHeight     =   296
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   496
      TabIndex        =   1
      Top             =   840
      Width           =   7500
      Begin VB.Image imgCompToken3 
         Height          =   1350
         Left            =   6075
         Stretch         =   -1  'True
         Top             =   3075
         Width           =   1350
      End
      Begin VB.Image imgUserToken3 
         DragMode        =   1  'Automatic
         Height          =   1350
         Left            =   4575
         Stretch         =   -1  'True
         Top             =   3075
         Width           =   1350
      End
      Begin VB.Image imgCompToken2 
         Height          =   1350
         Left            =   6075
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Image imgUserToken2 
         DragMode        =   1  'Automatic
         Height          =   1350
         Left            =   4575
         Stretch         =   -1  'True
         Top             =   1575
         Width           =   1350
      End
      Begin VB.Image imgCompToken1 
         Height          =   1350
         Left            =   6075
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1350
      End
      Begin VB.Image imgUserToken1 
         DragMode        =   1  'Automatic
         Height          =   1350
         Left            =   4575
         Stretch         =   -1  'True
         Top             =   75
         Width           =   1350
      End
      Begin VB.Line Line6 
         X1              =   300
         X2              =   300
         Y1              =   0
         Y2              =   400
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   300
         Y1              =   200
         Y2              =   200
      End
      Begin VB.Line Line1 
         X1              =   100
         X2              =   100
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line2 
         X1              =   200
         X2              =   200
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   300
         Y1              =   100
         Y2              =   100
      End
   End
   Begin VB.Label lblComputerScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label lblCompScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label lblUserScore 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Line Line4 
      X1              =   8
      X2              =   304
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label lblTrioTitle 
      Alignment       =   2  'Center
      Caption         =   "Trio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmTrio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Branavan Keethabaskaran
'Date : Thursday, January 9, 2020
'Purpose : To create the game Trio

Option Explicit

'Declare Global Variables
Dim intUser1 As Integer
Dim intUser2 As Integer
Dim intUser3 As Integer
Dim intComp1 As Integer
Dim intComp2 As Integer
Dim intComp3 As Integer
Dim intUserScore As Integer
Dim intComputerScore As Integer

Private Sub cmdComputer_Click()
MoveComputer

'Disable computers turn
cmdComputer.Enabled = False

'Enable drag and drop mode for user
imgUserToken1.DragMode = 1
imgUserToken2.DragMode = 1
imgUserToken3.DragMode = 1
End Sub

Private Sub cmdReset_Click()
'Reset User and Computer Values
intUser1 = -100
intUser2 = -100
intUser3 = -100
intComp1 = -100
intComp2 = -100
intComp3 = -100

'Place markers back in storage area
imgUserToken1.Top = 5
imgUserToken1.Left = 305
imgUserToken2.Top = 105
imgUserToken2.Left = 305
imgUserToken3.Top = 205
imgUserToken3.Left = 305
imgCompToken1.Top = 5
imgCompToken1.Left = 405
imgCompToken2.Top = 105
imgCompToken2.Left = 405
imgCompToken3.Top = 205
imgCompToken3.Left = 405

'Reset scores of user and computer
intUserScore = 0
intComputerScore = 0

'Display Scores of user and computer
lblUserScore.Caption = "User's wins: " & intUserScore
lblComputerScore.Caption = "Computer's wins:  " & intComputerScore

End Sub

Private Sub Form_Load()
'To Randomize computers marker positions
Randomize
'Print tokens on image boxes
imgUserToken1.Picture = LoadPicture(App.Path & "\Red Circle.jpeg")
imgUserToken2.Picture = LoadPicture(App.Path & "\Red Circle.jpeg")
imgUserToken3.Picture = LoadPicture(App.Path & "\Red Circle.jpeg")
imgCompToken1.Picture = LoadPicture(App.Path & "\Blue Circle.jpg")
imgCompToken2.Picture = LoadPicture(App.Path & "\Blue Circle.jpg")
imgCompToken3.Picture = LoadPicture(App.Path & "\Blue Circle.jpg")

'Intialize Global Variables
intUser1 = -100
intUser2 = -100
intUser3 = -100
intComp1 = -100
intComp2 = -100
intComp3 = -100
intUserScore = 0
intComputerScore = 0


'Display Scores of user and computer
lblUserScore.Caption = "User's wins: " & intUserScore
lblComputerScore.Caption = "Computer's wins:  " & intComputerScore

'Disable drag drop and computer
imgUserToken1.DragMode = 0
imgUserToken2.DragMode = 0
imgUserToken3.DragMode = 0
cmdComputer.Enabled = False

End Sub

Public Function ConvertXY(sglX As Single, sglY As Single) As Integer
'This Function is used to convert the area where the user wants to move from coordinates to an single integer that represents a square
'8 1 6
'3 5 7
'4 9 2

'For all the squares in the first row
If sglX \ 100 = 0 Then 'Token is in the first column
    If sglY \ 100 = 0 Then 'Token is in the first row
        ConvertXY = 8
    End If

ElseIf sglX \ 100 = 1 Then 'Token is in the second coloumn
    If sglY \ 100 = 0 Then 'Token is in first row
        ConvertXY = 1
    End If

ElseIf sglX \ 100 = 2 Then 'Token is in the third column
    If sglY \ 100 = 0 Then 'Token is in the first row
        ConvertXY = 6
    End If
End If
   
'For all the squares in the second row
If sglX \ 100 = 0 Then 'Token is in the first column
    If sglY \ 100 = 1 Then 'Token is in the first row
        ConvertXY = 3
    End If

ElseIf sglX \ 100 = 1 Then 'Token is in the second coloumn
    If sglY \ 100 = 1 Then 'Token is in first row
        ConvertXY = 5
    End If

ElseIf sglX \ 100 = 2 Then 'Token is in the third column
    If sglY \ 100 = 1 Then 'Token is in the first row
        ConvertXY = 7
    End If
End If

'For all the squares in the third row
If sglX \ 100 = 0 Then 'Token is in the first column
    If sglY \ 100 = 2 Then 'Token is in the first row
        ConvertXY = 4
    End If

ElseIf sglX \ 100 = 1 Then 'Token is in the second coloumn
    If sglY \ 100 = 2 Then 'Token is in first row
        ConvertXY = 9
    End If

ElseIf sglX \ 100 = 2 Then 'Token is in the third column
    If sglY \ 100 = 2 Then 'Token is in the first row
        ConvertXY = 2
    End If
End If
End Function


Public Sub putImage(token As Control, intSquare As Integer)
'This sub allows to user to put a token in a slot even if they are not exactly over the slot
'8 1 6
'3 5 7
'4 9 2

If intSquare = 8 Then 'first row, first coloumn
    token.Left = 5
    token.Top = 5

ElseIf intSquare = 1 Then 'first row, second coloumn
    token.Left = 105
    token.Top = 5

ElseIf intSquare = 6 Then ' first row, third coloumn
    token.Left = 205
    token.Top = 5

ElseIf intSquare = 3 Then 'second row, first coloumn
    token.Left = 5
    token.Top = 105

ElseIf intSquare = 5 Then 'second row, second column
    token.Left = 105
    token.Top = 105

ElseIf intSquare = 7 Then 'second row, third column
    token.Left = 205
    token.Top = 105

ElseIf intSquare = 4 Then 'third row, first column
    token.Left = 5
    token.Top = 205

ElseIf intSquare = 9 Then 'third row, second column
    token.Left = 105
    token.Top = 205

ElseIf intSquare = 2 Then 'third row, third column
    token.Left = 205
    token.Top = 205
End If
    
End Sub

Public Function IsSquareAvailable(intSquare As Integer) As Boolean
'This function is used to check if a square the user wants to go into is available

IsSquareAvailable = True 'Make this function's default as true

'Check if square is being used by another token
If intSquare = intUser1 Or intSquare = intUser2 Or intSquare = intUser3 Or _
    intSquare = intComp1 Or intSquare = intComp2 Or intSquare = intComp3 Then
    IsSquareAvailable = False
End If

End Function

Public Function IsGameOver() As Boolean
'This function is used to check if any player has won the game

IsGameOver = False 'Sets functions default value to false

If intUser1 + intUser2 + intUser3 = 15 Then
    IsGameOver = True
    MsgBox ("Game Over! User Wins!") 'This is used to indicate that the user has won.
    intUserScore = intUserScore + 1 'increase score of user for winning
    'reset first pick buttons
    optUserFirst = False
    optComputerFirst = False
    'Display wins of user and computer
    lblUserScore.Caption = "User's wins: " & intUserScore
    lblComputerScore.Caption = "Computer's wins:  " & intComputerScore
    'Reset User and Computer Values
    intUser1 = -100
    intUser2 = -100
    intUser3 = -100
    intComp1 = -100
    intComp2 = -100
    intComp3 = -100
    
    'Place markers back in storage area
    imgUserToken1.Top = 5
    imgUserToken1.Left = 305
    imgUserToken2.Top = 105
    imgUserToken2.Left = 305
    imgUserToken3.Top = 205
    imgUserToken3.Left = 305
    imgCompToken1.Top = 5
    imgCompToken1.Left = 405
    imgCompToken2.Top = 105
    imgCompToken2.Left = 405
    imgCompToken3.Top = 205
    imgCompToken3.Left = 405

ElseIf intComp1 + intComp2 + intComp3 = 15 Then
    IsGameOver = True
    MsgBox ("Game Over! Computer Wins!") 'This is used to indicate that the computer has one
    intComputerScore = intComputerScore + 1 'increase score of computer for winning.
    'reset first pick buttons
    optUserFirst = False
    optComputerFirst = False
    'Display wins of user and computer
    lblUserScore.Caption = "User's wins: " & intUserScore
    lblComputerScore.Caption = "Computer's wins:  " & intComputerScore
    'Reset User and Computer Values
    intUser1 = -100
    intUser2 = -100
    intUser3 = -100
    intComp1 = -100
    intComp2 = -100
    intComp3 = -100
    
    'Place markers back in storage area
    imgUserToken1.Top = 5
    imgUserToken1.Left = 305
    imgUserToken2.Top = 105
    imgUserToken2.Left = 305
    imgUserToken3.Top = 205
    imgUserToken3.Left = 305
    imgCompToken1.Top = 5
    imgCompToken1.Left = 405
    imgCompToken2.Top = 105
    imgCompToken2.Left = 405
    imgCompToken3.Top = 205
    imgCompToken3.Left = 405
End If

End Function

Private Sub optComputerFirst_Click()
'Disable drag and drop mode for user
imgUserToken1.DragMode = 0
imgUserToken2.DragMode = 0
imgUserToken3.DragMode = 0

'Enable computer's turn command button
cmdComputer.Enabled = True
End Sub

Private Sub optUserFirst_Click()
'Disable computer's turn
cmdComputer.Enabled = False

'Enable user drag and drop
imgUserToken1.DragMode = 1
imgUserToken2.DragMode = 1
imgUserToken3.DragMode = 1
End Sub

Private Sub picTrio_DragDrop(Source As Control, X As Single, Y As Single)
Dim intSquare As Integer 'represents the number of each square
Dim blnGameOver As Boolean 'represents the value of function IsGameOver

'This will convert the CovertXY into the intSquare
intSquare = ConvertXY(X, Y)

'Check If Square is available
If IsSquareAvailable(intSquare) = True Then
    putImage Source, intSquare
    If Source = imgUserToken1 Then 'if user drags first token
        intUser1 = intSquare
        
    ElseIf Source = imgUserToken2 Then 'if user drags second token
        intUser2 = intSquare
        
    ElseIf Source = imgUserToken3 Then ''if user drags third token
        intUser3 = intSquare
        
    End If
    
    'check if game is over
    blnGameOver = IsGameOver()
    
    'Enable computer's turn
    cmdComputer.Enabled = True

End If

If cmdComputer.Enabled = True Then
    imgUserToken1.DragMode = 0
    imgUserToken2.DragMode = 0
    imgUserToken3.DragMode = 0
End If
End Sub

Public Function FindRandomSquare() As Integer
'Function is used to generate a random square for computer to move
FindRandomSquare = Int(Rnd * 9) + 1

'Generate Random num until Square is available
Do While IsSquareAvailable(FindRandomSquare) = False
    FindRandomSquare = Int(Rnd * 9) + 1
Loop

End Function

Public Function getMarker() As Integer
'Function is used to make user choose which token to move

If intComp1 < 0 Then
    getMarker = 1

ElseIf intComp2 < 0 Then
    getMarker = 2

ElseIf intComp3 < 0 Then
    getMarker = 3

'8 1 6
'3 5 7
'4 9 2

'Choose certain marker for offense
'For First Row
ElseIf (intComp1 = 8 Or intComp1 = 1 Or intComp1 = 6) And (intComp2 = 8 Or intComp2 = 1 Or intComp2 = 6) Then
    getMarker = 3

ElseIf (intComp3 = 8 Or intComp3 = 1 Or intComp3 = 6) And (intComp2 = 8 Or intComp2 = 1 Or intComp2 = 6) Then
    getMarker = 1

ElseIf (intComp1 = 8 Or intComp1 = 1 Or intComp1 = 6) And (intComp3 = 8 Or intComp3 = 1 Or intComp3 = 6) Then
    getMarker = 2
    
'For Second Row
ElseIf (intComp1 = 3 Or intComp1 = 5 Or intComp1 = 7) And (intComp2 = 3 Or intComp2 = 5 Or intComp2 = 7) Then
    getMarker = 3

ElseIf (intComp3 = 3 Or intComp3 = 5 Or intComp3 = 7) And (intComp2 = 3 Or intComp2 = 5 Or intComp2 = 7) Then
    getMarker = 1

ElseIf (intComp1 = 3 Or intComp1 = 5 Or intComp1 = 7) And (intComp3 = 3 Or intComp3 = 5 Or intComp3 = 7) Then
    getMarker = 2

'For Third Row
ElseIf (intComp1 = 4 Or intComp1 = 9 Or intComp1 = 2) And (intComp2 = 4 Or intComp2 = 9 Or intComp2 = 2) Then
    getMarker = 3

ElseIf (intComp3 = 4 Or intComp3 = 9 Or intComp3 = 2) And (intComp2 = 4 Or intComp2 = 9 Or intComp2 = 2) Then
    getMarker = 1

ElseIf (intComp1 = 4 Or intComp1 = 9 Or intComp1 = 2) And (intComp3 = 4 Or intComp3 = 9 Or intComp3 = 2) Then
    getMarker = 2

'For First Column
ElseIf (intComp1 = 8 Or intComp1 = 3 Or intComp1 = 4) And (intComp2 = 8 Or intComp2 = 3 Or intComp2 = 4) Then
    getMarker = 3

ElseIf (intComp3 = 8 Or intComp3 = 3 Or intComp3 = 4) And (intComp2 = 8 Or intComp2 = 3 Or intComp2 = 4) Then
    getMarker = 1

ElseIf (intComp1 = 8 Or intComp1 = 3 Or intComp1 = 4) And (intComp3 = 8 Or intComp3 = 3 Or intComp3 = 4) Then
    getMarker = 2
    
'For Second Column
ElseIf (intComp1 = 1 Or intComp1 = 5 Or intComp1 = 9) And (intComp2 = 1 Or intComp2 = 5 Or intComp2 = 9) Then
    getMarker = 3

ElseIf (intComp3 = 1 Or intComp3 = 5 Or intComp3 = 9) And (intComp2 = 1 Or intComp2 = 5 Or intComp2 = 9) Then
    getMarker = 1

ElseIf (intComp1 = 1 Or intComp1 = 5 Or intComp1 = 9) And (intComp3 = 1 Or intComp3 = 5 Or intComp3 = 9) Then
    getMarker = 2
    
'For Third Column
ElseIf (intComp1 = 6 Or intComp1 = 7 Or intComp1 = 2) And (intComp2 = 6 Or intComp2 = 7 Or intComp2 = 2) Then
    getMarker = 3

ElseIf (intComp3 = 6 Or intComp3 = 7 Or intComp3 = 2) And (intComp2 = 6 Or intComp2 = 7 Or intComp2 = 2) Then
    getMarker = 1

ElseIf (intComp1 = 6 Or intComp1 = 7 Or intComp1 = 2) And (intComp3 = 6 Or intComp3 = 7 Or intComp3 = 2) Then
    getMarker = 2

'For Left Diagonal
ElseIf (intComp1 = 8 Or intComp1 = 5 Or intComp1 = 2) And (intComp2 = 8 Or intComp2 = 5 Or intComp2 = 2) Then
    getMarker = 3

ElseIf (intComp3 = 8 Or intComp3 = 5 Or intComp3 = 2) And (intComp2 = 8 Or intComp2 = 5 Or intComp2 = 2) Then
    getMarker = 1

ElseIf (intComp1 = 8 Or intComp1 = 5 Or intComp1 = 2) And (intComp3 = 8 Or intComp3 = 5 Or intComp3 = 2) Then
    getMarker = 2
   
'For Right Diagonal
ElseIf (intComp1 = 6 Or intComp1 = 5 Or intComp1 = 4) And (intComp2 = 6 Or intComp2 = 5 Or intComp2 = 4) Then
    getMarker = 3

ElseIf (intComp3 = 6 Or intComp3 = 5 Or intComp3 = 4) And (intComp2 = 6 Or intComp2 = 5 Or intComp2 = 4) Then
    getMarker = 1

ElseIf (intComp1 = 6 Or intComp1 = 5 Or intComp1 = 4) And (intComp3 = 6 Or intComp3 = 5 Or intComp3 = 4) Then
    getMarker = 2

        
'Choose certain marker for defense
ElseIf intUser1 + intUser2 + intComp1 = 15 Or intUser2 + intUser3 + intComp1 = 15 Or intUser1 + intUser3 + intComp1 = 15 Then
    getMarker = Int(Rnd * 3) + 2
    
ElseIf intUser1 + intUser2 + intComp2 = 15 Or intUser2 + intUser3 + intComp2 = 15 Or intUser1 + intUser3 + intComp2 = 15 Then
    getMarker = 1
    
ElseIf intUser1 + intUser2 + intComp3 = 15 Or intUser2 + intUser3 + intComp3 = 15 Or intUser1 + intUser3 + intComp3 = 15 Then
    getMarker = Int(Rnd * 2) + 1

End If
    
'Choose marker when there are two markers already blocking
'checks if marker one and marker two are occupied blocking and will use marker three
If (intUser1 + intUser2 + intComp1 = 15 Or intUser2 + intUser3 + intComp1 = 15 Or intUser1 + intUser3 + intComp1 = 15) _
    And (intUser1 + intUser2 + intComp2 = 15 Or intUser2 + intUser3 + intComp2 = 15 Or intUser1 + intUser3 + intComp2 = 15) Then
    getMarker = 3

'checks if marker two and marker three are occupied blocking and will use marker one
ElseIf (intUser1 + intUser2 + intComp2 = 15 Or intUser2 + intUser3 + intComp2 = 15 Or intUser1 + intUser3 + intComp2 = 15) _
    And (intUser1 + intUser2 + intComp3 = 15 Or intUser2 + intUser3 + intComp3 = 15 Or intUser1 + intUser3 + intComp3 = 15) Then
    getMarker = 1

'checks if marker one and marker three occupied and will use marker two
ElseIf (intUser1 + intUser2 + intComp1 = 15 Or intUser2 + intUser3 + intComp1 = 15 Or intUser1 + intUser3 + intComp1 = 15) _
        And (intUser1 + intUser2 + intComp3 = 15 Or intUser2 + intUser3 + intComp3 = 15 Or intUser1 + intUser3 + intComp3 = 15) Then
        
        getMarker = 2
End If





End Function

Public Sub doMove(intCompSquare As Integer, intMark As Integer)
'This sub is used to move the marker of the computer given the marker and square

'This will determine which marker to use and will place it in a square
If intMark = 1 Then
    putImage imgCompToken1, intCompSquare
    intComp1 = intCompSquare

ElseIf intMark = 2 Then
    putImage imgCompToken2, intCompSquare
    intComp2 = intCompSquare
    
ElseIf intMark = 3 Then
    putImage imgCompToken3, intCompSquare
    intComp3 = intCompSquare
End If
End Sub

Public Sub MoveComputer()
'This sub is what combines all of the computer's subs and functions and actually moves the computer
Dim intRandomSquare As Integer 'represents the function FindRandomSquare
Dim intMarker As Integer 'represents the function getMarker
Dim blnCompGameOver As Integer


'This will get a marker for the computer to move
getMarker
intMarker = getMarker

'8 1 6
'3 5 7
'4 9 2

'Check if game is over
blnCompGameOver = IsGameOver()

'Print out Computer wins
If blnCompGameOver = True Then
  
'Make Computer defensive and offensive rather than random
'For Offense
'For First Row First
ElseIf (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And (intComp1 = 1 Or intComp2 = 1 Or intComp3 = 1) And intUser1 <> 6 And intUser2 <> 6 _
    And intUser3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For First Row Second
ElseIf (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And (intComp1 = 1 Or intComp2 = 1 Or intComp3 = 1) And intUser1 <> 8 And intUser2 <> 8 _
    And intUser3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For First Row Third
ElseIf (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And intUser1 <> 1 And intUser2 <> 1 _
    And intUser3 <> 1 Then
    intRandomSquare = 1 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Second Row First
ElseIf (intComp1 = 3 Or intComp2 = 3 Or intComp3 = 3) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 7 And intUser2 <> 7 _
    And intUser3 <> 7 Then
    intRandomSquare = 7 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Row Second
ElseIf (intComp1 = 7 Or intComp2 = 7 Or intComp3 = 7) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 3 And intUser2 <> 3 _
    And intUser3 <> 3 Then
    intRandomSquare = 3 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Row Third
ElseIf (intComp1 = 7 Or intComp2 = 7 Or intComp3 = 7) And (intComp1 = 3 Or intComp2 = 3 Or intComp3 = 3) And intUser1 <> 5 And intUser2 <> 5 _
    And intUser3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block





'8 1 6
'3 5 7
'4 9 2
'For Third Row First
ElseIf (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And (intComp1 = 9 Or intComp2 = 9 Or intComp3 = 9) And intUser1 <> 2 And intUser2 <> 2 _
    And intUser3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Third Row Second
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 9 Or intComp2 = 9 Or intComp3 = 9) And intUser1 <> 4 And intUser2 <> 4 _
    And intUser3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    

'8 1 6
'3 5 7
'4 9 2
'For Third Row Third
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And intUser1 <> 9 And intUser2 <> 9 _
    And intUser3 <> 9 Then
    intRandomSquare = 9 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For First Column First
ElseIf (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And (intComp1 = 3 Or intComp2 = 3 Or intComp3 = 3) And intUser1 <> 4 And intUser2 <> 4 _
    And intUser3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For First Column Second
ElseIf (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And (intComp1 = 3 Or intComp2 = 3 Or intComp3 = 3) And intUser1 <> 8 And intUser2 <> 8 _
    And intUser3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For First Column Third
ElseIf (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And intUser1 <> 3 And intUser2 <> 3 _
    And intUser3 <> 3 Then
    intRandomSquare = 3 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Column First
ElseIf (intComp1 = 1 Or intComp2 = 1 Or intComp3 = 1) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 9 And intUser2 <> 9 _
    And intUser3 <> 9 Then
    intRandomSquare = 9 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Column Second
ElseIf (intComp1 = 9 Or intComp2 = 9 Or intComp3 = 9) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 1 And intUser2 <> 1 _
    And intUser3 <> 1 Then
    intRandomSquare = 1 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Column Third
ElseIf (intComp1 = 9 Or intComp2 = 9 Or intComp3 = 9) And (intComp1 = 1 Or intComp2 = 1 Or intComp3 = 1) And intUser1 <> 5 And intUser2 <> 5 _
    And intUser3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Third Column First
ElseIf (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And (intComp1 = 7 Or intComp2 = 7 Or intComp3 = 7) And intUser1 <> 2 And intUser2 <> 2 _
    And intUser3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Third Column Second
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 7 Or intComp2 = 7 Or intComp3 = 7) And intUser1 <> 6 And intUser2 <> 6 _
    And intUser3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Third Column Third
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And intUser1 <> 7 And intUser2 <> 7 _
    And intUser3 <> 7 Then
    intRandomSquare = 7 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
        
'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal First
ElseIf (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 2 And intUser2 <> 2 _
    And intUser3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal Second
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 8 And intUser2 <> 8 _
    And intUser3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal Third
ElseIf (intComp1 = 2 Or intComp2 = 2 Or intComp3 = 2) And (intComp1 = 8 Or intComp2 = 8 Or intComp3 = 8) And intUser1 <> 5 And intUser2 <> 5 _
    And intUser3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Right Diagonal First
ElseIf (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 4 And intUser2 <> 4 _
    And intUser3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
'8 1 6
'3 5 7
'4 9 2
'For Right Diagonal Second
ElseIf (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And (intComp1 = 5 Or intComp2 = 5 Or intComp3 = 5) And intUser1 <> 6 And intUser2 <> 6 _
    And intUser3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
'8 1 6
'3 5 7
'4 9 2
'For Right Diagonal Third
ElseIf (intComp1 = 4 Or intComp2 = 4 Or intComp3 = 4) And (intComp1 = 6 Or intComp2 = 6 Or intComp3 = 6) And intUser1 <> 5 And intUser2 <> 5 _
    And intUser3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'For Defense
'For First Row First
ElseIf (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And (intUser1 = 1 Or intUser2 = 1 Or intUser3 = 1) And intComp1 <> 6 And intComp2 <> 6 _
    And intComp3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For First Row Second
ElseIf (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And (intUser1 = 1 Or intUser2 = 1 Or intUser3 = 1) And intComp1 <> 8 And intComp2 <> 8 _
    And intComp3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For First Row Third
ElseIf (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And intComp1 <> 1 And intComp2 <> 1 _
    And intComp3 <> 1 Then
    intRandomSquare = 1 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block



'For Second Row First
ElseIf (intUser1 = 3 Or intUser2 = 3 Or intUser3 = 3) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 7 And intComp2 <> 7 _
    And intComp3 <> 7 Then
    intRandomSquare = 7 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'For Second Row Second
ElseIf (intUser1 = 7 Or intUser2 = 7 Or intUser3 = 7) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 3 And intComp2 <> 3 _
    And intComp3 <> 3 Then
    intRandomSquare = 3 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'For Second Row Third
ElseIf (intUser1 = 7 Or intUser2 = 7 Or intUser3 = 7) And (intUser1 = 3 Or intUser2 = 3 Or intUser3 = 3) And intComp1 <> 5 And intComp2 <> 5 _
    And intComp3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block



'For Third Row First
ElseIf (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And (intUser1 = 9 Or intUser2 = 9 Or intUser3 = 9) And intComp1 <> 2 And intComp2 <> 2 _
    And intComp3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'For Third Row Second
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 9 Or intUser2 = 9 Or intUser3 = 9) And intComp1 <> 4 And intComp2 <> 4 _
    And intComp3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For Third Row Third
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And intComp1 <> 9 And intComp2 <> 9 _
    And intComp3 <> 9 Then
    intRandomSquare = 9 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'8 1 6
'3 5 7
'4 9 2
'For First Column First
ElseIf (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And (intUser1 = 3 Or intUser2 = 3 Or intUser3 = 3) And intComp1 <> 4 And intComp2 <> 4 _
    And intComp3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
'8 1 6
'3 5 7
'4 9 2
'For First Column Second
ElseIf (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And (intUser1 = 3 Or intUser2 = 3 Or intUser3 = 3) And intComp1 <> 8 And intComp2 <> 8 _
    And intComp3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For First Column Third
ElseIf (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And intComp1 <> 3 And intComp2 <> 3 _
    And intComp3 <> 3 Then
    intRandomSquare = 3 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Second Column First
ElseIf (intUser1 = 1 Or intUser2 = 1 Or intUser3 = 1) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 9 And intComp2 <> 9 _
    And intComp3 <> 9 Then
    intRandomSquare = 9 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'8 1 6
'3 5 7
'4 9 2
'For Second Column Second
ElseIf (intUser1 = 9 Or intUser2 = 9 Or intUser3 = 9) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 1 And intComp2 <> 1 _
    And intComp3 <> 1 Then
    intRandomSquare = 1 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'8 1 6
'3 5 7
'4 9 2
'For Second Column Third
ElseIf (intUser1 = 9 Or intUser2 = 9 Or intUser3 = 9) And (intUser1 = 1 Or intUser2 = 1 Or intUser3 = 1) And intComp1 <> 5 And intComp2 <> 5 _
    And intComp3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Third Column First
ElseIf (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And (intUser1 = 7 Or intUser2 = 7 Or intUser3 = 7) And intComp1 <> 2 And intComp2 <> 2 _
    And intComp3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'8 1 6
'3 5 7
'4 9 2
'For Third Column Second
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 7 Or intUser2 = 7 Or intUser3 = 7) And intComp1 <> 6 And intComp2 <> 6 _
    And intComp3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Third Column Third
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And intComp1 <> 7 And intComp2 <> 7 _
    And intComp3 <> 7 Then
    intRandomSquare = 7 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal First
ElseIf (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 2 And intComp2 <> 2 _
    And intComp3 <> 2 Then
    intRandomSquare = 2 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal Second
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 8 And intComp2 <> 8 _
    And intComp3 <> 8 Then
    intRandomSquare = 8 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'8 1 6
'3 5 7
'4 9 2
'For Left Diagonal Third
ElseIf (intUser1 = 2 Or intUser2 = 2 Or intUser3 = 2) And (intUser1 = 8 Or intUser2 = 8 Or intUser3 = 8) And intComp1 <> 5 And intComp2 <> 5 _
    And intComp3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block
    
'8 1 6
'3 5 7
'4 9 2
'For Right Diagonal First
ElseIf (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 4 And intComp2 <> 4 _
    And intComp3 <> 4 Then
    intRandomSquare = 4 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block


'For Right Diagonal Second
ElseIf (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And (intUser1 = 5 Or intUser2 = 5 Or intUser3 = 5) And intComp1 <> 6 And intComp2 <> 6 _
    And intComp3 <> 6 Then
    intRandomSquare = 6 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

'For Right Diagonal Third
ElseIf (intUser1 = 4 Or intUser2 = 4 Or intUser3 = 4) And (intUser1 = 6 Or intUser2 = 6 Or intUser3 = 6) And intComp1 <> 5 And intComp2 <> 5 _
    And intComp3 <> 5 Then
    intRandomSquare = 5 'assign square to block
    doMove intRandomSquare, intMarker 'move to that position to block

 
Else
        'This will look for a random square that is available
        FindRandomSquare
        
        'This will get a marker for the computer to move
        getMarker
        
        'Assign the integers the value of the functions
        intRandomSquare = FindRandomSquare
        intMarker = getMarker
        
        'This will take the marker and move the image to the square. Performs the sub doMove
        doMove intRandomSquare, intMarker

End If

'Check if game is over
blnCompGameOver = IsGameOver()




End Sub




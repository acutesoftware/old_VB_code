VERSION 5.00
Begin VB.Form frmPuzzle 
   Caption         =   "Tile Puzzle"
   ClientHeight    =   5340
   ClientLeft      =   2535
   ClientTop       =   4860
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   7770
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   495
      Left            =   5880
      TabIndex        =   21
      Top             =   720
      Width           =   1575
   End
   Begin VB.ListBox Text1 
      Height          =   4740
      Left            =   5040
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.CommandButton cmdShuffle 
      Caption         =   "Shuffle"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   19
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   3735
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H000080FF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FF80FF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   960
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H000080FF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1680
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FF80FF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2400
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H000080FF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   240
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FF80FF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   960
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H000080FF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   1680
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FF80FF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   2400
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H0000C000&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   240
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FFFFFF&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   960
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H0000C000&
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   1680
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FFFFFF&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   2400
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H0000C000&
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   240
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H00FFFFFF&
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   960
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         BackColor       =   &H0000C000&
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   1680
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   2400
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdNumber 
      BackColor       =   &H000000C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   0
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Number of Moves ="
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
      Left            =   960
      TabIndex        =   22
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label lblInfo 
      Caption         =   "Label1"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
   End
End
Attribute VB_Name = "frmPuzzle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numy As Integer    'number of squares in puzzle
Dim numx As Integer
Dim spacing As Integer

Dim i As Long
Dim j As Long
Dim startx As Long
Dim starty As Long
Dim current As Integer 'current tile number
'Dim curx As Long     'current grid x pos
'Dim cury As Long     'current grid y pos
Dim blankx As Long     'current grid x pos of the 'blank'
Dim blanky As Long     'current grid y pos of the 'blank'
Dim grid As New clsMap 'create an instance of the map (grid) class


'Dim grid(16, 16)        'max grid size
'TODO : add clock with single reminder to it( calendar control?)
'       nice background with photo of kids and scanned image of coffee/smokes
'       control panel (shape of TV remote) to turn off sound, animation, messages, etc
'       high score screens with seperate peoples names


Public Sub Initialise(x, y)
    current = 1
    spacing = 735
    blankx = x
    blanky = y
    grid.GridHeight = y
    grid.GridWidth = x
    For j = 1 To y
      For i = 1 To x
        grid.CellDetails(i, j) = CLng(current)
        grid.CellPositionX(i, j) = CInt(startx + (spacing * (i)) - spacing / 2)
        grid.CellPositionY(i, j) = CInt(starty + (spacing * (j)) - spacing / 2)
        current = current + 1
        'GetNumberFromGrid(i, j)
        Debug.Print "Initialising grid: x=" & Str$(i) & "y=" & Str$(j) & " CellDetails=" & Str$(grid.CellDetails(i, j)) & " screen position is : x=" & grid.CellPositionX(i, j) & "y=" & grid.CellPositionY(i, j)
        'cmdNumber(current).Width = spacing
        'cmdNumber(current).Height = spacing
        'cmdNumber(current).Caption = Str$(current)
       ' grid(i, j) = current
        'cmdNumber(current).Move startx + (spacing * (i)) - spacing / 2, starty + (spacing * (j)) - spacing / 2
                    'cmdNumber(current).Move startx + (spacing * (i + 1)) - spacing / 2, starty + (spacing * (j + 1)) - spacing / 2
      Next i
    Next j
   
    Debug.Print " STARTING PROGRAM----------------" & Time$ & Date$

End Sub
Public Function DisplayGrid()
    Text1.Clear
    For j = 1 To grid.GridWidth
      For i = 1 To grid.GridHeight
            cmdNumber(Val(grid.CellDetails(i, j))).Move grid.CellPositionX(i, j), grid.CellPositionY(i, j)
            cmdNumber(grid.CellDetails(i, j)).Visible = True
            Text1.AddItem "x=" & Str$(j) & "y=" & Str$(i) & "grid.CellDetails(x,y)= " & grid.CellDetails(i, j) & Chr$(9) & "ScreenX= " & grid.CellPositionX(i, j) & "    ScreenY= " & grid.CellPositionY(i, j)

      Next i
    Next j

End Function

Public Sub ShuffleDown(x As Long, y_orig As Long, y_dest As Long)
        Dim trips As Integer    'number of trips
'        lblInfo.Caption = "MOVING DOWN from grid" & Str$(x) & Str$(x) & "  to grid" & Str$(x) & Str$(y_dest)
        trips = y_orig - y_dest
        Select Case trips
        Case 1:     'single shift down
            SwapCells x, y_orig, x, y_dest
            blanky = y_orig
        Case 2:     '2 moves down
        
        Case 3:     '3 moves down
        
        Case -1:    'single shift up
            SwapCells x, y_orig, x, y_dest
            blanky = y_orig
        Case -2:    '2 moves up
        
        Case -2:    '3 moves up
        
        End Select

End Sub
Public Sub ShuffleAcross(x_orig As Long, x_dest As Long, y As Long)
        Dim trips As Integer    'number of trips
'        lblInfo.Caption = "MOVING ACROSS from grid" & Str$(x_orig) & Str$(y) & "  to grid" & Str$(x_dest) & Str$(y)
        trips = x_orig - x_dest
        Select Case trips
        Case 1:     'single shift to right
            SwapCells x_orig, y, x_dest, y
            blankx = x_orig
        Case 2:     '2 moves to right
        
        Case 3:     '3 moves to right
        
        Case -1:    'single shift to left
            SwapCells x_orig, y, x_dest, y
            blankx = x_orig
        Case -2:    '2 moves to left
        
        Case -2:    '3 moves to left
        
        End Select
End Sub
Private Sub SwapCells(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    'this swaps all values for cells and display the changes
    Dim tempDetails As Integer
    Dim tempScreenX As Integer
    Dim tempScreenY As Integer
    i = sound("move.wav")
    tempScreenX = grid.CellPositionX(x1, y1)
    tempScreenY = grid.CellPositionY(x1, y1)
    tempDetails = grid.CellDetails(x1, y1)
    grid.CellDetails(x1, y1) = grid.CellDetails(x2, y2)
    grid.CellDetails(x2, y2) = tempDetails
    
    lblScore.Caption = Str$(Val(lblScore.Caption) + 1)
    'move 'slowly the 'nonblank cell to its new position
'    If grid.CellPositionX(x1, y1) = grid.CellPositionX(x2, y2) Then 'move down
'      For i = grid.CellPositionY(x1, y1) To grid.CellPositionY(x2, y2) Step 70
'        cmdNumber(grid.CellDetails(x1, y1)).Move grid.CellPositionX(x1, y1), i
'        cmdNumber(grid.CellDetails(x1, y1)).Refresh
'        For j = 1 To 25
'            DoEvents
'        Next j
'      Next i
'
'    Else    'slow cool move across
'      For i = grid.CellPositionX(x1, y1) To grid.CellPositionX(x2, y2) Step 70
'        cmdNumber(grid.CellDetails(x1, y1)).Move i, grid.CellPositionY(x1, y1)
'        cmdNumber(grid.CellDetails(x1, y1)).Refresh
'        For j = 1 To 25
'            DoEvents
'        Next j
'      Next i
'    End If
    'now move the 'blank' one fast
  'WORKS
  If Text1.Enabled = True Then  'this means we are still debugging
    DisplayGrid
  
  Else
    cmdNumber(grid.CellDetails(x1, y1)).Move grid.CellPositionX(x1, y1), grid.CellPositionY(x1, y1)
    cmdNumber(grid.CellDetails(x2, y2)).Move grid.CellPositionX(x2, y2), grid.CellPositionY(x2, y2)
  End If

End Sub

Private Sub cmdNumber_Click(Index As Integer)
    current = Index
    i = grid.CalculateCurrentCoordinates(current)    'gets the x,y co-ord for button clicked
    'curx = grid.x_current
    'cury = grid.y_current
    
    lblInfo.Caption = "clicked tile " & Str$(current)
   ' GetGridCoordinatesFromNumber (current)
    If blankx = grid.x_current Then ' correct x
        'cmdNumber(GetNumberFromGrid(cury, curx)).Move cmdNumber(16).Left, cmdNumber(16).Top
        ShuffleDown grid.x_current, grid.y_current, blanky
    End If
    If blanky = grid.y_current Then ' correct y
        'lblInfo.Caption = "MOVING Y"
        ShuffleAcross grid.x_current, blankx, grid.y_current
    
    End If
    
    
End Sub

Private Sub cmdRestart_Click()
        Initialise 4, 4
        lblScore.Caption = "0"
        'GenerateGrid
        DisplayGrid
        cmdShuffle.Enabled = True

End Sub

Private Sub cmdShuffle_Click()
    grid.Shuffle
    current = 15
   lblScore.Caption = "0"

    i = grid.CalculateCurrentCoordinates(16)
    blankx = grid.x_current
    blanky = grid.y_current
    grid.x_current = 1
    grid.y_current = 1
    
    DisplayGrid
End Sub


Private Sub Form_Load()
    ChDir ("c:\dev\src\vb6\games\puzzle1")
        Call cmdRestart_Click
End Sub

Private Sub Form_Resize()
    'If Me.Width > 5000 And cmdShuffle.Enabled Then
    '    Frame1.Width = Me.Width - 4000
    'End If
    'If Me.Height > 3000 And cmdShuffle.Enabled Then
    '    Frame1.Height = Me.Height - 2000
    'End If
End Sub


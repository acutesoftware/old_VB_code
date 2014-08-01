VERSION 5.00
Object = "{C3CBDAAF-C8D1-11D2-9F8E-0080C7CE5CDC}#1.0#0"; "ACTCNDY3.OCX"
Object = "{C3CBD80D-C8D1-11D2-9F8E-0080C7CE5CDC}#4.1#0"; "ACTCNDY2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "game1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "game1.frx":0442
   ScaleHeight     =   5865
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerIntro 
      Interval        =   1000
      Left            =   5640
      Top             =   4200
   End
   Begin VB.Timer TimerGame 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   6960
      Top             =   3720
   End
   Begin VB.Timer TimerBalls 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   6960
      Top             =   2280
   End
   Begin ActiveCandy.CandyCommand cmdSpeedUp 
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackPicture     =   1
      Caption         =   ">"
      ForeColor       =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasShadow       =   -1  'True
      PictureGoDown   =   -1  'True
      UseSound        =   -1  'True
      SoundMethod     =   1
   End
   Begin ActiveCandy2.CandyFill fillSpeed 
      Height          =   630
      Left            =   6960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   1111
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   8421376
      ForeColor       =   -2147483630
      CandyFill       =   1
      CaptionStyle    =   1
   End
   Begin ActiveCandy.CandySphere ball 
      Height          =   1080
      Index           =   0
      Left            =   7680
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      SphereType      =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconSize        =   3
      PictureMode     =   1
      UseOnTop        =   0   'False
      UseSound        =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4680
      Top             =   2280
   End
   Begin ActiveCandy.CandyCommand cmdScore 
      Height          =   1095
      Left            =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1931
      BackPicture     =   1
      Caption         =   "0000"
      ForeColor       =   255
      FontName        =   "MS Sans Serif"
      FontSize        =   29.25
      FontBold        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasShadow       =   -1  'True
      UseSound        =   -1  'True
      SoundMethod     =   1
   End
   Begin ActiveCandy.CandyHCommand CandyHCommand2 
      Height          =   480
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "News Gothic MT"
      FontSize        =   11.25
      ForeColor       =   12648447
      Caption         =   "QUIT"
      FocusColor      =   8454143
      UseAnimation    =   -1  'True
      UseOnTop        =   0   'False
      SoundMethod     =   1
   End
   Begin ActiveCandy2.CandyPad CandyPad1 
      Height          =   1800
      Left            =   10680
      TabIndex        =   1
      Top             =   360
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   3175
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseDirectionShortKeys=   -1  'True
   End
   Begin ActiveCandy.CandySphere planet 
      Height          =   600
      Index           =   0
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   600
      _ExtentX        =   847
      _ExtentY        =   847
      SphereType      =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconSize        =   2
      FocusColor      =   16576
      PictureMode     =   1
      UseSound        =   -1  'True
      SoundMethod     =   1
   End
   Begin ActiveCandy.CandySphere ball 
      CausesValidation=   0   'False
      Height          =   360
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   847
      _ExtentY        =   847
      ContentType     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconSize        =   3
      PictureMode     =   1
      UseSound        =   -1  'True
      SoundMethod     =   1
   End
   Begin ActiveCandy.CandySphere ball 
      Height          =   600
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      SphereType      =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconSize        =   3
      PictureMode     =   1
      UseSound        =   -1  'True
   End
   Begin ActiveCandy.CandySphere ball 
      Height          =   825
      Index           =   3
      Left            =   3600
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   825
      _ExtentX        =   847
      _ExtentY        =   847
      SphereType      =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureMode     =   1
      UseSound        =   -1  'True
   End
   Begin ActiveCandy.CandyCommand cmdSlowDown 
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackPicture     =   1
      Caption         =   "<"
      ForeColor       =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasShadow       =   -1  'True
      PictureGoDown   =   -1  'True
      UseSound        =   -1  'True
      SoundMethod     =   1
   End
   Begin ActiveCandy.CandyHCommand CandyHCommand1 
      Height          =   480
      Left            =   3840
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      FontName        =   "News Gothic MT"
      FontSize        =   11.25
      ForeColor       =   12648447
      Caption         =   "NEW GAME"
      FocusColor      =   8454143
      UseAnimation    =   -1  'True
      UseOnTop        =   0   'False
      SoundMethod     =   1
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00000000&
      Caption         =   "Level 1 - Demo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1350
      Left            =   4560
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   20
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   19
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   18
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   17
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   16
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   15
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   14
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   13
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   12
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   11
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   10
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   9
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   8
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   7
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   6
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   5
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   4
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   3
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   2
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   1
      X1              =   4080
      X2              =   0
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Line line1 
      BorderColor     =   &H00FF0000&
      Index           =   0
      X1              =   9000
      X2              =   4920
      Y1              =   5040
      Y2              =   -720
   End
   Begin ActiveCandy.CandyScreen CandyScreen1 
      Height          =   2775
      Left            =   240
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   4895
      MouseIcon       =   "game1.frx":0594
      ScreenType      =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xscreen As Integer      'max size of screen
Dim yscreen As Integer      'max size of screen
'your position and info
Dim xdir As Integer
Dim ydir As Integer
Dim xpos As Integer
Dim ypos As Integer
Dim speed As Integer
Dim maxspeed As Integer
Dim res As Integer        'general counter
Dim i As Integer        'general counter
Dim j As Integer        'general counter
Dim score As Integer    'duh
Dim orbits(4, 100)  'orbit calculation for planets
Dim xplan(3) As Integer 'x position of planets
Dim yplan(3) As Integer 'y position of planets
Dim speedplanet(3) As Integer   'speed of planets
Public Sub NewGame()
    xplan(0) = 1000: yplan(0) = 3000:    speedplanet(0) = 50
    xplan(1) = 2500: yplan(1) = 5000:    speedplanet(1) = 100
    xplan(2) = 5000: yplan(2) = 1000:    speedplanet(2) = 70
    xplan(3) = 7300: yplan(3) = 3500:    speedplanet(3) = 30
    speed = 50
    xpos = 3000
    ypos = 3000
    score = 0
    res = sound("103_comin.wav")
    For i = 0 To 3
        ball(i).Move xplan(i), yplan(i)
        ball(i).Visible = True
    Next i
    xdir = 1
    ydir = 1
    planet(0).Move xpos, ypos
    planet(0).Visible = True
    SendKeys "{TAB}"
    Timer1.Enabled = True
    TimerBalls.Enabled = True
    TimerGame.Enabled = True

End Sub

Public Sub Initialise()
    Timer1.Enabled = False
    TimerBalls.Enabled = False
    TimerGame.Enabled = False
    ChDir App.Path
    For i = 0 To 3
        For j = 0 To 100
        orbits(i, j) = Sin(j)
       ' ball(i).Move Rnd(8000) + 1000, Rnd(4000) + 1000
        Next j
    Next i
    maxspeed = 500
    xscreen = Screen.Width
    yscreen = Screen.Height - 3000
    
    'setup the screen -----------------------------------------
    CandyScreen1.Left = 0
    CandyScreen1.Top = Screen.Height - 2500
    CandyScreen1.Height = 2500
    CandyScreen1.Width = Screen.Width
    i = 600 + CandyScreen1.Top  'how far inside control box commands are
    cmdScore.Top = i
    CandyHCommand1.Top = i
    CandyHCommand2.Top = i + 600
    cmdSpeedUp.Top = i
    cmdSlowDown.Top = i
    fillSpeed.Top = i + 500
    CandyPad1.Top = i - 300
End Sub
Public Sub GameOver()
    Timer1.Enabled = False
    TimerBalls.Enabled = False
    TimerGame.Enabled = False
    lblMessage.Caption = "Game Over"
    lblMessage.Move Me.ScaleWidth / 2 - (lblMessage.Width / 2), Me.ScaleHeight / 2 - (lblMessage.Height / 2)
    lblMessage.Visible = True
    res = sound("kickass.wav")
    'MsgBox "Congratulations, you scored " & cmdScore.Caption & " points!!"
    

End Sub
Public Sub DrawTerrain()
    For i = 0 To 20
        'Dim line1(i) As New Line
        'line1(i).DrawMode
       ' Line (0, i * 200)-(xscreen, i * 200)
        line1(i).X1 = 500
        line1(i).Y1 = i * 200 + 1000
        line1(i).Y2 = i * 200 + 1000
        line1(i).X2 = xscreen
        line1(i).BorderStyle = 1
        'line1(i).BorderColor = i
      '  Debug.Print line1(i).BorderColor
        line1(i).Visible = False
     Next i
End Sub

Public Sub ball_Click(Index As Integer)
    'ball(i).Visible = False
End Sub

Private Sub CandyHCommand1_Click()
    NewGame
    lblMessage.Visible = False

End Sub

Private Sub CandyHCommand2_Click()
    End
End Sub

Private Sub CandyPad1_GoDown()
    ydir = 1
    'res = sound("blup1.wav") ' Or SND_ASYNC)

End Sub

Private Sub CandyPad1_GoLeft()
    xdir = -1
    'res = sound("blup1.wav") ' Or SND_ASYNC)

End Sub

Private Sub CandyPad1_GoRight()
    xdir = 1
    'res = sound("blup1.wav") ' Or SND_ASYNC)
End Sub

Private Sub CandyPad1_GoUp()
    ydir = -1
    'res = sound("blup1.wav") ' Or SND_ASYNC)
End Sub

Private Sub CandyScreen1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    res = sound("shot3.wav")

End Sub

Private Sub cmdScore_Click()
    MsgBox "Your score is " & cmdScore.Caption
End Sub

Private Sub cmdSlowDown_Click()
    If speed > 10 Then
        speed = speed - 10
    End If
    fillSpeed.Value = 100 * speed / maxspeed
        SendKeys "{TAB}"

End Sub

Private Sub cmdSpeedUp_Click()
    If speed < maxspeed Then
        speed = speed + 10 ' = CandyPad1.XValue
    End If
    fillSpeed.Value = 100 * (speed / maxspeed)
    SendKeys "{TAB}"

End Sub

Private Sub Form_Load()
    Me.WindowState = 2
    lblMessage.Caption = "Level 1 - Demo"
    lblMessage.Move Me.ScaleWidth / 2 - (lblMessage.Width / 2), Me.ScaleHeight / 2 - (lblMessage.Height / 2)
    lblMessage.Visible = True

   Initialise
    DrawTerrain
End Sub

Private Sub Timer1_Timer()
    'this timer moves you, the player (the blue ball)
    'change direction if we are near edges
    If xpos < 500 Then xpos = 500: xdir = 1: res = sound("hit1.wav"): score = score - 1: cmdScore.Caption = Str$(score)
    If xpos > xscreen Then xpos = xscreen: xdir = -1: res = sound("hit1.wav"): score = score - 1: cmdScore.Caption = Str$(score)
    If ypos < 500 Then ypos = 500: ydir = 1: res = sound("hit1.wav"): score = score - 1: cmdScore.Caption = Str$(score)
    If ypos > yscreen Then ypos = yscreen: ydir = -1: res = sound("hit1.wav"): score = score - 1: cmdScore.Caption = Str$(score)
    
    'calculate next position
    xpos = xpos + (speed * xdir)
    ypos = ypos + (speed * ydir)
    
    'move to next position
    planet(0).Move xpos, ypos
    
    'check if we are dead
    For i = 0 To 3
        If xpos > ball(i).Left - ball(i).Width And xpos < ball(i).Left + ball(i).Width Then
            If ypos > ball(i).Top - ball(i).Height And ypos < ball(i).Top + ball(i).Height Then
            'COLLISION
              ' res = sound("alex_setup.wav")
               res = sound("expols3.wav")

                ball_Click (i)
                score = score + 1
                xdir = -xdir
                ydir = -ydir
                cmdScore.Caption = Str$(score)
            End If
        End If
    Next i
End Sub

Private Sub TimerBalls_Timer()
    For i = 0 To 3
        yplan(i) = yplan(i) + speedplanet(i)
        If yplan(i) > yscreen Then
            yplan(i) = 100
            xplan(i) = Int(Rnd(1) * (xscreen - 100)) + 50
            speedplanet(i) = ((Rnd(1) * 10) * 5) + 50
            'ball(i).SphereType = Rnd(5) + 3
            'res = (Rnd(10) + 2) * 200 'size
            'ball(i).Width = res
            'ball(i).Height = res
        End If
        ball(i).Move xplan(i), yplan(i)
    Next i
End Sub

Private Sub TimerGame_Timer()
    speed = speed + 10
    fillSpeed.Value = 100 * (speed / maxspeed)
    If fillSpeed.Value > 99 Then GameOver   'TEMP end game here
    
End Sub

Private Sub TimerIntro_Timer()
    'only used once (yeh, right!) at start of game to play intro
    res = sndPlaySound("welcome to planet smasher.wav", &H1)
    TimerIntro.Enabled = False
End Sub

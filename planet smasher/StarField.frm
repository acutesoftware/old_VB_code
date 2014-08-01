VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Starfield"
   ClientHeight    =   7956
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7836
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   663
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   653
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar Scroll 
      Height          =   372
      LargeChange     =   100
      Left            =   120
      Max             =   1000
      Min             =   1
      SmallChange     =   5
      TabIndex        =   1
      Top             =   6480
      Value           =   1
      Width           =   7572
   End
   Begin VB.PictureBox BackRound 
      BackColor       =   &H00000000&
      Height          =   6288
      Left            =   120
      ScaleHeight     =   520
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   630
      TabIndex        =   0
      Top             =   120
      Width           =   7608
   End
   Begin VB.Label Label1 
      Caption         =   $"StarField.frx":0000
      Height          =   492
      Left            =   120
      TabIndex        =   3
      Top             =   7440
      Width           =   7572
   End
   Begin VB.Label Label 
      Caption         =   "Delay:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   7572
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StarArray(99) As Coordinate
Dim XChange As Integer
Dim YChange As Integer
Dim Key As Integer
Const White = &HFFFFFF
Const Black = &H0&

Private Sub BackRound_Click()

Dim Ctr As Integer
Dim Ctr2 As Integer

Do
For Ctr = 0 To 99

If StarArray(Ctr).X > 315 And StarArray(Ctr).Y > 210 Then
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal Black
     StarArray(Ctr).X = StarArray(Ctr).X + (Int((StarArray(Ctr).X - 315) / 10) + 1) + XChange
     StarArray(Ctr).Y = StarArray(Ctr).Y + (Int((StarArray(Ctr).Y - 210) / 10) + 1) + YChange
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal White
ElseIf StarArray(Ctr).X > 315 And StarArray(Ctr).Y <= 210 Then
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal Black
     StarArray(Ctr).X = StarArray(Ctr).X + (Int((StarArray(Ctr).X - 315) / 10) + 1) + XChange
     StarArray(Ctr).Y = StarArray(Ctr).Y - (Int((210 - StarArray(Ctr).Y) / 10) + 1) + YChange
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal White
ElseIf StarArray(Ctr).X <= 315 And StarArray(Ctr).Y > 210 Then
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal Black
     StarArray(Ctr).X = StarArray(Ctr).X - (Int((315 - StarArray(Ctr).X) / 10) + 1) + XChange
     StarArray(Ctr).Y = StarArray(Ctr).Y + (Int((StarArray(Ctr).Y - 210) / 10) + 1) + YChange
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal White
ElseIf StarArray(Ctr).X <= 315 And StarArray(Ctr).Y <= 210 Then
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal Black
     StarArray(Ctr).X = StarArray(Ctr).X - (Int((315 - StarArray(Ctr).X) / 10) + 1) + XChange
     StarArray(Ctr).Y = StarArray(Ctr).Y - (Int((210 - StarArray(Ctr).Y) / 10) + 1) + YChange
     PixelSet BackRound, ByVal StarArray(Ctr).X, ByVal StarArray(Ctr).Y, ByVal White
Else
End If
     
If StarArray(Ctr).X < 0 Or StarArray(Ctr).X > 630 Or StarArray(Ctr).Y < 0 Or StarArray(Ctr).Y > 520 Then
     BackRound.PSet (StarArray(Ctr).X, StarArray(Ctr).Y), Black
     StarArray(Ctr).X = Int(Rnd * 630)
     StarArray(Ctr).Y = Int(Rnd * 520)
     BackRound.PSet (StarArray(Ctr).X, StarArray(Ctr).Y), White
End If
     
Next
For Ctr2 = 1 To Scroll.Value
DoEvents
DoMainProc
Next
Loop

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Key = KeyAscii
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Key = 0
End Sub

Private Sub Form_Load()

Dim Ctr As Integer

For Ctr = 0 To 99
     StarArray(Ctr).X = Int(630 * Rnd)
     StarArray(Ctr).Y = Int(630 * Rnd)
     BackRound.PSet (StarArray(Ctr).X, StarArray(Ctr).Y), White
Next

Label = "Delay: " & Scroll

End Sub

Private Sub Scroll_Change()
Label = "Delay: " & Scroll
End Sub

Sub DoMainProc()

If Key = 0 Then
     XChange = 0
     YChange = 0
     Exit Sub
End If
If Key = KLeft Then XChange = 40
If Key = KRight Then XChange = -40
If Key = KUp Then YChange = -40
If Key = KDown Then YChange = 40
If Key = 27 Then End

End Sub

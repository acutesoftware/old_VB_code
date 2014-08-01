Attribute VB_Name = "StarModule"
Declare Function SetPixelV Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Type Coordinate
     X As Integer
     Y As Integer
End Type

'Here are some constants for use with functions, subs and games
Global Const KLeft = 52
Global Const KUp = 56
Global Const KRight = 54
Global Const KDown = 50

Sub PixelSet(Obj As PictureBox, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)

Dim Z As Long
Z = SetPixelV(ByVal Obj.hdc, ByVal X, ByVal Y, ByVal Color)

End Sub

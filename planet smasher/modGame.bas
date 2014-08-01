Attribute VB_Name = "modGame"
Option Explicit

Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10


Public Function sound(fname As String) As Integer
    Dim X%
    Dim wFlags%
    'wFlags% = SND_ASYNC Or SND_NODEFAULT
    'wFlags% = SND_ASYNC Or SND_NODEFAULT
    wFlags% = SND_ASYNC 'Or SND_NOSTOP
    X% = sndPlaySound(fname, wFlags%)

End Function

Attribute VB_Name = "modSound"
'*********************************************************************
'   Thank´s to Brian for this Module
'*********************************************************************
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long
Private Const SND_SYNC = &H0 ' Don't return until sound ends (default).
Private Const SND_ASYNC = &H1 ' Return immediately after the sound starts.
Private Const SND_NODEFAULT = &H2 ' If the sound file is not found, do NOT play default sound.
Private Const SND_MEMORY = &H4 ' Play a sound from a buffer in memory.
Private Const SND_LOOP = &H8 ' Loop sound continuously (used with SND_ASYNC)
Private Const SND_NOSTOP = &H10 ' Don't stop current sound to play another.

Public Sub s_Playsound(SoundName As String)
    
    SoundName = App.path & "\sound\" & SoundName & ".wav"
    
    sndPlaySound SoundName, SND_ASYNC Or SND_NODEFAULT

End Sub



Attribute VB_Name = "modStartmenyDelay"
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Delay1(MilliSeconds As Long)

Dim Start As Long
    Start = GetTickCount

Do While GetTickCount < Start + MilliSeconds
        DoEvents
Loop

End Sub

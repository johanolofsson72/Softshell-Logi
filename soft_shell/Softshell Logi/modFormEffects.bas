Attribute VB_Name = "modFormEffects"
Option Explicit

Public Sub MakeFormEffect(FormEffect As String, FormA As Form, DelayValue As Long, Optional StepValue As Long = 200)

Dim mTop As Long, mHeigh As Long, mB As Long, mA As Long, i As Long
Dim mHeight As Long
Dim mWidth As Long

 Select Case FormEffect
 
  Case "up"
      
      mTop = FormA.Top
      mA = FormA.Top
      mHeigh = FormA.Height
      FormA.Top = FormA.Top + FormA.Height
      FormA.Height = 0
      mB = FormA.Top
      
      FormA.Show
      FormA.Refresh
      
      For i = mB To mA Step -StepValue '100
        FormA.Top = i
        FormA.Height = mB - FormA.Top
        FormA.Refresh
        Call Delay1(DelayValue) '//Delaytime for the menu to be finished
      Next
      
      FormA.Top = mTop
      FormA.Height = mHeigh
      
  Case "down"
      
      mHeight = FormA.Height
      FormA.Height = 100
      
      FormA.Show
      FormA.Refresh
      
      For i = 10 To 1 Step -StepValue '1
        FormA.Height = mHeight / i
        FormA.Refresh
        Call Delay1(DelayValue) '//Delaytime for the menu to be finished
      Next
      
      FormA.Height = mHeight
      
  Case "side"
      
      mWidth = FormA.Width
      FormA.Width = 100
      
      FormA.Show
      FormA.Refresh
      
      For i = 10 To 1 Step -StepValue '1
        FormA.Width = mWidth / i
        FormA.Refresh
        Call Delay1(DelayValue) '//Delaytime for the menu to be finished
      Next
      
      FormA.Width = mWidth
      
  Case "win98"
    
      mHeight = FormA.Height
      mWidth = FormA.Width
      FormA.Height = 100
      FormA.Width = 100
      
      FormA.Show
      FormA.Refresh

      For i = 10 To 1 Step -StepValue '1
        FormA.Height = mHeight / i
        FormA.Width = mWidth / i
        FormA.Refresh
        Call Delay1(DelayValue) '//Delaytime for the menu to be finished
      Next

      FormA.Height = mHeight
      FormA.Width = mWidth

  End Select

End Sub

Attribute VB_Name = "ArcFade"
'ArcFade 1.0 By:Arc 12/2/01
'ArcSoftwareDesign@Hotmail.com

'This is the first Fader.Bas that I have seen, that does this.

'The PulseBackFade will fade any Object
'including(Forms,Labels,Buttons,Pictureboxes, any Object
'in VB that has a BackColor property.
'The PulseForeFade will fade any Object
'including(Forms,Labels,Buttons,Pictureboxes, any Object
'in VB that has a ForeColor property.
'The PulseFades fade your Objects from one color to another
'and then back again.

'Example of use:
'PulseBackFade 255,0,34,134,234,23,Label1,0.3,10
' THe first 3 numbers are your first color in RGB value. Red From 0-255,
'Green from 0-255,Blue from 0-255. The next 3 numbers are your second color.
'Then you have the object you want to fade.
'Then you set the Fade speed
'Then you set how much you want the color to change with each loop.
'For example: Setting Step to 1 will provide a much smoother fade then
'setting Step to 40. This also effects the speed of the fade.

'To see the Color Values in the 3 labels Uncomment the Lines in the Subs

Give credit where it is due. If you adapt this or add to it Add my name and e-mail address please.






Public Red As Long, Green As Long, Blue As Long, Color As Long
Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
 DoEvents
Loop
End Sub
Sub PulseBackFade(C1Val1 As Integer, C1val2 As Integer, C1Val3 As Integer, C2Val1 As Integer, C2Val2 As Integer, C2Val3 As Integer, OB As Object, FadeSpeed, Step)
Dim RedMinus As Boolean, RedPlus As Boolean, GreenMinus As Boolean, GreenPlus As Boolean, BlueMinus As Boolean, BluePlus As Boolean
Red = C1Val1
Green = C1val2
Blue = C1Val3
Normal:

Do
DoEvents
If Red = C2Val1 And Green = C2Val2 And Blue = C2Val3 Then
GoTo reverse
End If
 If C1Val1 >= C2Val1 Then
  RedMinus = True
   Red = Red - Step
    If Red <= C2Val1 Then Red = C2Val1
  End If
   If C1Val1 <= C2Val1 Then
  RedPlus = True
   Red = Red + Step
    If Red >= C2Val1 Then Red = C2Val1
  End If
  
   If C1val2 >= C2Val2 Then
 GreenMinus = True
   Green = Green - Step
    If Green <= C2Val2 Then Green = C2Val2
  End If
   If C1val2 <= C2Val2 Then
  GreenPlus = True
   Green = Green + Step
    If Green >= C2Val2 Then Green = C2Val2
  End If
  
  If C1Val3 >= C2Val3 Then
  BlueMinus = True
   Blue = Blue - Step
    If Blue <= C2Val3 Then Blue = C2Val3
  End If
   If C1Val3 <= C2Val3 Then
  BluePlus = True
   Blue = Blue + Step
    If Blue >= C2Val3 Then Blue = C2Val3
  End If
   Color = RGB(Red, Green, Blue)
   'UNCOMMENT THESE LINES TO SEE THE NUMERIC VALUE IN THE LABELS
 'ArcFadeFrm.lblRed = "Red:" & Red
 'ArcFadeFrm.lblGreen = "Green:" & Green
 'ArcFadeFrm.lblBlue = "Blue:" & Blue
OB.BackColor = Color
Pause (FadeSpeed)
Loop

reverse:
Do
DoEvents
If Red = C1Val1 And Green = C1val2 And Blue = C1Val3 Then
GoTo Normal
End If
If RedMinus = True Then
Red = Red + Step
If Red >= C1Val1 Then Red = C1Val1
End If
If RedPlus = True Then
Red = Red - Step
If Red <= C1Val1 Then Red = C1Val1
End If

If GreenMinus = True Then
Green = Green + Step
If Green >= C1val2 Then Green = C1val2
End If
If GreenPlus = True Then
Green = Green - Step
If Green <= C1val2 Then Green = C1val2
End If

If BlueMinus = True Then
Blue = Blue + Step
If Blue >= C1Val3 Then Blue = C1Val3
End If
If BluePlus = True Then
Blue = Blue - Step
If Blue <= C1Val3 Then Blue = C1Val3
End If
   Color = RGB(Red, Green, Blue)
   'UNCOMMENT THESE LINES TO SEE THE NUMERIC VALUE IN THE LABELS
 'ArcFadeFrm.lblRed = "Red:" & Red
 'ArcFadeFrm.lblGreen = "Green:" & Green
 'ArcFadeFrm.lblBlue = "Blue:" & Blue
OB.BackColor = Color
Pause (FadeSpeed)
Loop
End Sub
Sub PulseForeFade(C1Val1 As Integer, C1val2 As Integer, C1Val3 As Integer, C2Val1 As Integer, C2Val2 As Integer, C2Val3 As Integer, OB As Object, FadeSpeed, Step)
Dim RedMinus As Boolean, RedPlus As Boolean, GreenMinus As Boolean, GreenPlus As Boolean, BlueMinus As Boolean, BluePlus As Boolean
Red = C1Val1
Green = C1val2
Blue = C1Val3
Normal:

Do
DoEvents
If Red = C2Val1 And Green = C2Val2 And Blue = C2Val3 Then
GoTo reverse
End If
 If C1Val1 >= C2Val1 Then
  RedMinus = True
   Red = Red - Step
    If Red <= C2Val1 Then Red = C2Val1
  End If
   If C1Val1 <= C2Val1 Then
  RedPlus = True
   Red = Red + Step
    If Red >= C2Val1 Then Red = C2Val1
  End If
  
   If C1val2 >= C2Val2 Then
 GreenMinus = True
   Green = Green - Step
    If Green <= C2Val2 Then Green = C2Val2
  End If
   If C1val2 <= C2Val2 Then
  GreenPlus = True
   Green = Green + Step
    If Green >= C2Val2 Then Green = C2Val2
  End If
  
  If C1Val3 >= C2Val3 Then
  BlueMinus = True
   Blue = Blue - Step
    If Blue <= C2Val3 Then Blue = C2Val3
  End If
   If C1Val3 <= C2Val3 Then
  BluePlus = True
   Blue = Blue + Step
    If Blue >= C2Val3 Then Blue = C2Val3
  End If
   Color = RGB(Red, Green, Blue)
   'UNCOMMENT THESE LINES TO SEE THE NUMERIC VALUE IN THE LABELS
 'ArcFadeFrm.lblRed = "Red:" & Red
' ArcFadeFrm.lblGreen = "Green:" & Green
 'ArcFadeFrm.lblBlue = "Blue:" & Blue
OB.ForeColor = Color
Pause (FadeSpeed)
Loop

reverse:
Do
DoEvents
If Red = C1Val1 And Green = C1val2 And Blue = C1Val3 Then
GoTo Normal
End If
If RedMinus = True Then
Red = Red + Step
If Red >= C1Val1 Then Red = C1Val1
End If
If RedPlus = True Then
Red = Red - Step
If Red <= C1Val1 Then Red = C1Val1
End If

If GreenMinus = True Then
Green = Green + Step
If Green >= C1val2 Then Green = C1val2
End If
If GreenPlus = True Then
Green = Green - Step
If Green <= C1val2 Then Green = C1val2
End If

If BlueMinus = True Then
Blue = Blue + Step
If Blue >= C1Val3 Then Blue = C1Val3
End If
If BluePlus = True Then
Blue = Blue - Step
If Blue <= C1Val3 Then Blue = C1Val3
End If
   Color = RGB(Red, Green, Blue)
'UNCOMMENT THESE LINES TO SEE THE NUMERIC VALUE IN THE LABELS
 'ArcFadeFrm.lblRed = "Red:" & Red
 'ArcFadeFrm.lblGreen = "Green:" & Green
 'ArcFadeFrm.lblBlue = "Blue:" & Blue
   OB.ForeColor = Color
Pause (FadeSpeed)
Loop

End Sub

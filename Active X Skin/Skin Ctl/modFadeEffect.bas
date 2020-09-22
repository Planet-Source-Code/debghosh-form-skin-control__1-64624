Attribute VB_Name = "modFadeEffect"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const FadeDelay = 5
Const FadeMax = 255
 
 'Global Const GWL_EXSTYLE = (-20)
 'Global Const WS_EX_LAYERED = &H80000
 Global Const LWA_ALPHA = &H2

 


 Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
 Global LastAlpha As Long

Sub FadeIn(f As Form)
Dim c
Dim ne As Integer, en(32767) As Boolean
'For Each c In f.Controls
 'ne = ne + 1
 'en(ne) = c.Enabled
 'c.Enabled = False
'Next
   TransForm f.hwnd, 0
    f.Show
    Dim i As Long
    For i = 0 To FadeMax Step 3
        TransForm f.hwnd, CByte(i)
        DoEvents
        Call Sleep(FadeDelay)
    Next
    TransForm f.hwnd, FadeMax
    i = 0
'For Each c In f.Controls
 'i = i + 1
 'c.Enabled = en(i)
'Next
End Sub


Sub FadeOut(f As Form)
On Local Error Resume Next
Dim c
'For Each c In f.Controls
 'c.Enabled = False
'Next
Dim i As Long
    For i = 255 To 0 Step -3
        TransForm f.hwnd, CByte(i)
        DoEvents
        Call Sleep(FadeDelay)
    Next

End Sub

Public Function TransForm(fhWnd As Long, alpha As Byte) As Boolean
'Set alpha between 0-255
' 0 = Invisible , 128 = 50% transparent , 255 = Opaque
    SetWindowLong fhWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes fhWnd, 0, alpha, LWA_ALPHA
    LastAlpha = alpha
End Function





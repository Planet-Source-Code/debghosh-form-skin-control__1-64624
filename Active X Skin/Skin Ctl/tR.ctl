VERSION 5.00
Begin VB.UserControl tR 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   MouseIcon       =   "tR.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   24
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -420
      Top             =   -75
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   105
      Picture         =   "tR.ctx":0152
      Top             =   75
      Width           =   300
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   315
   End
End
Attribute VB_Name = "tR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Form Skin COntrol v 1.0
'Send Form To Tray Control
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event MouseUp()
Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.X, p.Y) = UserControl.hWnd Then
         MouseOver = True
    End If
End Function
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim hRgn As Long
    If Not MouseOver Then
        Shape1.FillColor = vbWhite
        Timer1.Enabled = False
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Height = 255
    UserControl.Width = 270
    UserControl.ScaleMode = vbPixels
    UserControl.AutoRedraw = True
    UserControl.Cls
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Image1.Move (UserControl.ScaleWidth - Image1.Width) / 2, (UserControl.ScaleHeight - Image1.Height) / 2
    Image1.ToolTipText = "Click Here For Menu"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Shape1.FillColor = &HE0E0E0
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
        Shape1.FillColor = &HFFC0C0
    Else
        Shape1.FillColor = vbWhite
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 255
    UserControl.Width = 270
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


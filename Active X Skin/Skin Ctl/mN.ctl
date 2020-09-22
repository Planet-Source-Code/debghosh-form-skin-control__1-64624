VERSION 5.00
Begin VB.UserControl mN 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   345
   MouseIcon       =   "mN.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   23
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -420
      Top             =   -75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   6
      X2              =   14
      Y1              =   13
      Y2              =   13
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   285
   End
End
Attribute VB_Name = "mN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Form Skin Control v 1.0
'Form Minimize Control
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click

Private Sub Timer1_Timer()
    If Not MouseOver Then
        Timer1.Enabled = False
        Call UCEnabled
    End If
End Sub
Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
    Shape1.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call UCEnabled
End Sub
Private Function MouseOver() As Boolean
    Dim p As POINTAPI
    Dim d As Long
    On Error Resume Next
    d = GetCursorPos(p)
    If WindowFromPoint(p.X, p.Y) = UserControl.hWnd Then
         MouseOver = True
    End If
End Function
Private Sub DefVal()
    UserControl.BackStyle = 0
    Shape1.FillStyle = 0
    Shape1.FillColor = vbWhite
    With Line1
        .X1 = 3
        .X2 = UserControl.ScaleWidth - 5
        .Y1 = UserControl.ScaleHeight / 2 + 3
        .Y2 = UserControl.ScaleHeight / 2 + 3
        .BorderColor = &HC000&
    End With
End Sub
Private Sub UcMouseOver()
    Shape1.FillColor = &H8000&
    Line1.BorderColor = vbWhite
End Sub

Private Sub UserControl_InitProperties()
    UserControl.ScaleMode = vbPixels
    UserControl.Height = 255
    UserControl.Width = 270
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Shape1.FillColor = &HE0E0E0
        Line1.BorderColor = &H808080
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X < 0 Or X > UserControl.ScaleWidth Or Y < 0 Or Y > UserControl.ScaleHeight Then
        Call UCEnabled
    Else
        Call UcMouseOver
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
    Exit Sub
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call UCEnabled
    PropertyChanged "Enabled"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Call UCEnabled
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub
Private Sub UCEnabled()
    If UserControl.Enabled = True Then
        With Line1
            .X1 = 3
            .X2 = UserControl.ScaleWidth - 5
            .Y1 = UserControl.ScaleHeight / 2 + 4
            .Y2 = UserControl.ScaleHeight / 2 + 4
            .BorderColor = &HC000&
        End With
        Shape1.FillStyle = 0
        Shape1.FillColor = vbWhite
    Else
        With Line1
            .X1 = 3
            .X2 = UserControl.ScaleWidth - 5
            .Y1 = UserControl.ScaleHeight / 2 + 3
            .Y2 = UserControl.ScaleHeight / 2 + 3
            .BorderColor = &HC0C0C0
        End With
        Shape1.FillStyle = 0
        Shape1.FillColor = &HE0E0E0
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub


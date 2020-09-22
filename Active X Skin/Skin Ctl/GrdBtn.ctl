VERSION 5.00
Begin VB.UserControl GrdBtn 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   PropertyPages   =   "GrdBtn.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   1245
   ToolboxBitmap   =   "GrdBtn.ctx":003F
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -500
      Top             =   -30
   End
   Begin VB.Image img 
      Height          =   225
      Left            =   240
      Top             =   90
      Width           =   240
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   870
      TabIndex        =   0
      Top             =   195
      Width           =   255
   End
End
Attribute VB_Name = "GrdBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Dim m_Red, m_d_Red, Rx, Rs, m_Green, m_d_Green, Gx, Gs, m_Blue, m_d_Blue, Bx, Bs, Rx1, Rs1, Gx1, Gs1, Bx1, Bs1, Y, X
Dim m_Dcolor, m_MDColor
Dim m_d_Fc As OLE_COLOR ' Default Label Fore Color
Dim m_Fc As OLE_COLOR 'Changed Label Fore Color

'Default Gradient
Public Enum DColor
    DRed = 1
    DGreen = 2
    DBlue = 3
End Enum
'Onmouse Move Gradient
Public Enum OnMouseMoveColor
    DMRed = 1
    DMGreen = 2
    DMBlue = 3
End Enum
'Mouse Pointer
Public Enum CmdMousePointer
    None = 0
    [Custom] = 99
End Enum
Dim hRgn As Long
Dim b As Integer
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseRightDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
' Check Whether Mouse On The Button Or Not
Private Function UMouseOver() As Boolean
    Dim typPoint As POINTAPI
    Dim dumpAway As Long
    On Error Resume Next
    dumpAway = GetCursorPos(typPoint)
    If WindowFromPoint(typPoint.X, typPoint.Y) = UserControl.hwnd Then
         UMouseOver = True
    End If
End Function
'Positioning Label And Image
Private Sub PicLabelPosition()
' Image And Label Position
If UserControl.Enabled = True Then
    img.Visible = True
    lblCaption.Enabled = True
    lblCaption.Top = (UserControl.ScaleHeight - lblCaption.Height) / 2
    If img.Picture <> 0 Then
        img.Top = (UserControl.ScaleHeight - img.Height) / 2
        If lblCaption.Caption <> "" Then
            lblCaption.Enabled = True
            img.Left = (UserControl.ScaleWidth - img.Width - lblCaption.Width - 5) / 2
        Else
            img.Left = (UserControl.ScaleWidth - img.Width) / 2
        End If
        lblCaption.Left = img.Left + img.Width + 5 '(UserControl.ScaleWidth - img.Left - img.Width - 5) / 2
    Else
        lblCaption.Left = (UserControl.ScaleWidth - lblCaption.Width) / 2
    End If
Else
    With img
        .Visible = False
    End With
    lblCaption.Enabled = False
    lblCaption.Top = (UserControl.ScaleHeight - lblCaption.Height) / 2
    lblCaption.Left = (UserControl.ScaleWidth - lblCaption.Width) / 2
End If
End Sub
Private Sub UCPaint()
'Paint User Control
    With UserControl
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Cls
    End With
    If UserControl.Enabled = True Then
    Select Case m_Dcolor
    
    Case 1
        'Paint User Control (RED)
        m_Red = 255: m_Green = 0: m_Blue = 0
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 2)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 2)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 2)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(160, 0, 0)), 1, 1
        SetWindowRgn UserControl.hwnd, hRgn, True
        DeleteObject hRgn
        
    Case 2
        'Paint User Control (GREEN)
        m_Red = 0: m_Green = 255: m_Blue = 0
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 2)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 2)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 2)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(0, 160, 0)), 1, 1
        SetWindowRgn UserControl.hwnd, hRgn, True
        DeleteObject hRgn
        
    Case 3
        'Paint User Control (BLUE)
        m_Red = 0: m_Green = 0: m_Blue = 255
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 1.8)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 1.8)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 1.8)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(0, 0, 160)), 1, 1
        SetWindowRgn UserControl.hwnd, hRgn, True
        DeleteObject hRgn
        
    End Select
    Call PicLabelPosition
    Else
        ' If Usercontrol Is Disabled Then Color Of Usercontrol
        m_Red = 255: m_Green = 255: m_Blue = 255
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 2)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 2)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 2)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(&H808080), 1, 1
        SetWindowRgn UserControl.hwnd, hRgn, True
        DeleteObject hRgn
        Call PicLabelPosition
    End If
End Sub
'Effect On Mouse Move
Private Sub UCMouseMove()
If UserControl.Enabled = True Then
    With UserControl
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Cls
    End With
    
    Select Case m_MDColor
        
        Case 1
        'Paint User Control (RED)
        m_Red = 255: m_Green = 0: m_Blue = 0
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 2)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 2)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 2)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(160, 0, 0)), 1, 1
        DeleteObject hRgn
    
    Case 2
        'Paint User Control (GREEN)
        m_Red = 0: m_Green = 255: m_Blue = 0
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 2)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 2)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 2)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(0, 160, 0)), 1, 1
        DeleteObject hRgn
        
    Case 3
        'Paint User Control (BLUE)
        m_Red = 0: m_Green = 0: m_Blue = 255
        Rx = m_Red
        Rs = ((m_Red - (m_Red / 1.8)) / (UserControl.ScaleHeight - 1))
        Gx = m_Green
        Gs = ((m_Green - (m_Green / 1.8)) / (UserControl.ScaleHeight - 1))
        Bx = m_Blue
        Bs = ((m_Blue - (m_Blue / 1.8)) / (UserControl.ScaleHeight - 1))
        For Y = 0 To UserControl.ScaleHeight
            UserControl.Line (0, Y)-(ScaleWidth, Y), RGB(Rx, Gx, Bx)
            Rx = Rx - Rs
            Gx = Gx - Gs
            Bx = Bx - Bs
        Next Y
        hRgn = CreateRoundRectRgn(0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 5, 5)
        FrameRgn UserControl.hdc, hRgn, CreateSolidBrush(RGB(0, 0, 160)), 1, 1
        DeleteObject hRgn
    End Select
    Call LabelChange
End If
End Sub
Private Sub DefLabelVal()
    'Default Label Color
    lblCaption.ForeColor = m_d_Fc
End Sub
Private Sub LabelChange()
    'Changed Label Color
    lblCaption.ForeColor = m_Fc
End Sub
Public Property Get DefaultGradient() As DColor
    'Default Gradient
    DefaultGradient = m_Dcolor
End Property
Public Property Let DefaultGradient(ByVal New_DColor As DColor)
    m_Dcolor = New_DColor
    Call UCPaint
    PropertyChanged "DefaultGradient"
End Property
Public Property Get OnMouseMoveGradient() As OnMouseMoveColor
    'Changed Gradient
    OnMouseMoveGradient = m_MDColor
End Property
Public Property Let OnMouseMoveGradient(ByVal New_COlor As OnMouseMoveColor)
    m_MDColor = New_COlor
    PropertyChanged "OnMouseMoveGradient"
End Property
Public Property Get OnMouseMoveForeColor() As OLE_COLOR
    'Changed Label Fore Color
    OnMouseMoveForeColor = m_Fc
End Property
Public Property Let OnMouseMoveForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_Fc = New_ForeColor
    PropertyChanged "OnMouseMoveForeColor"
End Property

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UserControl_MouseDown(Button, Shift, X, Y)
End If
End Sub
Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UserControl_MouseUp(Button, Shift, X, Y)
End If
End Sub
Private Sub lblCaption_DblClick()
    If UserControl.Enabled = True Then
        RaiseEvent DblClick
    End If
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UserControl_MouseDown(Button, Shift, X, Y)
End If
End Sub
Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UserControl_MouseUp(Button, Shift, X, Y)
End If
End Sub
Private Sub Timer1_Timer()
    If Not UMouseOver Then
        Call UCPaint
        Call DefLabelVal
        Timer1.Enabled = False
    End If
End Sub
Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    If UserControl.Enabled = True Then
        If (KeyAscii = 13 Or KeyAscii = 27) Then
            RaiseEvent Click
            Exit Sub
        End If
    End If
End Sub
Private Sub UserControl_GotFocus()
    lblCaption.FontUnderline = True
End Sub
Private Sub UserControl_Initialize()
    Call UCPaint
    UserControl.Refresh
End Sub
Private Sub UserControl_InitProperties()
    m_Dcolor = 2
    lblCaption.Caption = Extender.Name
End Sub
Private Sub UserControl_LostFocus()
    lblCaption.FontUnderline = False
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UCMouseMove
    For Y = 0 To ScaleHeight
        Line (0, Y)-(ScaleWidth, Y), RGB(Rx + 25, Gx + 25, Bx + 25)
        Rx = Rx + Rs
        Gx = Gx + Gs
        Bx = Bx + Bs
    Next Y
    img.Move img.Left + 1, img.Top + 1
    lblCaption.Move lblCaption.Left + 1, lblCaption.Top + 1
    RaiseEvent MouseDown(Button, Shift, X, Y)
End If
If Button = vbRightButton Then
    Call UCMouseMove
    For Y = 0 To ScaleHeight
        Line (0, Y)-(ScaleWidth, Y), RGB(Rx + 25, Gx + 25, Bx + 25)
        Rx = Rx + Rs
        Gx = Gx + Gs
        Bx = Bx + Bs
    Next Y
    img.Move img.Left + 1, img.Top + 1
    lblCaption.Move lblCaption.Left + 1, lblCaption.Top + 1
    RaiseEvent MouseRightDown(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X > 0 Or X < UserControl.ScaleWidth Or Y > 0 Or Y < UserControl.ScaleHeight Then
        Call UCMouseMove
    Else
        Timer1.Enabled = False
        Call DefLabelVal
        Call UCPaint
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call UCMouseMove
    Call PicLabelPosition
    RaiseEvent Click
    RaiseEvent MouseUp(Button, Shift, X, Y)
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Dcolor = PropBag.ReadProperty("DefaultGradient", 2)
    m_MDColor = PropBag.ReadProperty("OnMouseMoveGradient", 2)
    lblCaption.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", vbWhite)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 405)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 1440)
    m_d_Fc = lblCaption.ForeColor
    m_Fc = PropBag.ReadProperty("OnMouseMoveForeColor", m_d_Fc)
End Sub
Private Sub UserControl_Resize()
    If UserControl.Height < 180 Then
        UserControl.Height = 180
    End If
    If UserControl.Width < 180 Then
        UserControl.Width = 180
    End If
    Call UCPaint
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DefaultGradient", m_Dcolor, 2)
    Call PropBag.WriteProperty("OnMouseMoveGradient", m_MDColor, 2)
    Call PropBag.WriteProperty("Caption", lblCaption.Caption, UserControl.Name)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, vbWhite)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 405)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 1440)
    Call PropBag.WriteProperty("OnMouseMoveForeColor", m_Fc, m_d_Fc)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Caption
Public Property Get Caption() As String
    Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblCaption.Caption() = New_Caption
    Call UCPaint
    PropertyChanged "Caption"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Call UCPaint
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    Call UCPaint
    If lblCaption.FontUnderline = True Then
        lblCaption.FontUnderline = False
    End If
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    Call UCPaint
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            RaiseEvent Click
        Case vbKeySpace
            Call UCMouseMove
            For Y = 0 To ScaleHeight
                Line (0, Y)-(ScaleWidth, Y), RGB(Rx + 25, Gx + 25, Bx + 25)
                Rx = Rx + Rs
                Gx = Gx + Gs
                Bx = Bx + Bs
            Next Y
        Case 39, 40
            SendKeys "{Tab}"
        Case 37, 38
            SendKeys "+{Tab}"
    End Select
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeySpace) Then
        Call UCPaint
        Call DefLabelVal
        RaiseEvent Click
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As CmdMousePointer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As CmdMousePointer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=img,img,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = img.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set img.Picture = New_Picture
    Call UCPaint
    PropertyChanged "Picture"
End Property
Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub



VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.UserControl sK 
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   PropertyPages   =   "sK.ctx":0000
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "sK.ctx":002E
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -15000
      Top             =   -75
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3570
      Top             =   825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SkinCtl.tR tR1 
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin SkinCtl.mN mN1 
      Height          =   255
      Left            =   315
      TabIndex        =   2
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin SkinCtl.mX mX1 
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin SkinCtl.cL cL1 
      Height          =   255
      Left            =   900
      TabIndex        =   0
      Top             =   30
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   450
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   435
      Top             =   135
      Width           =   435
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1980
      TabIndex        =   4
      Top             =   105
      Width           =   735
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   930
      TabIndex        =   3
      Top             =   75
      Width           =   1200
   End
   Begin VB.Menu mnuChCol 
      Caption         =   "Choose Color"
      Begin VB.Menu mnuChCapCol 
         Caption         =   "Change Title Bar Color"
      End
      Begin VB.Menu mnuB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChFbCol 
         Caption         =   "Change Form Back Color"
      End
      Begin VB.Menu mnuB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChFmCapCol 
         Caption         =   "Change Caption Fore Color"
      End
      Begin VB.Menu mnuB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChFmCapShCol 
         Caption         =   "Change Caption Shade Color"
      End
   End
End
Attribute VB_Name = "sK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Form Skin Control v 1.0
'Enhanced Form Control
'Author :      Debasis Ghosh
'Email:        debughosh@vsnl.net
'         Copyright © 2004 by Debasis Ghosh
Option Explicit
Private WithEvents frm As Form
Attribute frm.VB_VarHelpID = -1

Private t_Icon As NOTIFYICONDATA
Private w_State As Long

Dim m_Uc As OLE_COLOR
Dim m_d_Uc As OLE_COLOR
Dim m_Fc As OLE_COLOR
Dim m_d_Fc As OLE_COLOR
Dim m_Ec As OLE_COLOR
Dim m_d_Ec As OLE_COLOR
Dim m_Cap As OLE_COLOR
Dim m_d_Cap As OLE_COLOR
Dim m_CapSh As OLE_COLOR
Dim m_d_CapSh As OLE_COLOR
Dim t_Info As String
Dim t_m_Info As String

Dim f_Max As Boolean
Dim f_Res As Boolean

Public Enum dlgShowActions
   fadeNone = 0
   FadeIn = 1
   FadeOut = 2
   fadeInOut = 3
End Enum

Private unloadAction As dlgShowActions
Private fadeMode As dlgShowActions
Private winstyle As Long

Private m_iOSver          As Byte         '/* OS 1=Win98/ME; 2=Win2000/XP

Private Sub Init()
    ' Initialize The Control And Parent
    Dim l As Long
    Dim fR As Long
    Dim r As RECT
    Dim fx As Integer, fy As Integer
    
    With frm
        .AutoRedraw = True
        .ScaleMode = vbPixels
    End With
    
    ' Get the current window style of the form.
    l = GetWindowLong(frm.hwnd, GWL_STYLE)
        
    ' Set the window style
    l = l And Not (WS_CAPTION)
    SetWindowLong frm.hwnd, GWL_STYLE, l
    
    ' Move And Size the Window
    fR = GetWindowRect(frm.hwnd, r)
    fx = r.Right - r.Left
    fy = r.Bottom - r.Top - GetSystemMetrics(SM_CYCAPTION)
    fR = MoveWindow(frm.hwnd, r.Left, r.Top, fx%, fy%, True)
    With Image1
        .Picture = frm.Icon
        .Stretch = True
        .Width = 18
        .Height = 18
    End With
    
    ' Form Caption
    With lblCaption
        .Caption = frm.Caption
    End With
    
    With lblS
        .Caption = lblCaption.Caption
        .Font = lblCaption.Font
        .FontBold = lblCaption.FontBold
        .FontItalic = lblCaption.FontItalic
        .FontUnderline = lblCaption.FontUnderline
        .FontSize = lblCaption.FontSize
    End With
    Call PositionControl
End Sub

Private Sub PositionControl()
        
    'Move image1 and Label
    Image1.Move 2, 2
    lblCaption.Move 0, 2, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblS.Move 1, 3, UserControl.ScaleWidth, UserControl.ScaleHeight
    lblCaption.ZOrder 0
    
    ' Checking  Form Borderstyle
    Select Case frm.BorderStyle
        Case 0
            cL1.Visible = False
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = False
            tR1.Move UserControl.ScaleWidth - tR1.Width - 4, 2
            
        Case 1
            cL1.Visible = True
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = True
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            tR1.Move cL1.Left - tR1.Width - 1, 2
            
        Case 2
            cL1.Visible = True
            mX1.Visible = True
            mN1.Visible = True
            tR1.Visible = True
            Image1.Visible = True
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            mX1.Move cL1.Left - mX1.Width - 1, 2
            mN1.Move mX1.Left - mN1.Width - 1, 2
            tR1.Move mN1.Left - tR1.Width - 1, 2
            
        Case 3
            cL1.Visible = True
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = True
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            tR1.Move cL1.Left - tR1.Width - 1, 2
            
        Case 4
            cL1.Visible = True
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = False
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            tR1.Move cL1.Left - tR1.Width - 1, 2
        
        Case 5
            cL1.Visible = True
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = False
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            tR1.Move cL1.Left - tR1.Width - 1, 2
            
        Case Else
            cL1.Visible = True
            mX1.Visible = False
            mN1.Visible = False
            tR1.Visible = True
            Image1.Visible = False
            cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
            tR1.Move cL1.Left - tR1.Width - 1, 2
    End Select
    
    'positioning form control box
    
End Sub
  Private Sub FormSysmenu()
    ' Sysmenu
    Dim cm As Long
    Dim hm As Long
    Dim lf As Long
    Dim r As RECT
    Dim ef As MenuControlConstants
    Dim pt As POINTAPI
    Dim l As Long
    
    GetCursorPos pt
    'Get System Menu
    hm = GetSystemMenu(frm.hwnd, &H0&)
    
    If hm <> 0 Then
    
        lf = ef Or (TPM_RETURNCMD)
        cm = TrackPopupMenu(hm, lf, pt.X, pt.Y, &H0&, frm.hwnd, r)
    End If
    
    If cm <> 0 Then
        Call PostMessage(frm.hwnd, WM_SYSCOMMAND, cm, hm)
    End If
 End Sub
 Private Sub FormCtlMenu()
    ' If Maximize is Enabled Or Disabled
    Dim hsMenu As Long
    Dim Cnt As Long
    If IsMaximizeEnable = False Then
        hsMenu = GetSystemMenu(frm.hwnd, False)
        If hsMenu Then
            Cnt = GetMenuItemCount(hsMenu)
                If Cnt Then
                    RemoveMenu hsMenu, Cnt - 3, MF_BYPOSITION Or MF_REMOVE  'Remove Maximize
                    RemoveMenu hsMenu, Cnt - 7, MF_BYPOSITION Or MF_REMOVE
                    DrawMenuBar frm.hwnd
                End If
        End If
    End If
 End Sub
Public Sub Paint_Form()
'Paint Form
On Error GoTo pError
    Dim rColor As Long
    Dim fColor As Long
    Dim fEcolor As Long
    TranslateColor CaptionColor, 0, rColor
    Dim frmWidth As Integer, frmHeight As Integer
    Dim X, Y
    Dim i As Integer
    Dim rndRG As Long
    Dim max As Long
    Dim l As Long
    Dim Rx As Integer
    Dim Gx As Integer
    Dim Bx As Integer
    Dim rxU, gxU, bxU As Integer
    rxU = CStr(rColor And &HFF&)
    gxU = CStr((rColor And &HFF00&) / 2 ^ 8)
    bxU = CStr((rColor And &HFF0000) / 2 ^ 16)
    
    'Paint Usercontrol Green
        For i = 0 To 24
            UserControl.Line (0, i)-(frm.ScaleWidth, i), RGB(rxU, gxU, bxU)
            If rxU > 0 Then
                rxU = rxU - 4
                    If rxU < 4 Then
                        rxU = 4
                    End If
            End If
            If gxU > 0 Then
                gxU = gxU - 4
                    If gxU < 4 Then
                        gxU = 4
                    End If
            End If
            If bxU > 0 Then
                bxU = bxU - 4
                    If bxU < 4 Then
                        bxU = 4
                    End If
            End If
        Next i
        
        TranslateColor FormBackColor, 0, fColor
        
        Rx = CStr(fColor And &HFF&)
        Gx = CStr((fColor And &HFF00&) / 2 ^ 8)
        Bx = CStr((fColor And &HFF0000) / 2 ^ 16)
        Rx = Rx - 60
        Gx = Gx - 60
        Bx = Bx - 60
        If Rx <= 0 Then
            Rx = 0
        End If
        If Gx <= 0 Then
            Gx = 0
        End If
        If Bx <= 0 Then
            Bx = 0
        End If
    
    frm.BackColor = FormBackColor
pError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error"
        Exit Sub
    End If
End Sub
Private Sub cL1_Click()
    On Error GoTo UnloadError
    Unload frm
UnloadError:
End Sub

Private Sub frm_Load()
    On Error GoTo fLoad
    Dim i As Integer
    
    Dim OSV        As OSVERSIONINFO
    
    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then m_iOSver = 1 '/* Win 98/ME
        If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then m_iOSver = 2  '/* Win 2000/XP
    End If
    
    'Initialize frm.hwnd
    Call Init ' Initialize Form And Control
    'Checking WindowState
    
    If m_iOSver = 2 Then
        Timer1.Enabled = True
        Call DialogAction(fadeInOut)
    End If
    
    If frm.WindowState = vbNormal Then
        mX1.frmState = False
        Call mX1.UcState
    Else
        mX1.frmState = True
        Call mX1.UcState
    End If
    
    mN1.ToolTipText = "Minimize"
    cL1.ToolTipText = "Close"
    
    Call FormCtlMenu
    
    Initialize frm.hwnd
    
fLoad:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
        Exit Sub
    End If
End Sub

Private Sub frm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ((UnloadMode = vbFormControlMenu) Or _
       (UnloadMode = vbFormCode)) And _
       (unloadAction = FadeOut) Or _
       (unloadAction = fadeInOut) Then

      Cancel = True

      fadeMode = FadeOut
      Timer1.Interval = 20
      Timer1.Enabled = True

   End If
End Sub

Private Sub frm_Resize()
    On Error GoTo fRError
    
    'Checking Window State
    If frm.WindowState = vbNormal Then
        mX1.frmState = False
        Call mX1.UcState
        mX1.ToolTipText = "Maximize"
    Else
        mX1.frmState = True
        Call mX1.UcState
        mX1.ToolTipText = "Restore Down"
    End If
    
    'Move Usercontrol To Top Of The Form
    MoveWindow UserControl.hwnd, 0, 0, frm.ScaleWidth, 24, 1
    
    
    
    Call PositionControl
    Call Paint_Form
    
    cL1.Refresh
    mN1.Refresh
    mX1.Refresh
    tR1.Refresh
    
    frm.Refresh
    
    
fRError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
        Exit Sub
    End If
End Sub

Private Sub frm_Unload(Cancel As Integer)
    Terminate frm.hwnd
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call FormSysmenu
    End If
End Sub
Private Sub lblCaption_DblClick()
    Call UserControl_DblClick
End Sub
Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lblS_DblClick()
    Call lblCaption_DblClick
End Sub

Private Sub lblS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblCaption_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblCaption_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub mN1_Click()
    On Error GoTo error_Min
    ShowWindow frm.hwnd, SW_MINIMIZE
error_Min:
   Exit Sub
End Sub

Private Sub mnuChCapCol_Click()
    On Error GoTo chColor
    Dim chC As Long
    cd.CancelError = True
    cd.ShowColor
    chC = cd.Color
    If chC <> 0 Then
        Me.CaptionColor = chC
        Call Paint_Form
    End If
chColor:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
        Exit Sub
    End If
End Sub

Private Sub mnuChFbCol_Click()
    On Error GoTo fbColor
    Dim fCo As Long
    cd.CancelError = True
    cd.ShowColor
    fCo = cd.Color
    If fCo <> 0 Then
        Me.FormBackColor = fCo
        Call Paint_Form
    End If
fbColor:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
    End If
End Sub

Private Sub mnuChFmCapCol_Click()
    On Error GoTo cUCol
    Dim Cu As Long
    cd.CancelError = True
    cd.ShowColor
    Cu = cd.Color
    If Cu <> 0 Then
        Me.ForeColor = Cu
    End If
cUCol:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
    End If
End Sub

Private Sub mnuChFmCapShCol_Click()
    On Error GoTo fCap
    Dim fC As Long
    cd.CancelError = True
    cd.ShowColor
    fC = cd.Color
    If fC <> 0 Then
        Me.Shade = fC
    End If
fCap:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, frm.Caption
    End If
End Sub

Private Sub mX1_Click()
    On Error GoTo maxError
    'On Error Resume Next
    If frm.WindowState = vbNormal Then
        mX1.frmState = False
        
        'ShowWindow frm.hWnd, SW_MAXIMIZE
        frm.WindowState = vbMaximized
        Call mX1.UcState
    Else
        mX1.frmState = True
        
        'ShowWindow frm.hWnd, SW_RESTORE
        frm.WindowState = vbNormal
        Call mX1.UcState
    End If

maxError:
    Exit Sub
End Sub
Public Property Get CaptionColor() As OLE_COLOR
    CaptionColor = m_Uc
End Property

Public Property Let CaptionColor(ByVal New_COlor As OLE_COLOR)
    m_Uc = New_COlor
    PropertyChanged "CaptionColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblCaption.ForeColor
End Property
Public Property Get FormBackColor() As OLE_COLOR
    FormBackColor = m_Fc
End Property

Public Property Let FormBackColor(ByVal New_COlor As OLE_COLOR)
    m_Fc = New_COlor
    PropertyChanged "FormBackColor"
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCaption,lblCaption,-1,Font
Public Property Get Font() As Font
    Set Font = lblCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblCaption.Font = New_Font
    If lblCaption.FontSize > 11 Then
        lblCaption.FontSize = 11
    End If
    'If lblCaption.FontItalic = True Then
        'lblCaption.FontItalic = False
    'End If
    If lblCaption.FontStrikethru = True Then
        lblCaption.FontStrikethru = False
    End If
    'If lblCaption.FontUnderline = True Then
        'lblCaption.FontUnderline = False
    'End If
    PropertyChanged "Font"
End Property
Public Property Get Shade() As OLE_COLOR
    Shade = lblS.ForeColor
End Property
Public Property Let Shade(ByVal New_Shade As OLE_COLOR)
    lblS.ForeColor = New_Shade
    PropertyChanged "Shade"
End Property
Public Property Get IsMaximizeEnable() As Boolean
    IsMaximizeEnable = mX1.Enabled
End Property
Public Property Let IsMaximizeEnable(ByVal New_Enable As Boolean)
    mX1.Enabled = New_Enable
    'Call FormCtlMenu
    PropertyChanged "IsMaximizeEnable"
End Property

Private Sub Timer1_Timer()
    Static fadeValue As Long
    Dim alpha As Long
   
    Select Case fadeMode
   
        Case FadeOut:
        unloadAction = 0
      
         If (fadeValue + (256 * 0.05)) >= 256 Then
             Timer1.Enabled = False
             fadeValue = 0
             Unload frm
             Exit Sub
         End If
         
         fadeValue = fadeValue + (256 * 0.05)
         alpha = (256 - fadeValue)
      
      Case FadeIn:
      
         If (fadeValue + (256 * 0.05)) >= 256 Then
             Timer1.Enabled = False
             fadeValue = 0
             alpha = 255
         Else
            fadeValue = fadeValue + (256 * 0.05)
            alpha = fadeValue
         End If
      
      Case Else
   End Select

   SetLayeredWindowAttributes frm.hwnd, 0&, alpha, LWA_ALPHA

End Sub

Private Sub tR1_Click()
    PopupMenu mnuChCol, , tR1.Left
End Sub

Private Sub UserControl_DblClick()
On Error Resume Next
    If IsMaximizeEnable = True Then
        If frm.BorderStyle = 2 Then
            If frm.WindowState = vbNormal Then
                ShowWindow frm.hwnd, SW_MAXIMIZE
            Else
                ShowWindow frm.hwnd, SW_RESTORE
            End If
        End If
    End If
End Sub
Private Sub UserControl_InitProperties()
    With UserControl
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .Cls
    End With
    Image1.Move 0, 2
    lblCaption.Move 0, 3, UserControl.ScaleWidth, UserControl.ScaleHeight
    cL1.Move UserControl.ScaleWidth - cL1.Width - 4, 2
    mX1.Move cL1.Left - mX1.Width - 1, 2
    mN1.Move mX1.Left - mN1.Width - 1, 2
    tR1.Move mN1.Left - tR1.Width - 1, 2
    CaptionColor = &HFFC0C0
    m_Ec = &HFFC0C0
    m_Fc = &HFFC0C0
    lblCaption.ForeColor = vbWhite
    lblS.ForeColor = vbBlack
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        ReleaseCapture
        SendMessage frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show Form Sysmenu
    If Button = vbRightButton Then
        Call FormSysmenu
    End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode = True Then
        Set frm = Parent
    End If
    m_Uc = PropBag.ReadProperty("CaptionColor", m_d_Uc)
    m_Fc = PropBag.ReadProperty("FormBackColor", m_d_Fc)
    Set lblCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    lblS.ForeColor = PropBag.ReadProperty("Shade", vbBlack)
    mX1.Enabled = PropBag.ReadProperty("IsMaximizeEnable", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("CaptionColor", m_Uc, m_d_Uc)
    Call PropBag.WriteProperty("FormBackColor", m_Fc, m_d_Fc)
    Call PropBag.WriteProperty("ForeColor", lblCaption.ForeColor, vbWhite)
    Call PropBag.WriteProperty("Shade", lblS.ForeColor, vbBlack)
    Call PropBag.WriteProperty("Font", lblCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsMaximizeEnable", mX1.Enabled, True)
End Sub
'Sub SendToTrayIcon()
    'With t_Icon
        '.cbSize = Len(t_Icon)
        '.hWnd = frm.hWnd
        '.uID = vbNull
        '.uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        '.uCallbackMessage = WM_MOUSEMOVE
        '.hIcon = frm.Icon
        '.szTip = TrayMInfo '"Developed By Debasis Ghosh" & vbNullChar
        '.dwState = 0
        '.dwStateMask = 0
        '.szInfo = TrayInfo '"Developed By Debasis Ghosh, Form Skin Control v 1.8" & vbCrLf & "Click Here" & Chr(0)
        '.szInfoTitle = "" & frm.Caption & ""
        '.dwInfoFlags = NIIF_INFO
        '.uTimeout = 3000
   'End With
        'Shell_NotifyIcon NIM_ADD, t_Icon 'Add Icon To Systray
        'w_State = frm.WindowState ' Hold Window State
        'frm.Hide
'End Sub
Sub ShowAboutBox()
Attribute ShowAboutBox.VB_UserMemId = -552
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub
Private Function AdjustWindowStyle()

   Dim style As Long

  'in order to have transparent windows, the
  'WS_EX_LAYERED window style must be applied
  'to the form
   style = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
   
   If Not (style And WS_EX_LAYERED = WS_EX_LAYERED) Then
      
      style = style Or WS_EX_LAYERED
      SetWindowLong frm.hwnd, GWL_EXSTYLE, style
      
   End If
    
End Function

Public Sub DialogAction(dlgEffectsMethod As dlgShowActions)
   
   Dim alpha As Long
 
   Select Case dlgEffectsMethod
   
      Case fadeNone
         Exit Sub
      
      Case FadeOut
         unloadAction = dlgEffectsMethod
         Call AdjustWindowStyle
         alpha = 255
         SetLayeredWindowAttributes frm.hwnd, 0&, alpha, LWA_ALPHA
      
      Case FadeIn, fadeInOut
         Call AdjustWindowStyle
         fadeMode = FadeIn
         Timer1.Interval = 20
         Timer1.Enabled = True
        
         If dlgEffectsMethod = fadeInOut Then unloadAction = FadeOut
         
   End Select
   
End Sub

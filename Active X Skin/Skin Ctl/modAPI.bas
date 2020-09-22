Attribute VB_Name = "modAPI"
'API Declaration
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Public Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Public Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

Public Const NIM_SETFOCUS = &H3
Public Const NIM_SETVERSION = &H4

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const NIF_STATE = &H8
Public Const NIF_INFO = &H10
 
Public Const NIS_HIDDEN = &H1
Public Const NIS_SHAREDICON = &H2
 
Public Const NIIF_NONE = &H0
Public Const NIIF_WARNING = &H2
Public Const NIIF_ERROR = &H3
Public Const NIIF_INFO = &H1
Public Const NIIF_GUID = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_NCACTIVATE As Long = &H86
Public Const WM_ACTIVATE As Long = &H6

Public Const SW_NORMAL = 1
Public Const SW_SHOWNORMAL = 1
Public Const SW_RESTORE As Long = 9

Public Const SW_MAXIMIZE As Long = 3
Public Const SW_MINIMIZE As Long = 6
Public Const SW_SHOWMAXIMIZED As Long = 3
Public Const SW_SHOWMINIMIZED As Long = 2

Public Const RGN_AND As Long = 1
Public Const RGN_COPY As Long = 5
Public Const RGN_DIFF As Long = 4
Public Const RGN_MAX As Long = RGN_COPY
Public Const RGN_MIN As Long = RGN_AND
Public Const RGN_OR As Long = 2
Public Const RGN_XOR As Long = 3

'Global frmState As Boolean
Public Const SM_CXFRAME As Long = 32
Public Const SM_CXSIZE As Long = 30
Public Const SM_CYBORDER As Long = 6
Public Const SM_CYCAPTION As Long = 4
Public Const SM_CYFRAME As Long = 33
Public Const SRCCOPY As Long = &HCC0020

Public Const HWND_TOP As Long = 0
Public Const HWND_TOPMOST As Long = -1
Public Const HWND_NOTOPMOST As Long = -2
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_SHOWWINDOW As Long = &H40

' Menu flags for Add/Check/EnableMenuItem().
Public Const MF_INSERT = &H0&
Public Const MF_CHANGE = &H80&
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_REMOVE = &H1000&

Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

Public Const MF_SEPARATOR = &H800&

Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_DISABLED = &H2&
Public Const MF_DEFAULT As Long = &H1000&

Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_USECHECKBITMAPS = &H200&

Public Const MF_STRING = &H0&
Public Const MF_BITMAP = &H4&
Public Const MF_OWNERDRAW = &H100&

Public Const MF_POPUP = &H10&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&

Public Const MF_UNHILITE = &H0&
Public Const MF_HILITE = &H80&

Public Const MF_SYSMENU = &H2000&
Public Const MF_HELP = &H4000&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_LINKS As Long = &H20000000
Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_STYLE As Long = -16
Public Const GWL_WNDPROC As Long = -4
Public Const SC_CLOSE As Long = &HF060&
Public Const SC_DEFAULT As Long = &HF160
Public Const SC_MAXIMIZE As Long = &HF030
Public Const SC_MINIMIZE As Long = &HF020
Public Const WS_MAXIMIZE As Long = &H1000000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const WS_MINIMIZE As Long = &H20000000
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_POPUP As Long = &H80000000
Public Const WS_SYSMENU As Long = &H80000
Public Const WS_THICKFRAME As Long = &H40000

Public Const TPM_BOTTOMALIGN As Long = &H20&
Public Const TPM_CENTERALIGN As Long = &H4&
Public Const TPM_HORIZONTAL As Long = &H0&
Public Const TPM_HORNEGANIMATION As Long = &H800&
Public Const TPM_HORPOSANIMATION As Long = &H400&
Public Const TPM_LEFTALIGN As Long = &H0&
Public Const TPM_LEFTBUTTON As Long = &H0&
Public Const TPM_NOANIMATION As Long = &H4000&
Public Const TPM_NONOTIFY As Long = &H80&
Public Const TPM_RECURSE As Long = &H1&
Public Const TPM_RETURNCMD As Long = &H100&
Public Const TPM_RIGHTALIGN As Long = &H8&
Public Const TPM_RIGHTBUTTON As Long = &H2&
Public Const TPM_TOPALIGN As Long = &H0&
Public Const TPM_VCENTERALIGN As Long = &H10&
Public Const TPM_VERNEGANIMATION As Long = &H2000&
Public Const TPM_VERPOSANIMATION As Long = &H1000&
Public Const TPM_VERTICAL As Long = &H40&

Public Const SW_SHOW As Long = 5
Public Const SW_SHOWDEFAULT As Long = 10

Public Const WM_SYSCOMMAND As Long = &H112
Public Const WM_SYSDEADCHAR As Long = &H107

Public Const WS_BORDER As Long = &H800000
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_DLGFRAME As Long = &H400000

Public Const WS_ACTIVECAPTION As Long = &H1
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_POPUPWINDOW As Long = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX As Long = WS_THICKFRAME

Public Const WM_NCLBUTTONDBLCLK As Long = &HA3
Public Const WM_NCLBUTTONDOWN As Long = &HA1
Public Const WM_NCLBUTTONUP As Long = &HA2

Public Const HTBORDER As Long = 18
Public Const HTCAPTION As Long = 2
Public Const HTCLIENT As Long = 1
Public Const HTCLOSE As Long = 20

Public Declare Function SetForegroundWindow Lib "user32" _
(ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Public Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Public Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Public Declare Function GetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Public Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lprc As RECT) As Long

Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function ReleaseCapture Lib "user32.dll" () As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Global NoFocus As Boolean
Global HasFocus As Boolean

'Public Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Public Const LWA_ALPHA As Long = &H2
Public Const LWA_COLORKEY As Long = &H1

Public Type OSVERSIONINFO
    OSVSize       As Long
    dwVerMajor    As Long
    dwVerMinor    As Long
    dwBuildNumber As Long
    PlatformID    As Long
    szCSDVersion  As String * 128
End Type


Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long


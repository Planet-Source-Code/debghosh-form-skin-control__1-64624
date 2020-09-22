Attribute VB_Name = "modSubclass"
Option Explicit
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4&)
Private Const HWND_TOP As Long = 0
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim PrevWndProc As Long
Dim mnuHandle As Long
Dim nmenuHandle As Long
Public Sub Initialize(hWnd As Long)
    mnuHandle = GetSystemMenu(hWnd, False)
    ' Add menu
    'Call FormOnTop(hWnd)
    Call AppendMenu(mnuHandle, MF_SEPARATOR, 0, "")
    Call AppendMenu(mnuHandle, MF_STRING, &H200, "Always On Top")
    Call AppendMenu(mnuHandle, MF_SEPARATOR, 0, "")
    Call AppendMenu(mnuHandle, MF_STRING, &H201, "About.....")
    PrevWndProc = SetWindowLong(mnuHandle, GWL_WNDPROC, AddressOf SubWndProc)
    PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SubWndProc)
End Sub
Public Sub Terminate(hWnd As Long)
    Call SetWindowLong(hWnd, GWL_WNDPROC, PrevWndProc)
End Sub
Public Function SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim Result As Long
    Select Case Msg
    Case WM_SYSCOMMAND
        Select Case wParam
            'Form Always On Top
            Case &H200
                Result = GetMenuState(mnuHandle, &H200, MF_BYCOMMAND)
                If Result And MF_CHECKED Then ' Checking Checked Menu
                    Call CheckMenuItem(mnuHandle, &H200, MF_BYCOMMAND Or MF_UNCHECKED)
                    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
                Else
                    Call CheckMenuItem(mnuHandle, &H200, MF_BYCOMMAND Or MF_CHECKED)
                    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
                End If
            Case &H201
                SetWindowPos frmAbout.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
        End Select
    End Select
    
    SubWndProc = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
End Function
Public Sub FormOnTop(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub



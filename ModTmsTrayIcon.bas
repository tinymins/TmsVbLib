Attribute VB_Name = "ModTmsTrayIcon"
Option Explicit
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONUP = &H202
Private Const WM_USER = &H400
Private Const WM_NOTIFYICON = WM_USER + 1
Private Const WM_LBUTTONDBLCLK = &H203
Private Const GWL_WNDPROC = (-4)
'0-9
Private Const WM_MOUSEUP = &H200
Private Const NIN_BALLOONSHOW = (WM_USER + &H2)
Private Const NIN_BALLOONHIDE = (WM_USER + &H3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + &H4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + &H5)
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeoutOrVersion As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIF_ICON As Long = &H2
Private Const NIF_INFO As Long = &H10
Private Const NIF_MESSAGE As Long = &H1
Private Const NIF_STATE As Long = &H8
Private Const NIF_TIP As Long = &H4
Private Const NIM_ADD As Long = &H0
Private Const NIM_DELETE As Long = &H2
Private Const NIM_MODIFY As Long = &H1
Private Const NIM_SETFOCUS As Long = &H3
Private Const lngNIM_SETVERSION As Long = &H4
Private lngPreWndProc As Long

Private TheForm As Form
Private TheMenu As Menu
Private TheData As NOTIFYICONDATA
'AddNotifyIcon Form1, "系统信息获取", 1  '添加系统托盘
Public Sub AddNotifyIcon(frm As Form, strTitle As String, Optional lngType As Long = 1, Optional lngTime As Long = 10000)
    strTitle = strTitle & vbNullChar
    Set TheForm = frm
    With TheData
        .cbSize = Len(TheData)
        .hwnd = TheForm.hwnd
        .uID = 0
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_STATE
        .uCallbackMessage = WM_NOTIFYICON
        .szTip = strTitle
        .hIcon = TheForm.Icon.Handle
        .dwState = 0
        .dwStateMask = 0
        .szInfo = vbNullChar
        .szInfoTitle = strTitle
        .dwInfoFlags = lngType
        .uTimeoutOrVersion = lngTime
    End With
    
    If lngPreWndProc = 0 Then
       Shell_NotifyIcon NIM_ADD, TheData
       lngPreWndProc = SetWindowLong(TheForm.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    Else
       Shell_NotifyIcon NIM_MODIFY, TheData
    End If
End Sub

Public Sub BindNotifyMenu(mnu As Menu)
    Set TheMenu = mnu
End Sub

Public Sub DelNotifyIcon()
    If lngPreWndProc <> 0 Then
        With TheData
            .cbSize = Len(TheData)
            .hwnd = TheForm.hwnd
            .uID = 0
            .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE
            .uCallbackMessage = WM_NOTIFYICON
            .szTip = ""
            .hIcon = TheForm.Icon.Handle
        End With
        Shell_NotifyIcon NIM_DELETE, TheData
        SetWindowLong TheForm.hwnd, GWL_WNDPROC, lngPreWndProc
        lngPreWndProc = 0
    End If
End Sub

Public Sub ChangeNotifyIcon(pic As Picture)
    If lngPreWndProc = 0 Then Exit Sub
    With TheData
        .cbSize = Len(TheData)
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_STATE
        .hIcon = pic.Handle
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Public Sub ShowNotifyTip(strInfo As String)
    If lngPreWndProc = 0 Then Exit Sub
    strInfo = strInfo & vbNullChar
    With TheData
        .cbSize = Len(TheData)
        .szInfo = strInfo
        .uFlags = NIF_TIP Or NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_STATE
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub

Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   'On Error Resume Next
    If Msg = WM_NOTIFYICON Then
        Select Case lParam
            Case WM_LBUTTONUP
                 '左键点击托盘图标响应事件
                 If TheForm.WindowState = vbMinimized Or TheForm.Visible = False Then
                    TheForm.Show
                    TheForm.WindowState = vbNormal
                 Else
                    TheForm.Hide
                    TheForm.WindowState = vbMinimized
                 End If
            Case WM_RBUTTONUP
                 '右键点击托盘图标响应事件
                 TheForm.PopupMenu TheMenu
                 
            Case WM_MOUSEUP
                 '鼠村移到托盘图标响应事件
               
            Case NIN_BALLOONSHOW
                Debug.Print "显示气球提示"
            Case NIN_BALLOONHIDE
                Debug.Print "删除托盘图标"
            Case NIN_BALLOONTIMEOUT
                Debug.Print "气球提示消失"
            Case NIN_BALLOONUSERCLICK
                Debug.Print "单击气球提示响应事件"
        End Select
    End If
    WindowProc = CallWindowProc(lngPreWndProc, hwnd, Msg, wParam, lParam)
End Function

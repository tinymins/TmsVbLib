Attribute VB_Name = "modWin32Constants"
'//
'// Win32 Constants
'//

'#region Peek Message Flags
Public Const PM_NOREMOVEConst = 0
Public Const PM_REMOVEConst = 1
Public Const PM_NOYIELDConst = 2
'#End Region

'#Region Windows Messages
Public Const WM_NULL = &H0
Public Const WM_CREATE = &H1
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_ENABLE = &HA
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_PAINT = &HF
Public Const WM_CLOSE = &H10
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUIT = &H12
Public Const WM_QUERYOPEN = &H13
Public Const WM_ERASEBKGND = &H14
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_ENDSESSION = &H16
Public Const WM_SHOWWINDOW = &H18
Public Const WM_CTLCOLOR = &H19
Public Const WM_WININICHANGE = &H1A
Public Const WM_SETTINGCHANGE = &H1A
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_FONTCHANGE = &H1D
Public Const WM_TIMECHANGE = &H1E
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_QUEUESYNC = &H23
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_PAINTICON = &H26
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_DRAWITEM = &H2B
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DELETEITEM = &H2D
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_CHARTOITEM = &H2F
Public Const WM_SETFONT = &H30
Public Const WM_GETFONT = &H31
Public Const WM_SETHOTKEY = &H32
Public Const WM_GETHOTKEY = &H33
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_COMPAREITEM = &H39
Public Const WM_GETOBJECT = &H3D
Public Const WM_COMPACTING = &H41
Public Const WM_COMMNOTIFY = &H44
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_POWER = &H48
Public Const WM_COPYDATA = &H4A
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_NOTIFY = &H4E
Public Const WM_INPUTLANGCHANGEREQUEST = &H50
Public Const WM_INPUTLANGCHANGE = &H51
Public Const WM_TCARD = &H52
Public Const WM_HELP = &H53
Public Const WM_USERCHANGED = &H54
Public Const WM_NOTIFYFORMAT = &H55
Public Const WM_CONTEXTMENU = &H7B
Public Const WM_STYLECHANGING = &H7C
Public Const WM_STYLECHANGED = &H7D
Public Const WM_DISPLAYCHANGE = &H7E
Public Const WM_GETICON = &H7F
Public Const WM_SETICON = &H80
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_GETDLGCODE = &H87
Public Const WM_SYNCPAINT = &H88
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_DEADCHAR = &H103
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_KEYLAST = &H108
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_TIMER = &H113
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_ENTERIDLE = &H121
Public Const WM_MENURBUTTONUP = &H122
Public Const WM_MENUDRAG = &H123
Public Const WM_MENUGETOBJECT = &H124
Public Const WM_UNINITMENUPOPUP = &H125
Public Const WM_MENUCOMMAND = &H126
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_NEXTMENU = &H213
Public Const WM_SIZING = &H214
Public Const WM_CAPTURECHANGED = &H215
Public Const WM_MOVING = &H216
Public Const WM_DEVICECHANGE = &H219
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_DROPFILES = &H233
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_REQUEST = &H288
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYUP = &H291
Public Const WM_MOUSEHOVER = &H2A1
Public Const WM_MOUSELEAVE = &H2A3
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const WM_CLEAR = &H303
Public Const WM_UNDO = &H304
Public Const WM_RENDERFORMAT = &H305
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PALETTECHANGED = &H311
Public Const WM_HOTKEY = &H312
Public Const WM_PRINT = &H317
Public Const WM_PRINTCLIENT = &H318
Public Const WM_HANDHELDFIRST = &H358
Public Const WM_HANDHELDLAST = &H35F
Public Const WM_AFXFIRST = &H360
Public Const WM_AFXLAST = &H37F
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_APP = &H8000
Public Const WM_USER = &H400
Public Const WM_REFLECT = WM_USER + &H1C00
'#End Region

'#Region Window Styles
Public Const WS_OVERLAPPED = &H0
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_DLGFRAME = &H400000
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_TILED = &H0
Public Const WS_ICONIC = &H20000000
Public Const WS_SIZEBOX = &H40000
Public Const WS_POPUPWINDOW = &H80880000
Public Const WS_OVERLAPPEDWINDOW = &HCF0000
Public Const WS_TILEDWINDOW = &HCF0000
Public Const WS_CHILDWINDOW = &H40000000
'#End Region

'#Region Window Extended Styles
Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_NOPARENTNOTIFY = &H4
Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_LEFT = &H0
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LTRREADING = &H0
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_RIGHTSCROLLBAR = &H0
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_EX_OVERLAPPEDWINDOW = &H300
Public Const WS_EX_PALETTEWINDOW = &H188
Public Const WS_EX_LAYERED = &H80000
'#End Region

'#Region ShowWindow Styles
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10
Public Const SW_FORCEMINIMIZE = 11
Public Const SW_MAX = 11
'#End Region

'#Region SetWindowPos Z Order
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'#End Region

'#Region SetWindowPosFlags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOSENDCHANGING = &H400
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_DEFERERASE = &H2000
Public Const SWP_ASYNCWINDOWPOS = &H4000
'#End Region

'#Region Virtual Keys
Public Const VK_LBUTTON = &H1
Public Const VK_CANCEL = &H3
Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD
Public Const VK_ENTER = &HD
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_CAPITAL = &H14
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_HELP = &H2F
Public Const VK_0 = &H30
Public Const VK_1 = &H31
Public Const VK_2 = &H32
Public Const VK_3 = &H33
Public Const VK_4 = &H34
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_7 = &H37
Public Const VK_8 = &H38
Public Const VK_9 = &H39
Public Const VK_A = &H41
Public Const VK_B = &H42
Public Const VK_C = &H43
Public Const VK_D = &H44
Public Const VK_E = &H45
Public Const VK_F = &H46
Public Const VK_G = &H47
Public Const VK_H = &H48
Public Const VK_I = &H49
Public Const VK_J = &H4A
Public Const VK_K = &H4B
Public Const VK_L = &H4C
Public Const VK_M = &H4D
Public Const VK_N = &H4E
Public Const VK_O = &H4F
Public Const VK_P = &H50
Public Const VK_Q = &H51
Public Const VK_R = &H52
Public Const VK_S = &H53
Public Const VK_T = &H54
Public Const VK_U = &H55
Public Const VK_V = &H56
Public Const VK_W = &H57
Public Const VK_X = &H58
Public Const VK_Y = &H59
Public Const VK_Z = &H5A
Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE
Public Const VK_LWIN = &H5B
Public Const VK_RWIN = &H5C
Public Const VK_APPS = &H5D
Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5
'#End Region

'#Region PatBlt Types
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const NOTSRCCOPY = &H330008
Public Const NOTSRCERASE = &H1100A6
Public Const MERGECOPY = &HC000CA
Public Const MERGEPAINT = &HBB0226
Public Const PATCOPY = &HF00021
Public Const PATPAINT = &HFB0A09
Public Const PATINVERT = &H5A0049
Public Const DSTINVERT = &H550009
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
'#End Region

'#Region Clipboard Formats
Public Const CF_TEXT = 1
Public Const CF_BITMAP = 2
Public Const CF_METAFILEPICT = 3
Public Const CF_SYLK = 4
Public Const CF_DIF = 5
Public Const CF_TIFF = 6
Public Const CF_OEMTEXT = 7
Public Const CF_DIB = 8
Public Const CF_PALETTE = 9
Public Const CF_PENDATA = 10
Public Const CF_RIFF = 11
Public Const CF_WAVE = 12
Public Const CF_UNICODETEXT = 13
Public Const CF_ENHMETAFILE = 14
Public Const CF_HDROP = 15
Public Const CF_LOCALE = 16
Public Const CF_MAX = 17
Public Const CF_OWNERDISPLAY = &H80
Public Const CF_DSPTEXT = &H81
Public Const CF_DSPBITMAP = &H82
Public Const CF_DSPMETAFILEPICT = &H83
Public Const CF_DSPENHMETAFILE = &H8E
Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF
Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF
'#End Region

'#Region Common Controls Initialization flags
Public Const ICC_LISTVIEW_CLASSES = &H1
Public Const ICC_TREEVIEW_CLASSES = &H2
Public Const ICC_BAR_CLASSES = &H4
Public Const ICC_TAB_CLASSES = &H8
Public Const ICC_UPDOWN_CLASS = &H10
Public Const ICC_PROGRESS_CLASS = &H20
Public Const ICC_HOTKEY_CLASS = &H40
Public Const ICC_ANIMATE_CLASS = &H80
Public Const ICC_WIN95_CLASSES = &HFF
Public Const ICC_DATE_CLASSES = &H100
Public Const ICC_USEREX_CLASSES = &H200
Public Const ICC_COOL_CLASSES = &H400
Public Const ICC_INTERNET_CLASSES = &H800
Public Const ICC_PAGESCROLLER_CLASS = &H1000
Public Const ICC_NATIVEFNTCTL_CLASS = &H2000
'#End Region

'#Region Common Controls Styles
Public Const CCS_TOP = &H1
Public Const CCS_NOMOVEY = &H2
Public Const CCS_BOTTOM = &H3
Public Const CCS_NORESIZE = &H4
Public Const CCS_NOPARENTALIGN = &H8
Public Const CCS_ADJUSTABLE = &H20
Public Const CCS_NODIVIDER = &H40
Public Const CCS_VERT = &H80
Public Const CCS_LEFT = (CCS_VERT Or CCS_TOP)
Public Const CCS_RIGHT = (CCS_VERT Or CCS_BOTTOM)
Public Const CCS_NOMOVEX = (CCS_VERT Or CCS_NOMOVEY)
'#End Region

'#Region Toolbar button styles
Public Const TBSTYLE_BUTTON = &H0
Public Const TBSTYLE_SEP = &H1
Public Const TBSTYLE_CHECK = &H2
Public Const TBSTYLE_GROUP = &H4
Public Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
Public Const TBSTYLE_DROPDOWN = &H8
Public Const TBSTYLE_AUTOSIZE = &H10
Public Const TBSTYLE_NOPREFIX = &H20
Public Const TBSTYLE_TOOLTIPS = &H100
Public Const TBSTYLE_WRAPABLE = &H200
Public Const TBSTYLE_ALTDRAG = &H400
Public Const TBSTYLE_FLAT = &H800
Public Const TBSTYLE_LIST = &H1000
Public Const TBSTYLE_CUSTOMERASE = &H2000
Public Const TBSTYLE_REGISTERDROP = &H4000
Public Const TBSTYLE_TRANSPARENT = &H8000
Public Const TBSTYLE_DRAWDDARROWS = &H1
'#End Region

'#Region ToolBar Ex Styles
Public Const TBSTYLE_EX_DRAWDDARROWS = &H1
Public Const TBSTYLE_EX_HIDECLIPPEDBUTTONS = &H10
Public Const TBSTYLE_EX_DOUBLEBUFFER = &H80
'#End Region

'#Region ToolBar Messages
Public Const TB_ENABLEBUTTON = (WM_USER + 1)
Public Const TB_CHECKBUTTON = (WM_USER + 2)
Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_HIDEBUTTON = (WM_USER + 4)
Public Const TB_INDETERMINATE = (WM_USER + 5)
Public Const TB_MARKBUTTON = (WM_USER + 6)
Public Const TB_ISBUTTONENABLED = (WM_USER + 9)
Public Const TB_ISBUTTONCHECKED = (WM_USER + 10)
Public Const TB_ISBUTTONPRESSED = (WM_USER + 11)
Public Const TB_ISBUTTONHIDDEN = (WM_USER + 12)
Public Const TB_ISBUTTONINDETERMINATE = (WM_USER + 13)
Public Const TB_ISBUTTONHIGHLIGHTED = (WM_USER + 14)
Public Const TB_SETSTATE = (WM_USER + 17)
Public Const TB_GETSTATE = (WM_USER + 18)
Public Const TB_ADDBITMAP = (WM_USER + 19)
Public Const TB_ADDBUTTONSA = (WM_USER + 20)
Public Const TB_INSERTBUTTONA = (WM_USER + 21)
Public Const TB_ADDBUTTONS = (WM_USER + 20)
Public Const TB_INSERTBUTTON = (WM_USER + 21)
Public Const TB_DELETEBUTTON = (WM_USER + 22)
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_COMMANDTOINDEX = (WM_USER + 25)
Public Const TB_SAVERESTOREA = (WM_USER + 26)
Public Const TB_CUSTOMIZE = (WM_USER + 27)
Public Const TB_ADDSTRINGA = (WM_USER + 28)
Public Const TB_GETITEMRECT = (WM_USER + 29)
Public Const TB_BUTTONSTRUCTSIZE = (WM_USER + 30)
Public Const TB_SETBUTTONSIZE = (WM_USER + 31)
Public Const TB_SETBITMAPSIZE = (WM_USER + 32)
Public Const TB_AUTOSIZE = (WM_USER + 33)
Public Const TB_GETTOOLTIPS = (WM_USER + 35)
Public Const TB_SETTOOLTIPS = (WM_USER + 36)
Public Const TB_SETPARENT = (WM_USER + 37)
Public Const TB_SETROWS = (WM_USER + 39)
Public Const TB_GETROWS = (WM_USER + 40)
Public Const TB_GETBITMAPFLAGS = (WM_USER + 41)
Public Const TB_SETCMDID = (WM_USER + 42)
Public Const TB_CHANGEBITMAP = (WM_USER + 43)
Public Const TB_GETBITMAP = (WM_USER + 44)
Public Const TB_GETBUTTONTEXTA = (WM_USER + 45)
Public Const TB_GETBUTTONTEXTW = (WM_USER + 75)
Public Const TB_REPLACEBITMAP = (WM_USER + 46)
Public Const TB_SETINDENT = (WM_USER + 47)
Public Const TB_SETIMAGELIST = (WM_USER + 48)
Public Const TB_GETIMAGELIST = (WM_USER + 49)
Public Const TB_LOADIMAGES = (WM_USER + 50)
Public Const TB_GETRECT = (WM_USER + 51)
Public Const TB_SETHOTIMAGELIST = (WM_USER + 52)
Public Const TB_GETHOTIMAGELIST = (WM_USER + 53)
Public Const TB_SETDISABLEDIMAGELIST = (WM_USER + 54)
Public Const TB_GETDISABLEDIMAGELIST = (WM_USER + 55)
Public Const TB_SETSTYLE = (WM_USER + 56)
Public Const TB_GETSTYLE = (WM_USER + 57)
Public Const TB_GETBUTTONSIZE = (WM_USER + 58)
Public Const TB_SETBUTTONWIDTH = (WM_USER + 59)
Public Const TB_SETMAXTEXTROWS = (WM_USER + 60)
Public Const TB_GETTEXTROWS = (WM_USER + 61)
Public Const TB_GETOBJECT = (WM_USER + 62)
Public Const TB_GETBUTTONINFOW = (WM_USER + 63)
Public Const TB_SETBUTTONINFOW = (WM_USER + 64)
Public Const TB_GETBUTTONINFOA = (WM_USER + 65)
Public Const TB_SETBUTTONINFOA = (WM_USER + 66)
Public Const TB_INSERTBUTTONW = (WM_USER + 67)
Public Const TB_ADDBUTTONSW = (WM_USER + 68)
Public Const TB_HITTEST = (WM_USER + 69)
Public Const TB_SETDRAWTEXTFLAGS = (WM_USER + 70)
Public Const TB_GETHOTITEM = (WM_USER + 71)
Public Const TB_SETHOTITEM = (WM_USER + 72)
Public Const TB_SETANCHORHIGHLIGHT = (WM_USER + 73)
Public Const TB_GETANCHORHIGHLIGHT = (WM_USER + 74)
Public Const TB_SAVERESTOREW = (WM_USER + 76)
Public Const TB_ADDSTRINGW = (WM_USER + 77)
Public Const TB_MAPACCELERATORA = (WM_USER + 78)
Public Const TB_GETINSERTMARK = (WM_USER + 79)
Public Const TB_SETINSERTMARK = (WM_USER + 80)
Public Const TB_INSERTMARKHITTEST = (WM_USER + 81)
Public Const TB_MOVEBUTTON = (WM_USER + 82)
Public Const TB_GETMAXSIZE = (WM_USER + 83)
Public Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)
Public Const TB_GETEXTENDEDSTYLE = (WM_USER + 85)
Public Const TB_GETPADDING = (WM_USER + 86)
Public Const TB_SETPADDING = (WM_USER + 87)
Public Const TB_SETINSERTMARKCOLOR = (WM_USER + 88)
Public Const TB_GETINSERTMARKCOLOR = (WM_USER + 89)
'#End Region

'#Region ToolBar Notifications
Public Const TTN_NEEDTEXTA = ((0 - 520) - 0)
Public Const TTN_NEEDTEXTW = ((0 - 520) - 10)
Public Const TBN_QUERYINSERT = ((0 - 700) - 6)
Public Const TBN_DROPDOWN = ((0 - 700) - 10)
Public Const TBN_HOTITEMCHANGE = ((0 - 700) - 13)
'#End Region

'#Region Reflected Messages
Public Const OCM__BASE = (WM_USER + &H1C00)
Public Const OCM_COMMAND = (OCM__BASE + WM_COMMAND)
Public Const OCM_CTLCOLORBTN = (OCM__BASE + WM_CTLCOLORBTN)
Public Const OCM_CTLCOLOREDIT = (OCM__BASE + WM_CTLCOLOREDIT)
Public Const OCM_CTLCOLORDLG = (OCM__BASE + WM_CTLCOLORDLG)
Public Const OCM_CTLCOLORLISTBOX = (OCM__BASE + WM_CTLCOLORLISTBOX)
Public Const OCM_CTLCOLORMSGBOX = (OCM__BASE + WM_CTLCOLORMSGBOX)
Public Const OCM_CTLCOLORSCROLLBAR = (OCM__BASE + WM_CTLCOLORSCROLLBAR)
Public Const OCM_CTLCOLORSTATIC = (OCM__BASE + WM_CTLCOLORSTATIC)
Public Const OCM_CTLCOLOR = (OCM__BASE + WM_CTLCOLOR)
Public Const OCM_DRAWITEM = (OCM__BASE + WM_DRAWITEM)
Public Const OCM_MEASUREITEM = (OCM__BASE + WM_MEASUREITEM)
Public Const OCM_DELETEITEM = (OCM__BASE + WM_DELETEITEM)
Public Const OCM_VKEYTOITEM = (OCM__BASE + WM_VKEYTOITEM)
Public Const OCM_CHARTOITEM = (OCM__BASE + WM_CHARTOITEM)
Public Const OCM_COMPAREITEM = (OCM__BASE + WM_COMPAREITEM)
Public Const OCM_HSCROLL = (OCM__BASE + WM_HSCROLL)
Public Const OCM_VSCROLL = (OCM__BASE + WM_VSCROLL)
Public Const OCM_PARENTNOTIFY = (OCM__BASE + WM_PARENTNOTIFY)
Public Const OCM_NOTIFY = (OCM__BASE + WM_NOTIFY)
'#End Region

'#Region Notification Messages
Public Const NM_FIRST = (0 - 0)
Public Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Public Const NM_NCHITTEST = (NM_FIRST - 14)
'#End Region

'#Region ToolTip Flags
Public Const TTF_CENTERTIP = &H2
Public Const TTF_RTLREADING = &H4
Public Const TTF_SUBCLASS = &H10
Public Const TTF_TRACK = &H20
Public Const TTF_ABSOLUTE = &H80
Public Const TTF_TRANSPARENT = &H100
Public Const TTF_DI_SETITEM = &H8000
'#End Region

'#Region Custom Draw Return Flags
Public Const CDRF_DODEFAULT = &H0
Public Const CDRF_NEWFONT = &H2
Public Const CDRF_SKIPDEFAULT = &H4
Public Const CDRF_NOTIFYPOSTPAINT = &H10
Public Const CDRF_NOTIFYITEMDRAW = &H20
Public Const CDRF_NOTIFYSUBITEMDRAW = &H20
Public Const CDRF_NOTIFYPOSTERASE = &H40
'#End Region

'#Region Custom Draw Item State Flags
Public Const CDIS_SELECTED = &H1
Public Const CDIS_GRAYED = &H2
Public Const CDIS_DISABLED = &H4
Public Const CDIS_CHECKED = &H8
Public Const CDIS_FOCUS = &H10
Public Const CDIS_DEFAULT = &H20
Public Const CDIS_HOT = &H40
Public Const CDIS_MARKED = &H80
Public Const CDIS_INDETERMINATE = &H100
'#End Region

'#Region Custom Draw Draw State Flags
Public Const CDDS_PREPAINT = &H1
Public Const CDDS_POSTPAINT = &H2
Public Const CDDS_PREERASE = &H3
Public Const CDDS_POSTERASE = &H4
Public Const CDDS_ITEM = &H10000
Public Const CDDS_ITEMPREPAINT = (CDDS_ITEM Or CDDS_PREPAINT)
Public Const CDDS_ITEMPOSTPAINT = (CDDS_ITEM Or CDDS_POSTPAINT)
Public Const CDDS_ITEMPREERASE = (CDDS_ITEM Or CDDS_PREERASE)
Public Const CDDS_ITEMPOSTERASE = (CDDS_ITEM Or CDDS_POSTERASE)
Public Const CDDS_SUBITEM = &H20000
'#End Region

'#Region Toolbar button info flags
Public Const TBIF_IMAGE = &H1
Public Const TBIF_TEXT = &H2
Public Const TBIF_STATE = &H4
Public Const TBIF_STYLE = &H8
Public Const TBIF_LPARAM = &H10
Public Const TBIF_COMMAND = &H20
Public Const TBIF_SIZE = &H40
Public Const I_IMAGECALLBACK = -1
Public Const I_IMAGENONE = -2
'#End Region

'#Region Toolbar button state
Public Const TBSTATE_CHECKED = &H1
Public Const TBSTATE_PRESSED = &H2
Public Const TBSTATE_ENABLED = &H4
Public Const TBSTATE_HIDDEN = &H8
Public Const TBSTATE_INDETERMINATE = &H10
Public Const TBSTATE_WRAP = &H20
Public Const TBSTATE_ELLIPSES = &H40
Public Const TBSTATE_MARKED = &H80
'#End Region

'#Region Windows Hook Codes
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_MOUSE = 7
Public Const WH_HARDWARE = 8
Public Const WH_DEBUG = 9
Public Const WH_SHELL = 10
Public Const WH_FOREGROUNDIDLE = 11
Public Const WH_CALLWNDPROCRET = 12
Public Const WH_KEYBOARD_LL = 13
Public Const WH_MOUSE_LL = 14
'#End Region

'#Region Hook Status
Public Const HC_ACTION = 0
Public Const HC_GETNEXT = 1
Public Const HC_SKIP = 2
Public Const HC_NOREMOVE = 3
Public Const HC_NOREM = HC_NOREMOVE
Public Const HC_SYSMODALON = 4
Public Const HC_SYSMODALOFF = 5
'#End Region

'#Region Mouse Hook Filters
Public Const MSGF_DIALOGBOX = 0
Public Const MSGF_MESSAGEBOX = 1
Public Const MSGF_MENU = 2
Public Const MSGF_SCROLLBAR = 5
Public Const MSGF_NEXTWINDOW = 6
'#End Region

'#Region Draw Text format flags
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000
Public Const DT_EDITCONTROL = &H2000
Public Const DT_PATH_ELLIPSIS = &H4000
Public Const DT_END_ELLIPSIS = &H8000
Public Const DT_MODIFYSTRING = &H10000
Public Const DT_RTLREADING = &H20000
Public Const DT_WORD_ELLIPSIS = &H40000
'#End Region

'#Region Rebar Styles
Public Const RBS_TOOLTIPS = &H100
Public Const RBS_VARHEIGHT = &H200
Public Const RBS_BANDBORDERS = &H400
Public Const RBS_FIXEDORDER = &H800
Public Const RBS_REGISTERDROP = &H1000
Public Const RBS_AUTOSIZE = &H2000
Public Const RBS_VERTICALGRIPPER = &H4000
Public Const RBS_DBLCLKTOGGLE = &H8000
'#End Region

'#Region Rebar Notifications
Public Const RBN_FIRST = (0 - 831)
Public Const RBN_HEIGHTCHANGE = (RBN_FIRST - 0)
Public Const RBN_GETOBJECT = (RBN_FIRST - 1)
Public Const RBN_LAYOUTCHANGED = (RBN_FIRST - 2)
Public Const RBN_AUTOSIZE = (RBN_FIRST - 3)
Public Const RBN_BEGINDRAG = (RBN_FIRST - 4)
Public Const RBN_ENDDRAG = (RBN_FIRST - 5)
Public Const RBN_DELETINGBAND = (RBN_FIRST - 6)
Public Const RBN_DELETEDBAND = (RBN_FIRST - 7)
Public Const RBN_CHILDSIZE = (RBN_FIRST - 8)
Public Const RBN_CHEVRONPUSHED = (RBN_FIRST - 10)
'#End Region

'#Region Rebar Messages
Public Const CCM_FIRST = &H2000
Public Const RB_INSERTBANDA = (WM_USER + 1)
Public Const RB_DELETEBAND = (WM_USER + 2)
Public Const RB_GETBARINFO = (WM_USER + 3)
Public Const RB_SETBARINFO = (WM_USER + 4)
Public Const RB_GETBANDINFO = (WM_USER + 5)
Public Const RB_SETBANDINFOA = (WM_USER + 6)
Public Const RB_SETPARENT = (WM_USER + 7)
Public Const RB_HITTEST = (WM_USER + 8)
Public Const RB_GETRECT = (WM_USER + 9)
Public Const RB_INSERTBANDW = (WM_USER + 10)
Public Const RB_SETBANDINFOW = (WM_USER + 11)
Public Const RB_GETBANDCOUNT = (WM_USER + 12)
Public Const RB_GETROWCOUNT = (WM_USER + 13)
Public Const RB_GETROWHEIGHT = (WM_USER + 14)
Public Const RB_IDTOINDEX = (WM_USER + 16)
Public Const RB_GETTOOLTIPS = (WM_USER + 17)
Public Const RB_SETTOOLTIPS = (WM_USER + 18)
Public Const RB_SETBKCOLOR = (WM_USER + 19)
Public Const RB_GETBKCOLOR = (WM_USER + 20)
Public Const RB_SETTEXTCOLOR = (WM_USER + 21)
Public Const RB_GETTEXTCOLOR = (WM_USER + 22)
Public Const RB_SIZETORECT = (WM_USER + 23)
Public Const RB_SETCOLORSCHEME = (CCM_FIRST + 2)
Public Const RB_GETCOLORSCHEME = (CCM_FIRST + 3)
Public Const RB_BEGINDRAG = (WM_USER + 24)
Public Const RB_ENDDRAG = (WM_USER + 25)
Public Const RB_DRAGMOVE = (WM_USER + 26)
Public Const RB_GETBARHEIGHT = (WM_USER + 27)
Public Const RB_GETBANDINFOW = (WM_USER + 28)
Public Const RB_GETBANDINFOA = (WM_USER + 29)
Public Const RB_MINIMIZEBAND = (WM_USER + 30)
Public Const RB_MAXIMIZEBAND = (WM_USER + 31)
Public Const RB_GETDROPTARGET = (CCM_FIRST + 4)
Public Const RB_GETBANDBORDERS = (WM_USER + 34)
Public Const RB_SHOWBAND = (WM_USER + 35)
Public Const RB_SETPALETTE = (WM_USER + 37)
Public Const RB_GETPALETTE = (WM_USER + 38)
Public Const RB_MOVEBAND = (WM_USER + 39)
Public Const RB_SETUNICODEFORMAT = (CCM_FIRST + 5)
Public Const RB_GETUNICODEFORMAT = (CCM_FIRST + 6)
'#End Region

'#Region Rebar Info Mask
Public Const RBBIM_STYLE = &H1
Public Const RBBIM_COLORS = &H2
Public Const RBBIM_TEXT = &H4
Public Const RBBIM_IMAGE = &H8
Public Const RBBIM_CHILD = &H10
Public Const RBBIM_CHILDSIZE = &H20
Public Const RBBIM_SIZE = &H40
Public Const RBBIM_BACKGROUND = &H80
Public Const RBBIM_ID = &H100
Public Const RBBIM_IDEALSIZE = &H200
Public Const RBBIM_LPARAM = &H400
Public Const BBIM_HEADERSIZE = &H800
'#End Region

'#Region Rebar Styles
Public Const RBBS_BREAK = &H1
Public Const RBBS_CHILDEDGE = &H4
Public Const RBBS_FIXEDBMP = &H20
Public Const RBBS_GRIPPERALWAYS = &H80
Public Const RBBS_USECHEVRON = &H200
'#End Region

'#Region Object types
Public Const OBJ_PEN = 1
Public Const OBJ_BRUSH = 2
Public Const OBJ_DC = 3
Public Const OBJ_METADC = 4
Public Const OBJ_PAL = 5
Public Const OBJ_FONT = 6
Public Const OBJ_BITMAP = 7
Public Const OBJ_REGION = 8
Public Const OBJ_METAFILE = 9
Public Const OBJ_MEMDC = 10
Public Const OBJ_EXTPEN = 11
Public Const OBJ_ENHMETADC = 12
Public Const OBJ_ENHMETAFILE = 13
'#End Region

'#Region WM_MENUCHAR Return values
Public Const MNC_IGNORE = 0
Public Const MNC_CLOSE = 1
Public Const MNC_EXECUTE = 2
Public Const MNC_SELECT = 3
'#End Region

'#Region Background Mode
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2
'#End Region

'#Region ListView Messages
Public Const LVM_FIRST = &H1000
Public Const LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
Public Const LVM_GETITEMSTATE = (LVM_FIRST + 44)
Public Const LVM_GETITEMTEXTW = (LVM_FIRST + 115)
'#End Region

'#Region Header Control Messages
Public Const HDM_FIRST = &H1200
Public Const HDM_GETITEMRECT = (HDM_FIRST + 7)
Public Const HDM_HITTEST = (HDM_FIRST + 6)
Public Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
Public Const HDM_GETITEMW = (HDM_FIRST + 11)
Public Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
'#End Region

'#Region Header Control Notifications
Public Const HDN_FIRST = (0 - 300)
Public Const HDN_BEGINTRACKW = (HDN_FIRST - 26)
Public Const HDN_ENDTRACKW = (HDN_FIRST - 27)
Public Const HDN_ITEMCLICKW = (HDN_FIRST - 22)
'#End Region

'#Region Header Control HitTest Flags
Public Const HHT_NOWHERE = &H1
Public Const HHT_ONHEADER = &H2
Public Const HHT_ONDIVIDER = &H4
Public Const HHT_ONDIVOPEN = &H8
Public Const HHT_ABOVE = &H100
Public Const HHT_BELOW = &H200
Public Const HHT_TORIGHT = &H400
Public Const HHT_TOLEFT = &H800
'#End Region

'#Region List View sub item portion
Public Const LVIR_BOUNDS = 0
Public Const LVIR_ICON = 1
Public Const LVIR_LABEL = 2
'#End Region

'#Region Tracker Event Flags
Public Const TME_HOVER = &H1
Public Const TME_LEAVE = &H2
Public Const TME_QUERY = &H40000000
Public Const TME_CANCEL = &H80000000
'#End Region

'#Region Mouse Activate Flags
Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4
'#End Region

'#Region Dialog Codes
Public Const DLGC_WANTARROWS = &H1
Public Const DLGC_WANTTAB = &H2
Public Const DLGC_WANTALLKEYS = &H4
Public Const DLGC_WANTMESSAGE = &H4
Public Const DLGC_HASSETSEL = &H8
Public Const DLGC_DEFPUSHBUTTON = &H10
Public Const DLGC_UNDEFPUSHBUTTON = &H20
Public Const DLGC_RADIOBUTTON = &H40
Public Const DLGC_WANTCHARS = &H80
Public Const DLGC_STATIC = &H100
Public Const DLGC_BUTTON = &H2000
'#End Region

'#Region Update Layered Windows Flags
Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const ULW_OPAQUE = &H4
'#End Region

'#Region Blend Flags
Public Const AC_SRC_OVER = &H0
Public Const AC_SRC_ALPHA = &H1
'#End Region

'#Region ComboBox messages
Public Const CB_GETDROPPEDSTATE = &H157
'#End Region

'#Region TreeView Messages
Public Const TV_FIRST = &H1100
Public Const TVM_GETITEMRECT = (TV_FIRST + 4)
Public Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Public Const TVM_HITTEST = (TV_FIRST + 17)
Public Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)
Public Const TVM_GETITEMW = (TV_FIRST + 62)
Public Const TVM_SETITEMW = (TV_FIRST + 63)
Public Const TVM_INSERTITEMW = (TV_FIRST + 50)
'#End Region

'#Region TreeViewImageListFlags
Public Const TVSIL_NORMAL = 0
Public Const TVSIL_STATE = 2
'#End Region

'#Region TreeViewItem Flags
Public Const TVIF_NONE = &H0
Public Const TVIF_TEXT = &H1
Public Const TVIF_IMAGE = &H2
Public Const TVIF_PARAM = &H4
Public Const TVIF_STATE = &H8
Public Const TVIF_HANDLE = &H10
Public Const TVIF_SELECTEDIMAGE = &H20
Public Const TVIF_CHILDREN = &H40
Public Const TVIF_INTEGRAL = &H80
Public Const I_CHILDRENCALLBACK = -1
Public Const LPSTR_TEXTCALLBACK = -1
'Public Const I_IMAGECALLBACK = -1
'Public Const I_IMAGENONE = -2
'#End Region

'#Region ListViewItem flags
Public Const LVIF_TEXT = &H1
Public Const LVIF_IMAGE = &H2
Public Const LVIF_PARAM = &H4
Public Const LVIF_STATE = &H8
Public Const LVIF_INDENT = &H10
Public Const LVIF_NORECOMPUTE = &H800
'#End Region

'#Region HeaderItem flags
Public Const HDI_WIDTH = &H1
Public Const HDI_HEIGHT = HDI_WIDTH
Public Const HDI_TEXT = &H2
Public Const HDI_FORMAT = &H4
Public Const HDI_LPARAM = &H8
Public Const HDI_BITMAP = &H10
Public Const HDI_IMAGE = &H20
Public Const HDI_DI_SETITEM = &H40
Public Const HDI_ORDER = &H80
'#End Region

'#Region GetDCExFlags
Public Const DCX_WINDOW = &H1
Public Const DCX_CACHE = &H2
Public Const DCX_NORESETATTRS = &H4
Public Const DCX_CLIPCHILDREN = &H8
Public Const DCX_CLIPSIBLINGS = &H10
Public Const DCX_PARENTCLIP = &H20
Public Const DCX_EXCLUDERGN = &H40
Public Const DCX_INTERSECTRGN = &H80
Public Const DCX_EXCLUDEUPDATE = &H100
Public Const DCX_INTERSECTUPDATE = &H200
Public Const DCX_LOCKWINDOWUPDATE = &H400
Public Const DCX_VALIDATE = &H200000
'#End Region

'#Region HitTest
Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const HTOBJECT = 19
Public Const HTCLOSE = 20
Public Const HTHELP = 21
'#End Region

'#Region ActivateFlags
Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1
Public Const WA_CLICKACTIVE = 2
'#End Region

'#Region StrechModeFlags
Public Const BLACKONWHITEConst = 1
Public Const WHITEONBLACK = 2
Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4
Public Const MAXSTRETCHBLTMODE = 4
'#End Region

'#Region ScrollBarFlags
Public Const SBS_HORZ = &H0
Public Const SBS_VERT = &H1
Public Const SBS_TOPALIGN = &H2
Public Const SBS_LEFTALIGN = &H2
Public Const SBS_BOTTOMALIGN = &H4
Public Const SBS_RIGHTALIGN = &H4
Public Const SBS_SIZEBOXTOPLEFTALIGN = &H2
Public Const SBS_SIZEBOXBOTTOMRIGHTALIGN = &H4
Public Const SBS_SIZEBOX = &H8
Public Const SBS_SIZEGRIP = &H10
'#End Region

'#Region System Metrics Codes
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVSCROLL = 2
Public Const SM_CYHSCROLL = 3
Public Const SM_CYCAPTION = 4
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CXDLGFRAME = 7
Public Const SM_CYDLGFRAME = 8
Public Const SM_CYVTHUMB = 9
Public Const SM_CXHTHUMB = 10
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXCURSOR = 13
Public Const SM_CYCURSOR = 14
Public Const SM_CYMENU = 15
Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CYKANJIWINDOW = 18
Public Const SM_MOUSEPRESENT = 19
Public Const SM_CYVSCROLL = 20
Public Const SM_CXHSCROLL = 21
Public Const SM_DEBUG = 22
Public Const SM_SWAPBUTTON = 23
Public Const SM_RESERVED1 = 24
Public Const SM_RESERVED2 = 25
Public Const SM_RESERVED3 = 26
Public Const SM_RESERVED4 = 27
Public Const SM_CXMIN = 28
Public Const SM_CYMIN = 29
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CXFRAME = 32
Public Const SM_CYFRAME = 33
Public Const SM_CXMINTRACK = 34
Public Const SM_CYMINTRACK = 35
Public Const SM_CXDOUBLECLK = 36
Public Const SM_CYDOUBLECLK = 37
Public Const SM_CXICONSPACING = 38
Public Const SM_CYICONSPACING = 39
Public Const SM_MENUDROPALIGNMENT = 40
Public Const SM_PENWINDOWS = 41
Public Const SM_DBCSENABLED = 42
Public Const SM_CMOUSEBUTTONS = 43
Public Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Public Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Public Const SM_CXSIZEFRAME = SM_CXFRAME
Public Const SM_CYSIZEFRAME = SM_CYFRAME
Public Const SM_SECURE = 44
Public Const SM_CXEDGE = 45
Public Const SM_CYEDGE = 46
Public Const SM_CXMINSPACING = 47
Public Const SM_CYMINSPACING = 48
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50
Public Const SM_CYSMCAPTION = 51
Public Const SM_CXSMSIZE = 52
Public Const SM_CYSMSIZE = 53
Public Const SM_CXMENUSIZE = 54
Public Const SM_CYMENUSIZE = 55
Public Const SM_ARRANGE = 56
Public Const SM_CXMINIMIZED = 57
Public Const SM_CYMINIMIZED = 58
Public Const SM_CXMAXTRACK = 59
Public Const SM_CYMAXTRACK = 60
Public Const SM_CXMAXIMIZED = 61
Public Const SM_CYMAXIMIZED = 62
Public Const SM_NETWORK = 63
Public Const SM_CLEANBOOT = 67
Public Const SM_CXDRAG = 68
Public Const SM_CYDRAG = 69
Public Const SM_SHOWSOUNDS = 70
Public Const SM_CXMENUCHECK = 71
Public Const SM_CYMENUCHECK = 72
Public Const SM_SLOWMACHINE = 73
Public Const SM_MIDEASTENABLED = 74
Public Const SM_MOUSEWHEELPRESENT = 75
Public Const SM_XVIRTUALSCREEN = 76
Public Const SM_YVIRTUALSCREEN = 77
Public Const SM_CXVIRTUALSCREEN = 78
Public Const SM_CYVIRTUALSCREEN = 79
Public Const SM_CMONITORS = 80
Public Const SM_SAMEDISPLAYFORMAT = 81
Public Const SM_CMETRICS = 83
'#End Region

'#Region ScrollBarTypes
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3
'#End Region

'#Region SrollBarInfoFlags
Public Const SIF_RANGE = &H1
Public Const SIF_PAGE = &H2
Public Const SIF_POS = &H4
Public Const SIF_DISABLENOSCROLL = &H8
Public Const SIF_TRACKPOS = &H10
Public Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
'#End Region

'#Region Enable ScrollBar flags
Public Const ESB_ENABLE_BOTH = &H0
Public Const ESB_DISABLE_BOTH = &H3
Public Const ESB_DISABLE_LEFT = &H1
Public Const ESB_DISABLE_RIGHT = &H2
Public Const ESB_DISABLE_UP = &H1
Public Const ESB_DISABLE_DOWN = &H2
Public Const ESB_DISABLE_LTUP = ESB_DISABLE_LEFT
Public Const ESB_DISABLE_RTDN = ESB_DISABLE_RIGHT
'#End Region

'#Region Scroll Requests
Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1
Public Const SB_PAGEUP = 2
Public Const SB_PAGELEFT = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGERIGHT = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_THUMBTRACK = 5
Public Const SB_TOP = 6
Public Const SB_LEFT = 6
Public Const SB_BOTTOM = 7
Public Const SB_RIGHT = 7
Public Const SB_ENDSCROLL = 8
'#End Region

'#Region SrollWindowEx flags
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_INVALIDATE = &H2
Public Const SW_ERASE = &H4
Public Const SW_SMOOTHSCROLL = &H10
'#End Region

'#region ImageListFlags
Public Const ILC_MASK = &H1
Public Const ILC_COLOR = &H0
Public Const ILC_COLORDDB = &HFE
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILC_PALETTE = &H800
'#end region

'#region ImageListDrawFlags
Public Const ILD_NORMAL = &H0
Public Const ILD_TRANSPARENT = &H1
Public Const ILD_MASK = &H10
Public Const ILD_IMAGE = &H20
Public Const ILD_ROP = &H40
Public Const ILD_BLEND25 = &H2
Public Const ILD_BLEND50 = &H4
Public Const ILD_OVERLAYMASK = &HF00
'#end region

'#region List View Notifications
Public Const LVN_FIRST = (0 - 100)
Public Const LVN_GETDISPINFOW = (LVN_FIRST - 77)
Public Const LVN_SETDISPINFOA = (LVN_FIRST - 51)
'#end region

'#region Drive Type
Public Const DRIVE_UNKNOWN = 0
Public Const DRIVE_NO_ROOT_DIR = 1
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
'#End region

'#region Shell File Info Flags
Public Const SHGFI_ICON = &H100
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_TYPENAME = &H400
Public Const SHGFI_ATTRIBUTES = &H800
Public Const SHGFI_ICONLOCATION = &H1000
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LINKOVERLAY = &H8000
Public Const SHGFI_SELECTED = &H10000
Public Const SHGFI_ATTR_SPECIFIED = &H20000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_OPENICON = &H2
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_PIDL = &H8
Public Const SHGFI_USEFILEATTRIBUTES = &H10
'#end region

'#region Shell Special Folder
Public Const CSIDL_DESKTOP = &H0
Public Const CSIDL_INTERNET = &H1
Public Const CSIDL_PROGRAMS = &H2
Public Const CSIDL_CONTROLS = &H3
Public Const CSIDL_PRINTERS = &H4
Public Const CSIDL_PERSONAL = &H5
Public Const CSIDL_FAVORITES = &H6
Public Const CSIDL_STARTUP = &H7
Public Const CSIDL_RECENT = &H8
Public Const CSIDL_SENDTO = &H9
Public Const CSIDL_BITBUCKET = &HA
Public Const CSIDL_STARTMENU = &HB
Public Const CSIDL_DESKTOPDIRECTORY = &H10
Public Const CSIDL_DRIVES = &H11
Public Const CSIDL_NETWORK = &H12
Public Const CSIDL_NETHOOD = &H13
Public Const CSIDL_FONTS = &H14
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16
Public Const CSIDL_COMMON_PROGRAMS = &H17
Public Const CSIDL_COMMON_STARTUP = &H18
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Public Const CSIDL_APPDATA = &H1A
Public Const CSIDL_PRINTHOOD = &H1B
Public Const CSIDL_ALTSTARTUP = &H1D
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
'#end region

'#region ImageList Draw Colors
Public Const CLR_NONE = &HFFFFFFFF
Public Const CLR_DEFAULT = &HFF000000
'#end region

'#region ShellEnumFlags
Public Const SHCONTF_FOLDERS = 32 '// For shell browser
Public Const SHCONTF_NONFOLDERS = 64 '// For Default view
Public Const SHCONTF_INCLUDEHIDDEN = 128 '// For hidden/system objects
'#end region

'#region ShellGetDisplayNameOfFlags
Public Const SHGDN_NORMALConst = 0 '// Default (display purpose)
Public Const SHGDN_INFOLDERConst = 1 '// displayed under a folder (relative)
Public Const SHGDN_INCLUDE_NONFILESYS = &H2000 '// If Not set display names For shell name space items that are Not in the file system will fail.
Public Const SHGDN_FORADDRESSBARConst = &H4000 '// For displaying in the address (drives dropdown) bar
Public Const SHGDN_FORPARSINGConst = &H8000 '// For ParseDisplayName Or path
'#end region

'#region STRRETFlags
Public Const STRRET_WSTR = &H0     '// Use STRRET.pOleStr
Public Const STRRET_OFFSET = &H1       '// Use STRRET.uOffset To Ansi
Public Const STRRET_CSTR = &H2     '// Use STRRET.cStr
'#end region

'#region GetAttributeOfFlags
Public Const DROPEFFECT_NONE = 0
Public Const DROPEFFECT_COPY = 1
Public Const DROPEFFECT_MOVE = 2
Public Const DROPEFFECT_LINK = 4
Public Const DROPEFFECT_SCROLL = &H80000000
Public Const SFGAO_CANCOPY = DROPEFFECT_COPY   '// Objects can be copied
Public Const SFGAO_CANMOVE = DROPEFFECT_MOVE   '// Objects can be moved
Public Const SFGAO_CANLINK = DROPEFFECT_LINK   '// Objects can be linked
Public Const SFGAO_CANRENAME = &H10        '// Objects can be renamed
Public Const SFGAO_CANDELETE = &H20        '// Objects can be deleted
Public Const SFGAO_HASPROPSHEET = &H40         '// Objects have property sheets
Public Const SFGAO_DROPTARGET = &H100      '// Objects are drop target
Public Const SFGAO_CAPABILITYMASK = &H177
Public Const SFGAO_LINK = &H10000      '// Shortcut (link)
Public Const SFGAO_SHARE = &H20000     '// shared
Public Const SFGAO_READONLY = &H40000      '// Read-only
Public Const SFGAO_GHOSTED = &H80000       '// ghosted icon
Public Const SFGAO_HIDDEN = &H80000    '// hidden Object
Public Const SFGAO_DISPLAYATTRMASK = &HF0000
Public Const SFGAO_FILESYSANCESTOR = &H10000000    '// It contains file system folder
Public Const SFGAO_FOLDER = &H20000000 '// It's a folder.
Public Const SFGAO_FILESYSTEM = &H40000000 '// is a file system thing (file/folder/root)
Public Const SFGAO_HASSUBFOLDER = &H80000000   '// Expandable in the map pane
Public Const SFGAO_CONTENTSMASK = &H80000000
Public Const SFGAO_VALIDATE = &H1000000    '// invalidate cached information
Public Const SFGAO_REMOVABLE = &H2000000   '// is this removeable media?
Public Const SFGAO_COMPRESSED = &H4000000  '// Object is compressed (use alt Color)
Public Const SFGAO_BROWSABLE = &H8000000   '// is in-place browsable
Public Const SFGAO_NONENUMERATED = &H100000    '// is a non-enumerated Object
Public Const SFGAO_NEWCONTENT = &H200000   '// should show bold in explorer tree
'#end region

'#region ListViewItemState
Public Const LVIS_FOCUSED = &H1
Public Const LVIS_SELECTED = &H2
Public Const LVIS_CUT = &H4
Public Const LVIS_DROPHILITED = &H8
Public Const LVIS_ACTIVATING = &H20
Public Const LVIS_OVERLAYMASK = &HF00
Public Const LVIS_STATEIMAGEMASK = &HF000
'#end region

'#region TreeViewItemInsertPosition
Public Const TVI_ROOT = &HFFFF0000
Public Const TVI_FIRST = &HFFFF0001
Public Const TVI_LAST = &HFFFF0002
Public Const TVI_SORT = &HFFFF0003
'#end region

'#region TreeViewNotifications
Public Const TVN_FIRST = -400
Public Const TVN_SELCHANGINGA = (TVN_FIRST - 1)
Public Const TVN_SELCHANGINGW = (TVN_FIRST - 50)
Public Const TVN_SELCHANGEDA = (TVN_FIRST - 2)
Public Const TVN_SELCHANGEDW = (TVN_FIRST - 51)
Public Const TVN_GETDISPINFOA = (TVN_FIRST - 3)
Public Const TVN_GETDISPINFOW = (TVN_FIRST - 52)
Public Const TVN_SETDISPINFOA = (TVN_FIRST - 4)
Public Const TVN_SETDISPINFOW = (TVN_FIRST - 53)
Public Const TVN_ITEMEXPANDINGA = (TVN_FIRST - 5)
Public Const TVN_ITEMEXPANDINGW = (TVN_FIRST - 54)
Public Const TVN_ITEMEXPANDEDA = (TVN_FIRST - 6)
Public Const TVN_ITEMEXPANDEDW = (TVN_FIRST - 55)
Public Const TVN_BEGINDRAGA = (TVN_FIRST - 7)
Public Const TVN_BEGINDRAGW = (TVN_FIRST - 56)
Public Const TVN_BEGINRDRAGA = (TVN_FIRST - 8)
Public Const TVN_BEGINRDRAGW = (TVN_FIRST - 57)
Public Const TVN_DELETEITEMA = (TVN_FIRST - 9)
Public Const TVN_DELETEITEMW = (TVN_FIRST - 58)
Public Const TVN_BEGINLABELEDITA = (TVN_FIRST - 10)
Public Const TVN_BEGINLABELEDITW = (TVN_FIRST - 59)
Public Const TVN_ENDLABELEDITA = (TVN_FIRST - 11)
Public Const TVN_ENDLABELEDITW = (TVN_FIRST - 60)
Public Const TVN_KEYDOWN = (TVN_FIRST - 12)
Public Const TVN_GETINFOTIPA = (TVN_FIRST - 13)
Public Const TVN_GETINFOTIPW = (TVN_FIRST - 14)
Public Const TVN_SINGLEEXPAND = (TVN_FIRST - 15)
'#end region

'#region TreeViewItemExpansion
Public Const TVE_COLLAPSE = &H1
Public Const TVE_EXPAND = &H2
Public Const TVE_TOGGLE = &H3
Public Const TVE_EXPANDPARTIAL = &H4000
Public Const TVE_COLLAPSERESET = &H8000
'#end region

'#region WinErrors
Public Const NOERROR = &H0
'#end region

'#region TreeViewHitTest
Public Const TVHT_NOWHERE = &H1
Public Const TVHT_ONITEMICON = &H2
Public Const TVHT_ONITEMLABEL = &H4
Public Const TVHT_ONITEMINDENT = &H8
Public Const TVHT_ONITEMBUTTON = &H10
Public Const TVHT_ONITEMRIGHT = &H20
Public Const TVHT_ONITEMSTATEICON = &H40
Public Const TVHT_ABOVE = &H100
Public Const TVHT_BELOW = &H200
Public Const TVHT_TORIGHT = &H400
Public Const TVHT_TOLEFT = &H800
Public Const TVHT_ONITEM = (TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON)
'#End Region

'#Region TreeViewItemState
Public Const TVIS_SELECTED = &H2
Public Const TVIS_CUT = &H4
Public Const TVIS_DROPHILITED = &H8
Public Const TVIS_BOLD = &H10
Public Const TVIS_EXPANDED = &H20
Public Const TVIS_EXPANDEDONCE = &H40
Public Const TVIS_EXPANDPARTIAL = &H80
Public Const TVIS_OVERLAYMASK = &HF00
Public Const TVIS_STATEIMAGEMASK = &HF000
Public Const TVIS_USERMASK = &HF000
'#End Region

'#Region Windows System Objects
'// Reserved IDs For system objects
Public Const OBJID_WINDOW = &H0
Public Const OBJID_SYSMENU = &HFFFFFFFF
Public Const OBJID_TITLEBAR = &HFFFFFFFE
Public Const OBJID_MENU = &HFFFFFFFD
Public Const OBJID_CLIENT = &HFFFFFFFC
Public Const OBJID_VSCROLL = &HFFFFFFFB
Public Const OBJID_HSCROLL = &HFFFFFFFA
Public Const OBJID_SIZEGRIP = &HFFFFFFF9
Public Const OBJID_CARET = &HFFFFFFF8
Public Const OBJID_CURSOR = &HFFFFFFF7
Public Const OBJID_ALERT = &HFFFFFFF6
Public Const OBJID_SOUND = &HFFFFFFF5
'#End Region

'#Region SystemState

Public Const STATE_SYSTEM_UNAVAILABLE = &H1        '// Disabled
Public Const STATE_SYSTEM_SELECTED = &H2
Public Const STATE_SYSTEM_FOCUSED = &H4
Public Const STATE_SYSTEM_PRESSED = &H8
Public Const STATE_SYSTEM_CHECKED = &H10
Public Const STATE_SYSTEM_MIXED = &H20       '// 3-state checkbox Or toolbar button
Public Const STATE_SYSTEM_READONLY = &H40
Public Const STATE_SYSTEM_HOTTRACKED = &H80
Public Const STATE_SYSTEM_DEFAULT = &H100
Public Const STATE_SYSTEM_EXPANDED = &H200
Public Const STATE_SYSTEM_COLLAPSED = &H400
Public Const STATE_SYSTEM_BUSY = &H800
Public Const STATE_SYSTEM_FLOATING = &H1000     '// Children "owned" Not "contained" by parent
Public Const STATE_SYSTEM_MARQUEED = &H2000
Public Const STATE_SYSTEM_ANIMATED = &H4000
Public Const STATE_SYSTEM_INVISIBLE = &H8000
Public Const STATE_SYSTEM_OFFSCREEN = &H10000
Public Const STATE_SYSTEM_SIZEABLE = &H20000
Public Const STATE_SYSTEM_MOVEABLE = &H40000
Public Const STATE_SYSTEM_SELFVOICING = &H80000
Public Const STATE_SYSTEM_FOCUSABLE = &H100000
Public Const STATE_SYSTEM_SELECTABLE = &H200000
Public Const STATE_SYSTEM_LINKED = &H400000
Public Const STATE_SYSTEM_TRAVERSED = &H800000
Public Const STATE_SYSTEM_MULTISELECTABLE = &H1000000  '// Supports multiple selection
Public Const STATE_SYSTEM_EXTSELECTABLE = &H2000000  '// Supports extended selection
Public Const STATE_SYSTEM_ALERT_LOW = &H4000000  '// This information is of low priority
Public Const STATE_SYSTEM_ALERT_MEDIUM = &H8000000  '// This information is of medium priority
Public Const STATE_SYSTEM_ALERT_HIGH = &H10000000 '// This information is of high priority
Public Const STATE_SYSTEM_VALID = &H1FFFFFFF
'#End Region


'#Region QueryContextMenuFlags
Public Const CMF_NORMAL = &H0
Public Const CMF_DEFAULTONLY = &H1
Public Const CMF_VERBSONLY = &H2
Public Const CMF_EXPLORE = &H4
Public Const CMF_NOVERBS = &H8
Public Const CMF_CANRENAME = &H10
Public Const CMF_NODEFAULT = &H20
Public Const CMF_INCLUDESTATIC = &H40
Public Const CMF_RESERVED = &HFFFF0000
'#End Region

'#Region GetWindowLongFlags
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_USERDATA = (-21)
Public Const GWL_ID = (-12)
'#End Region

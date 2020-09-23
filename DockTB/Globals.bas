Attribute VB_Name = "Globals"
' ----------------------------------------------------------------- '
' Filename: Globals.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Global constants, types, API function declarations etc.
' ----------------------------------------------------------------- '

Option Explicit

'----------- Made up constants -----------
Public Const afxComCtlVersion = 400
Public Const VERSION_IE4 = 400
'-----------------------------------------

' Frame window styles
Public Const FWS_ADDTOTITLE = &H8000            ' modify title based on content
Public Const FWS_PREFIXTITLE = &H4000           ' show document name before app name
Public Const FWS_SNAPTOBARS = &H2000            ' snap size to size of contained bars
'-----------------------------------------

' parts of Main Frame
Public Const AFX_IDW_PANE_FIRST = &HE900&       ' first pane (256 max)
Public Const AFX_IDW_PANE_LAST = &HE9FF&
Public Const AFX_IDW_HSCROLL_FIRST = &HEA00&    ' first Horz scrollbar (16 max)
Public Const AFX_IDW_VSCROLL_FIRST = &HEA10&    ' first Vert scrollbar (16 max)
'-----------------------------------------

Public Const MFS_SYNCACTIVE = &H100             ' syncronize activation w/ parent
Public Const MFS_4THICKFRAME = &H200            ' thick frame all around (no tiles)
Public Const MFS_THICKFRAME = &H400             ' use instead of WS_THICKFRAME
Public Const MFS_MOVEFRAME = &H800              ' no sizing, just moving
Public Const MFS_BLOCKSYSMENU = &H1000          ' block hit testing on system menu

'----------------------------------------------------------------------------------------------------------------
' Layout Modes for CalcDynamicLayout
Public Const LM_STRETCH = &H1     ' same meaning as bStretch in CalcFixedLayout.  If set, ignores nLength
                                  ' and returns dimensions based on LM_HORZ state, otherwise LM_HORZ is used
                                  ' to determine if nLength is the desired horizontal or vertical length
                                  ' and dimensions are returned based on nLength
Public Const LM_HORZ = &H2        ' same as bHorz in CalcFixedLayout
Public Const LM_MRUWIDTH = &H4    ' Most Recently Used Dynamic Width
Public Const LM_HORZDOCK = &H8    ' Horizontal Docked Dimensions
Public Const LM_VERTDOCK = &H10   ' Vertical Docked Dimensions
Public Const LM_LENGTHY = &H20    ' Set if nLength is a Height instead of a Width
Public Const LM_COMMIT = &H40     ' Remember MRUWidth

' Note: If your application supports docking toolbars, you should
' not use the following IDs for your own toolbars.  The IDs chosen
' are at the top of the first 32 such that the bars will be hidden
' while in print preview mode, and are not likely to conflict with
' IDs your application may have used succesfully in the past.
Public Const AFX_IDW_DOCKBAR_TOP = &HE81B&
Public Const AFX_IDW_DOCKBAR_LEFT = &HE81C&
Public Const AFX_IDW_DOCKBAR_RIGHT = &HE81D&
Public Const AFX_IDW_DOCKBAR_BOTTOM = &HE81E&
Public Const AFX_IDW_DOCKBAR_FLOAT = &HE81F&
'----------------------------------------------------------------------------------------------------------------


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetDCEx Lib "user32" (ByVal hWnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Public Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetLastActivePopup Lib "user32" (ByVal hwndOwnder As Long) As Long
Public Declare Function InvalidateRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bErase As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function StretchDIBits_LONG Lib "gdi32" Alias "StretchDIBits" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Long, lpBitsInfo As Long, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Public Declare Function FindResourceNEW Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    ' ShowWindow Selectors
    Public Const SW_HIDE = 0
    Public Const SW_SHOW = 5
    Public Const SW_SHOWNA = 8

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    ' GetWindowLong constants
    Public Const GWL_HWNDPARENT = (-8)
    Public Const GWL_WNDPROC = (-4)
    Public Const GWL_USERDATA = (-21)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Public Const MAX_PATH_ = 260

Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function GetClipBox Lib "gdi32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function AdjustWindowRectEx Lib "user32" (lpRect As RECT, ByVal dsStyle As Long, ByVal bMenu As Long, ByVal dwEsStyle As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
    Public Const WHITE_BRUSH = 0
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Const SM_CXFRAME = 32
    Public Const SM_CYFRAME = 33
    Public Const SM_CXBORDER = 5
    Public Const SM_CYBORDER = 6
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Public Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Public Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Public Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDlgCtrlID Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Public Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function IntersectClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

'-----------------------------------------

Public Type COLORREF
    x As Long
End Type

'------------ Window Messages ------------
Public Const WM_USER = &H400
Public Const WM_DDE_FIRST = &H3E0
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_CANCELMODE = &H1F
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_CHAR = &H102
Public Const WM_CHARTOITEM = &H2F
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_CHOOSEFONT_GETLOGFONT = (WM_USER + 1)
Public Const WM_CHOOSEFONT_SETFLAGS = (WM_USER + 102)
Public Const WM_CHOOSEFONT_SETLOGFONT = (WM_USER + 101)
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMMNOTIFY = &H44
Public Const WM_COMPACTING = &H41
Public Const WM_COMPAREITEM = &H39
Public Const WM_CONVERTREQUESTEX = &H108
Public Const WM_COPY = &H301
Public Const WM_COPYDATA = &H4A
Public Const WM_CREATE = &H1
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CUT = &H300
Public Const WM_DDE_ACK = (WM_DDE_FIRST + 4)
Public Const WM_DDE_ADVISE = (WM_DDE_FIRST + 2)
Public Const WM_DDE_DATA = (WM_DDE_FIRST + 5)
Public Const WM_DDE_EXECUTE = (WM_DDE_FIRST + 8)
Public Const WM_DDE_INITIATE = (WM_DDE_FIRST)
Public Const WM_DDE_LAST = (WM_DDE_FIRST + 8)
Public Const WM_DDE_POKE = (WM_DDE_FIRST + 7)
Public Const WM_DDE_REQUEST = (WM_DDE_FIRST + 6)
Public Const WM_DDE_TERMINATE = (WM_DDE_FIRST + 1)
Public Const WM_DDE_UNADVISE = (WM_DDE_FIRST + 3)
Public Const WM_DEADCHAR = &H103
Public Const WM_DELETEITEM = &H2D
Public Const WM_DESTROY = &H2
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_DRAWITEM = &H2B
Public Const WM_DROPFILES = &H233
Public Const WM_ENABLE = &HA
Public Const WM_ENDSESSION = &H16
Public Const WM_ENTERIDLE = &H121
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_ERASEBKGND = &H14
Public Const WM_EXITMENULOOP = &H212
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETDLGCODE = &H87
Public Const WM_GETFONT = &H31
Public Const WM_GETHOTKEY = &H33
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_IME_CHAR = &H286
Public Const WM_IME_COMPOSITION = &H10F
Public Const WM_IME_COMPOSITIONFULL = &H284
Public Const WM_IME_CONTROL = &H283
Public Const WM_IME_ENDCOMPOSITION = &H10E
Public Const WM_IME_KEYDOWN = &H290
Public Const WM_IME_KEYLAST = &H10F
Public Const WM_IME_KEYUP = &H291
Public Const WM_IME_NOTIFY = &H282
Public Const WM_IME_SELECT = &H285
Public Const WM_IME_SETCONTEXT = &H281
Public Const WM_IME_STARTCOMPOSITION = &H10D
Public Const WM_INITDIALOG = &H110
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDINEXT = &H224
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDISETMENU = &H230
Public Const WM_MDITILE = &H226
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_NULL = &H0
Public Const WM_OTHERWINDOWCREATED = &H42
Public Const WM_OTHERWINDOWDESTROYED = &H43
Public Const WM_PAINT = &HF
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_PAINTICON = &H26
Public Const WM_PALETTECHANGED = &H311
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_PASTE = &H302
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_POWER = &H48
Public Const WM_PSD_ENVSTAMPRECT = (WM_USER + 5)
Public Const WM_PSD_FULLPAGERECT = (WM_USER + 1)
Public Const WM_PSD_GREEKTEXTRECT = (WM_USER + 4)
Public Const WM_PSD_MARGINRECT = (WM_USER + 3)
Public Const WM_PSD_MINMARGINRECT = (WM_USER + 2)
Public Const WM_PSD_PAGESETUPDLG = (WM_USER)
Public Const WM_PSD_YAFULLPAGERECT = (WM_USER + 6)
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_QUERYOPEN = &H13
Public Const WM_QUEUESYNC = &H23
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_RENDERFORMAT = &H305
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETHOTKEY = &H32
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSCOMMAND = &H112
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_TIMECHANGE = &H1E
Public Const WM_TIMER = &H113
Public Const WM_UNDO = &H304
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_VSCROLL = &H115
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WININICHANGE = &H1A
Public Const WM_EXITSIZEMOVE = &H232
Public Const WM_MOVING = &H216

Public Const WM_ACTIVATE = &H6
    Public Const WA_ACTIVE = 1
    Public Const WA_CLICKACTIVE = 2
    Public Const WA_INACTIVE = 0

Public Const MA_ACTIVATE = 1
Public Const MA_ACTIVATEANDEAT = 2
Public Const MA_NOACTIVATE = 3
Public Const MA_NOACTIVATEANDEAT = 4
'-----------------------------------------


Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_POPUP = &H80000000
Public Const WS_SYSMENU = &H80000
Public Const WS_TABSTOP = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_ICONIC = WS_MINIMIZE


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

Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

Public Declare Function SetWindowPos Lib "user32" _
                (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                 ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                 ByVal wFlags As Long) As Long
    '---------- SetWindowPos Flags ----------
    Public Const SWP_NOSIZE = &H1
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOZORDER = &H4
    Public Const SWP_NOREDRAW = &H8
    Public Const SWP_NOACTIVATE = &H10
    Public Const SWP_FRAMECHANGED = &H20                ' The frame changed: send WM_NCCALCSIZE
    Public Const SWP_SHOWWINDOW = &H40
    Public Const SWP_HIDEWINDOW = &H80
    Public Const SWP_NOCOPYBITS = &H100
    Public Const SWP_NOOWNERZORDER = &H200              ' Don't do owner Z ordering
    Public Const SWP_NOSENDCHANGING = &H400             ' Don't send WM_WINDOWPOSCHANGING
    Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
    Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
    Public Const SWP_DEFERERASE = &H2000
    Public Const SWP_ASYNCWINDOWPOS = &H4000

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type WINDOWPOS
        hWnd As Long
        hWndInsertAfter As Long
        x As Long
        y As Long
        cx As Long
        cy As Long
        flags As Long
End Type

Type TBADDBITMAP
        hInst As Long
        nID As Long
End Type


'------------------- CUSTOM WINDOW CLASS FUNCTIONS (BEGIN) -------------------
Public Type WNDCLASS
    style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type


Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

Public Const CS_BYTEALIGNCLIENT = &H1000
Public Const CS_BYTEALIGNWINDOW = &H2000
Public Const CS_CLASSDC = &H40
Public Const CS_DBLCLKS = &H8
Public Const CS_HREDRAW = &H2
Public Const CS_INSERTCHAR = &H2000
Public Const CS_KEYCVTWINDOW = &H4
Public Const CS_NOCLOSE = &H200
Public Const CS_NOKEYCVT = &H100
Public Const CS_NOMOVECARET = &H4000
Public Const CS_OWNDC = &H20
Public Const CS_PARENTDC = &H80
Public Const CS_PUBLICCLASS = &H4000
Public Const CS_SAVEBITS = &H800
Public Const CS_VREDRAW = &H1
'-------------------- CUSTOM WINDOW CLASS FUNCTIONS (END) --------------------


'------------------------- CURSOR FUNCTIONS (BEGIN) --------------------------
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Public Const IDC_APPSTARTING = 32650&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZE = 32640&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&
Public Const IDCANCEL = 2
'-------------------------- CURSOR FUNCTIONS (END) ---------------------------


Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_FRAME = &H400
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_INVALIDATE = &H1
Public Const RDW_NOCHILDREN = &H40
Public Const RDW_NOERASE = &H20
Public Const RDW_NOFRAME = &H800
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_UPDATENOW = &H100
Public Const RDW_VALIDATE = &H8


Public Const PATINVERT = &H5A0049           ' (DWORD) dest = pattern XOR dest
Public Const RGN_XOR = 3


Public Type BITMAP                          ' 14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type Size
        cx As Long
        cy As Long
End Type


Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100       ' shorts
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8


Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    dwExStyle As Long
End Type


'------------------------------------
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
'------------------------------------


'---------- Size Constants ----------
Public Const SIZE_MAXHIDE = 4
Public Const SIZE_MAXIMIZED = 2
Public Const SIZE_MAXSHOW = 3
Public Const SIZE_MINIMIZED = 1
Public Const SIZE_RESTORED = 0
Public Const SIZEFULLSCREEN = SIZE_MAXIMIZED
Public Const SIZEICONIC = SIZE_MINIMIZED
Public Const SIZENORMAL = SIZE_RESTORED
Public Const SIZEPALETTE = 104                  ' Number of entries in physical palette
Public Const SIZEZOOMHIDE = SIZE_MAXHIDE
Public Const SIZEZOOMSHOW = SIZE_MAXSHOW
'------------------------------------


' special struct for WM_SIZEPARENT
Type AFX_SIZEPARENTPARAMS
    hDWP As Long            ' handle for DeferWindowPos
    recta As RECT           ' parent client rectangle (trim as appropriate)
    sizeTotal As Size       ' total size on each side as layout proceeds
    bStretch As Boolean     ' should stretch to fill all space
End Type

Public Const WM_SIZEPARENT = &H361          ' lParam = &AFX_SIZEPARENTPARAMS

'--------------------------------------------------------
Public Type NCCALCSIZE_PARAMS
    rgrc(0 To 2) As RECT
    lppos As WINDOWPOS
End Type


Public Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type


Public Const VK_CONTROL = &H11
Public Const VK_SHIFT = &H10
Public Const VK_ESCAPE = &H1B
Public Const PM_NOREMOVE = &H0
Public Const DCX_WINDOW = &H1&
Public Const DCX_LOCKWINDOWUPDATE = &H400&
Public Const DCX_CACHE = &H2&


'-----------------------------------------
Public Const SC_ARRANGE = &HF110&
Public Const SC_CLOSE = &HF060&
Public Const SC_GROUP_IDENTIFIER = "+"
Public Const SC_HOTKEY = &HF150&
Public Const SC_HSCROLL = &HF080&
Public Const SC_KEYMENU = &HF100&
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MAXIMIZE = &HF030&
Public Const SC_MINIMIZE = &HF020&
Public Const SC_ICON = SC_MINIMIZE
Public Const SC_MOUSEMENU = &HF090&
Public Const SC_MOVE = &HF010&
Public Const SC_NEXTWINDOW = &HF040&
Public Const SC_PREVWINDOW = &HF050&
Public Const SC_RESTORE = &HF120&
Public Const SC_SCREENSAVE = &HF140&
Public Const SC_SIZE = &HF000&
Public Const SC_TASKLIST = &HF130&
Public Const SC_VSCROLL = &HF070&
Public Const SC_ZOOM = SC_MAXIMIZE
'-----------------------------------------


'-----------------------------------------
Public Const MF_APPEND = &H100&
Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&
'-----------------------------------------


'-----------------------------------------
Public Const WM_FLOATSTATUS = &H36D         ' wParam combination of FS_* flags below
' flags for wParam in the WM_FLOATSTATUS message
Public Const FS_SHOW = &H1
Public Const FS_HIDE = &H2
Public Const FS_ACTIVATE = &H4
Public Const FS_DEACTIVATE = &H8
Public Const FS_ENABLE = &H10
Public Const FS_DISABLE = &H20
Public Const FS_SYNCACTIVE = &H40
'-----------------------------------------


'-----------------------------------------
Public Const WF_STAYACTIVE = &H20           ' look active even though not active
Public Const WF_NOPOPMSG = &H40             ' ignore WM_POPMESSAGESTRING calls
Public Const WF_MODALDISABLE = &H80         ' window is disabled
Public Const WF_KEEPMINIACTIVE = &H200      ' stay activate even though you are deactivated
'-----------------------------------------


Public Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type
Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Const RT_BITMAP = 2&


'-----------------------------------------------------------------------
Global DockBars(0 To 3) As Long

Type tagAFX_OLDTOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    recta As RECT
    hInst As Long
    lpszText As Long
End Type

Type tagTOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    recta As RECT
    hInst As Long
    lpszText As Long
    lParam As Long
End Type

Public Type HKEY
    x As Long
End Type

Public Type LPCSTR
    x As String
End Type

Public Type LPCWSTR
    x As String
End Type

Type AFX_CONTROLPOS
    nIndex As Long
    nID As Long
    rectOldPos As RECT
End Type


Public Const CW_USEDEFAULT As Long = &H80000000

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_ID = (-12)

Public Const CBRS_ALIGN_LEFT = &H1000&
Public Const CBRS_ALIGN_TOP = &H2000&
Public Const CBRS_ALIGN_RIGHT = &H4000&
Public Const CBRS_ALIGN_BOTTOM = &H8000&
Public Const CBRS_ALIGN_ANY = &HF000&
Public Const CBRS_BORDER_LEFT = &H100&
Public Const CBRS_BORDER_TOP = &H200&
Public Const CBRS_BORDER_RIGHT = &H400&
Public Const CBRS_BORDER_BOTTOM = &H800&
Public Const CBRS_BORDER_ANY = &HF00&
Public Const CBRS_TOOLTIPS = &H10&
Public Const CBRS_FLYBY = &H20&
Public Const CBRS_FLOAT_MULTI = &H40&
Public Const CBRS_BORDER_3D = &H80&
Public Const CBRS_HIDE_INPLACE = &H8&
Public Const CBRS_SIZE_DYNAMIC = &H4&
Public Const CBRS_SIZE_FIXED = &H2&
Public Const CBRS_FLOATING = &H1&
Public Const CBRS_GRIPPER = &H400000
Public Const CBRS_ORIENT_HORZ = (CBRS_ALIGN_TOP Or CBRS_ALIGN_BOTTOM)
Public Const CBRS_ORIENT_VERT = (CBRS_ALIGN_LEFT Or CBRS_ALIGN_RIGHT)
Public Const CBRS_ORIENT_ANY = (CBRS_ORIENT_HORZ Or CBRS_ORIENT_VERT)
Public Const CBRS_ALL = &H40FFFF

Public Const CBRS_NOALIGN = &H0&
Public Const CBRS_LEFT = (CBRS_ALIGN_LEFT Or CBRS_BORDER_RIGHT)
Public Const CBRS_TOP = (CBRS_ALIGN_TOP Or CBRS_BORDER_BOTTOM)
Public Const CBRS_RIGHT = (CBRS_ALIGN_RIGHT Or CBRS_BORDER_LEFT)
Public Const CBRS_BOTTOM = (CBRS_ALIGN_BOTTOM Or CBRS_BORDER_TOP)


Public Const HTBORDER = 18
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTCAPTION = 2
Public Const HTCLIENT = 1
Public Const HTERROR = (-2)
Public Const HTGROWBOX = 4
Public Const HTHSCROLL = 6
Public Const HTLEFT = 10
Public Const HTMAXBUTTON = 9
Public Const HTMENU = 5
Public Const HTMINBUTTON = 8
Public Const HTNOWHERE = 0
Public Const HTREDUCE = HTMINBUTTON
Public Const HTRIGHT = 11
Public Const HTSIZE = HTGROWBOX
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT
Public Const HTSYSMENU = 3
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTTRANSPARENT = (-1)
Public Const HTVSCROLL = 7
Public Const HTZOOM = HTMAXBUTTON
'-----------------------------------------------------------------------


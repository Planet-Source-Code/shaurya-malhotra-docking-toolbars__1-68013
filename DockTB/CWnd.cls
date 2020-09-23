VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CWnd.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CWnd class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Public hWnd As Long
Public hWndOwner As Long
Public dwStyle As Long
Public nFlags As Long


Private Property Get m_nFlags() As Long
    m_nFlags = nFlags
End Property

Private Property Let m_nFlags(RHS As Long)
    nFlags = RHS
End Property

Friend Function Invalidate(Optional bErase As Boolean = True)
    Call InvalidateRgn(m_hWnd, 0, bErase)
End Function

Friend Function Create(lpszClassName As String, lpszWindowName As String, dwStyle As Long, _
                    recta As RECT, pParentWnd As Long, nID As Long, Optional pContext As Long = 0, _
                    Optional obj As Object = Nothing) As Boolean
    ' can't use for desktop or pop-up windows (use CreateEx instead)
    'ASSERT(pParentWnd != NULL);
    'ASSERT((dwStyle & WS_POPUP) == 0);

    Create = CreateEx(0, lpszClassName, lpszWindowName, dwStyle Or WS_CHILD, _
                recta.Left, recta.Top, recta.Right - recta.Left, recta.Bottom - recta.Top, _
                pParentWnd, nID, pContext, obj)
End Function


Public Function CreateEx(dwExStyle As Long, lpszClassName As String, _
                    lpszWindowName As String, dwStyle As Long, _
                    x As Long, y As Long, nWidth As Long, nHeight As Long, _
                    hWndParent As Long, nIDorHMenu As Long, lpParam As Long, _
                    Optional obj As Object = Nothing) As Boolean

    ' allow modification of several common create parameters
    Dim cs As CREATESTRUCT
    cs.dwExStyle = dwExStyle
    cs.lpszClass = lpszClassName
    cs.lpszName = lpszWindowName
    cs.style = dwStyle
    cs.x = x
    cs.y = y
    cs.cx = nWidth
    cs.cy = nHeight
    cs.hWndParent = hWndParent
    cs.hMenu = nIDorHMenu
    cs.hInstance = App.hInstance
    cs.lpCreateParams = lpParam

    If (Not (obj Is Nothing)) Then
        Call PreCreateWindow(cs, obj)
    End If

    Dim hWnd As Long
    hWnd = CreateWindowEx(cs.dwExStyle, cs.lpszClass, _
                    cs.lpszName, cs.style, cs.x, cs.y, cs.cx, cs.cy, _
                    cs.hWndParent, cs.hMenu, cs.hInstance, cs.lpCreateParams)

    If (hWnd = 0) Then
        CreateEx = False
        Exit Function
    End If

    Me.hWnd = hWnd
    CreateEx = True

End Function


Friend Function OnSize(nType As Long, cx As Long, cy As Long)
'
End Function


Friend Function RepositionBars(nIDFirst As Long, nIDLast As Long, nIDLeftOver As Long, _
                    nFlags As Long, ByRef lpRectParam As RECT, ByRef lpRectClient As RECT, Optional bStretch As Boolean = True)

    'ASSERT(nFlags == 0 || nFlags == reposQuery || nFlags == reposExtra);

    ' walk kids in order, control bars get the resize notification
    ' which allow them to shrink the client area
    ' remaining size goes to the 'nIDLeftOver' pane
    ' NOTE: nIDFirst->nIDLast are usually 0->0xffff

    Dim layout As AFX_SIZEPARENTPARAMS
    Dim hWndLeftOver As Long
    hWndLeftOver = 0

    layout.bStretch = bStretch

    layout.sizeTotal.cx = 0
    layout.sizeTotal.cy = 0

    If Not CBool(IsRectEmpty(lpRectClient)) Then
        layout.recta = lpRectClient                 ' starting rect comes from parameter
    Else
        Call GetClientRect(m_hWnd, layout.recta)    ' starting rect comes from client rect
    End If

    If (nFlags <> reposQuery) Then
        layout.hDWP = BeginDeferWindowPos(8)    ' reasonable guess
    Else
        layout.hDWP = 0                         ' not actually doing layout
    End If

    Dim hWndChild As Long
    hWndChild = GetTopWindow(m_hWnd)

    Do While (hWndChild <> 0)
        Dim nIDC As Long
        nIDC = AfxGetDlgCtrlID(hWndChild)

        Dim pWnd As Long
        pWnd = FromHandlePermanent(hWndChild)

        If (nIDC = nIDLeftOver) Then
            hWndLeftOver = hWndChild
        ElseIf ((nIDC >= nIDFirst) And (nIDC <= nIDLast) And (pWnd <> 0)) Then
            Call SendMessage(hWndChild, WM_SIZEPARENT, 0, layout)
        End If

        hWndChild = GetNextWindow(hWndChild, GW_HWNDNEXT)
    Loop

    ' if just getting the available rectangle, return it now...
    If (nFlags = reposQuery) Then
        'ASSERT(lpRectParam != NULL);

        If (bStretch) Then
            Call CopyRect(lpRectParam, layout.recta)
        Else
            lpRectParam.Left = 0
            lpRectParam.Top = 0

            lpRectParam.Right = layout.sizeTotal.cx
            lpRectParam.Bottom = layout.sizeTotal.cy
        End If

        Exit Function
    End If

    ' the rest is the client size of the left-over pane
    If ((nIDLeftOver <> 0) And (hWndLeftOver <> 0)) Then
        Dim pLeftOver As Long
        pLeftOver = FromHandle(hWndLeftOver)

        ' allow extra space as specified by lpRectBorder
        If (nFlags = reposExtra) Then
            'ASSERT(lpRectParam != NULL);

            layout.recta.Left = layout.recta.Left + lpRectParam.Left
            layout.recta.Top = layout.recta.Top + lpRectParam.Top
            layout.recta.Right = layout.recta.Right - lpRectParam.Right
            layout.recta.Bottom = layout.recta.Bottom - lpRectParam.Bottom
        End If

        ' reposition the window
        Call Custom.CalcWindowRect(pLeftOver, layout.recta)
        Call AfxRepositionWindow(layout, hWndLeftOver, layout.recta)
    End If

    ' move and resize all the windows at once!
    If ((layout.hDWP = 0) Or (Not CBool(EndDeferWindowPos(layout.hDWP)))) Then
        Debug.Print "Warning: DeferWindowPos failed - low system resources.\n"
    End If

End Function


Private Property Get m_hWnd() As Long
    m_hWnd = hWnd
End Property

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

Private Property Get m_hWndOwner() As Long
    m_hWndOwner = hWndOwner
End Property

Private Property Let m_hWndOwner(RHS As Long)
    hWndOwner = RHS
End Property

Friend Function CalcWindowRect(lpClientRect As RECT, nAdjustType As Long)
    Dim dwExStyle As Long
    dwExStyle = GetExStyle(m_hWnd)
    If (nAdjustType = 0) Then
        dwExStyle = dwExStyle And (Not WS_EX_CLIENTEDGE)
    End If

    Call AdjustWindowRectEx(lpClientRect, GetStyle(m_hWnd), False, dwExStyle)
End Function


Private Function PreCreateWindow(ByRef cs As CREATESTRUCT, obj As Object)
    Select Case TypeName(obj)
        Case "CControlBar"
            Dim cb As CControlBar
            Set cb = obj
            Call cb.PreCreateWindow(cs)
            Set cb = Nothing
        Case "CMiniFrameWnd"
            Dim mdf As CMiniFrameWnd
            Set mdf = obj
            Call mdf.PreCreateWindow(cs)
            Set mdf = Nothing
    End Select
End Function


Friend Function ActivateTopParent()
    ' special activate logic for floating toolbars and palettes
    Dim pTopLevel As Long
    pTopLevel = GetTopLevelParent()

    Dim pActiveWnd As Long
    pActiveWnd = GetForegroundWindow()

    If (pActiveWnd = 0 Or Not ((pActiveWnd = m_hWnd) Or CBool(IsChild(pActiveWnd, m_hWnd)))) Then
        ' clicking on floating frame when it does not have
        ' focus itself -- activate the toplevel frame instead.
        Call SetForegroundWindow(pTopLevel)
    End If

End Function


Friend Function GetTopLevelParent() As Long
    'if (GetSafeHwnd() == NULL) // no Window attached
    '    return NULL;

    'ASSERT_VALID(this);

    Dim hWndParent As Long
    hWndParent = m_hWnd

    Dim hWndT As Long

    hWndT = AfxGetParentOwner(hWndParent)
    Do While (hWndT <> 0)
        hWndParent = hWndT
        hWndT = AfxGetParentOwner(hWndParent)
    Loop

    GetTopLevelParent = hWndParent
End Function


Private Function AfxGetParentOwner(Wnd As Long) As Long
    ' check for permanent-owned window first
    Dim pWnd As Long
    pWnd = Wnd

    If (pWnd <> 0) Then
        AfxGetParentOwner = GetOwner(pWnd)
        Exit Function
    End If

    ' otherwise, return parent in the Windows sense
    If CBool(GetWindowLong(Wnd, GWL_STYLE) And WS_CHILD) Then
        AfxGetParentOwner = GetParent(Wnd)
    Else
        AfxGetParentOwner = GetWindow(Wnd, GW_OWNER)
    End If
End Function


Private Function GetOwner(pWnd As Long) As Long
    GetOwner = GetParent(pWnd)
End Function


Friend Function GetTopLevelFrame(obj As Object) As Long

    If hWnd = 0 Then            ' no Window attached
        GetTopLevelFrame = 0
        Exit Function
    End If

    'ASSERT_VALID(this);

    Dim pFrameWnd As Long
    pFrameWnd = Me.hWnd

    If (Not obj.IsFrameWnd()) Then
        pFrameWnd = GetParentFrameMFC(m_hWnd)
    End If

    If (pFrameWnd <> 0) Then
        Dim pTemp As Long

        pTemp = GetParentFrameMFC(pFrameWnd)
        Do While (pTemp <> 0)
            pFrameWnd = pTemp
            pTemp = GetParentFrameMFC(pFrameWnd)
        Loop
    End If

    GetTopLevelFrame = pFrameWnd
End Function


Friend Function IsWindowEnabled() As Boolean
    IsWindowEnabled = True
End Function


Friend Function OnMouseActivate(pDesktopWnd As Long, nHitTest As Integer, nMsg As Integer) As Long
    OnMouseActivate = DefWindowProc(m_hWnd, WM_MOUSEACTIVATE, pDesktopWnd, MAKELONG(nHitTest, nMsg))
End Function

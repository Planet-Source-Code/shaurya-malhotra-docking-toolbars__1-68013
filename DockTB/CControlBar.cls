VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CControlBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CControlBar.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CControlBar class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CWnd
Private Wnd As New CWnd

Private m_bFloating As Boolean
Private m_bLayoutQuery As Boolean
Private m_bAutoDelete As Boolean

Private m_cxLeftBorder As Long, m_cxRightBorder As Long
Private m_cyTopBorder As Long, m_cyBottomBorder As Long
Private m_cxDefaultGap As Long          ' default gap value
Private m_nMRUWidth As Long             ' for dynamic resizing

Private m_nCount As Long

Enum StateFlags
    delayHide = 1
    delayShow = 2
    tempHide = 4
    statusSet = 8
End Enum

Private m_nStateFlags As Long
Private m_pDockContext As CDockContext  ' used during dragging
Private m_pDockSite As CFrame           ' current dock site, if dockable
Private m_pDockBar As CDockBar          ' current dock bar, if dockable

Private m_dwDockStyle As Long           ' indicates how bar can be docked
Public arrBars As New CPtrArray         ' each element is a CControlBar


Public Property Let pDockContext(ByVal RHS As CDockContext)
    Set m_pDockContext = RHS
End Property

Public Property Get pDockContext() As CDockContext
    Set pDockContext = m_pDockContext
End Property

Public Property Let bFloating(ByVal RHS As Boolean)
    m_bFloating = RHS
End Property

Public Property Get bFloating() As Boolean
    bFloating = m_bFloating
End Property

Public Property Let bLayoutQuery(ByVal RHS As Boolean)
    m_bLayoutQuery = RHS
End Property

Public Property Get bLayoutQuery() As Boolean
    bLayoutQuery = m_bLayoutQuery
End Property

Public Property Let cxLeftBorder(ByVal RHS As Long)
    m_cxLeftBorder = RHS
End Property

Public Property Get cxLeftBorder() As Long
    cxLeftBorder = m_cxLeftBorder
End Property

Public Property Let cxRightBorder(ByVal RHS As Long)
    m_cxRightBorder = RHS
End Property

Public Property Get cxRightBorder() As Long
    cxRightBorder = m_cxRightBorder
End Property

Public Property Let cyTopBorder(ByVal RHS As Long)
    m_cyTopBorder = RHS
End Property

Public Property Get cyTopBorder() As Long
    cyTopBorder = m_cyTopBorder
End Property

Public Property Let cyBottomBorder(ByVal RHS As Long)
    m_cyBottomBorder = RHS
End Property

Public Property Get cyBottomBorder() As Long
    cyBottomBorder = m_cyBottomBorder
End Property

Public Property Let cxDefaultGap(ByVal RHS As Long)
    m_cxDefaultGap = RHS
End Property

Public Property Get cxDefaultGap() As Long
    cxDefaultGap = m_cxDefaultGap
End Property

Public Property Let nMRUWidth(ByVal RHS As Long)
    m_nMRUWidth = RHS
End Property

Public Property Get nMRUWidth() As Long
    nMRUWidth = m_nMRUWidth
End Property

Public Property Let nCount(ByVal RHS As Long)
    m_nCount = RHS
End Property

Public Property Get nCount() As Long
    nCount = m_nCount
End Property

Public Property Let nStateFlags(ByVal RHS As Long)
    m_nStateFlags = RHS
End Property

Public Property Get nStateFlags() As Long
    nStateFlags = m_nStateFlags
End Property

Public Property Let dwStyle(ByVal RHS As Long)
    Wnd.dwStyle = RHS
End Property

Public Property Get dwStyle() As Long
    dwStyle = Wnd.dwStyle
End Property

Public Property Let pDockSite(ByVal RHS As CFrame)
    Set m_pDockSite = RHS
End Property

Public Property Get pDockSite() As CFrame
    Set pDockSite = m_pDockSite
End Property

Public Property Let pDockBar(ByVal RHS As CDockBar)
    Set m_pDockBar = RHS
End Property

Public Property Get pDockBar() As CDockBar
    Set pDockBar = m_pDockBar
End Property

Public Property Let dwDockStyle(ByVal RHS As Long)
    m_dwDockStyle = RHS
End Property

Public Property Get dwDockStyle() As Long
    dwDockStyle = m_dwDockStyle
End Property

Private Property Let CWnd_dwStyle(ByVal RHS As Long)
    Wnd.dwStyle = RHS
End Property

Private Property Get CWnd_dwStyle() As Long
    CWnd_dwStyle = Wnd.dwStyle
End Property

Private Property Let CWnd_hWnd(ByVal RHS As Long)
'
End Property

Private Property Get CWnd_hWnd() As Long
    CWnd_hWnd = Wnd.hWnd
End Property

Public Property Get hWnd() As Long
    hWnd = Wnd.hWnd
End Property
Friend Property Let hWnd(RHS As Long)
    Wnd.hWnd = RHS
End Property

Private Sub Class_Initialize()
    Call Initialize
End Sub

Private Function CWnd_CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long, Optional obj As Object = Nothing) As Boolean
    CWnd_CreateEx = Wnd.CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam)
End Function

Public Function CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long) As Boolean
    CreateEx = CWnd_CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam, this)
End Function

Friend Function Create(lpszClassName As String, lpszWindowName As String, dwStyle As Long, recta As RECT, pParentWnd As Long, nID As Long, Optional pContext As Long = 0) As Boolean
    Create = Wnd.Create(lpszClassName, lpszWindowName, dwStyle, recta, pParentWnd, nID, pContext, this)
End Function


Friend Sub SetBorders(cxLeft As Long, cyTop As Long, cxRight As Long, cyBottom As Long)
    'ASSERT(cxLeft >= 0);
    'ASSERT(cyTop >= 0);
    'ASSERT(cxRight >= 0);
    'ASSERT(cyBottom >= 0);

    m_cxLeftBorder = cxLeft
    m_cyTopBorder = cyTop
    m_cxRightBorder = cxRight
    m_cyBottomBorder = cyBottom
End Sub

Friend Function Invalidate(Optional bErase As Boolean = True) As Variant
    Wnd.Invalidate (bErase)
End Function

Friend Function CalcInsideRect(ByRef recta As RECT, bHorz As Boolean)   'const
    'ASSERT_VALID(this);
    Dim dwStyle As Long
    dwStyle = m_dwStyle

    If ((dwStyle And CBRS_BORDER_LEFT) <> 0) Then
        recta.Left = recta.Left + cxBorder2
    End If

    If (dwStyle And CBRS_BORDER_TOP) <> 0 Then
        recta.Top = recta.Top + cyBorder2
    End If

    If (dwStyle And CBRS_BORDER_RIGHT) <> 0 Then
        recta.Right = recta.Right - cxBorder2
    End If

    If (dwStyle And CBRS_BORDER_BOTTOM) <> 0 Then
        recta.Bottom = recta.Bottom - cyBorder2
    End If

    ' inset the top and bottom.
    If (bHorz = True) Then
        recta.Left = recta.Left + m_cxLeftBorder
        recta.Top = recta.Top + m_cyTopBorder
        recta.Right = recta.Right - m_cxRightBorder
        recta.Bottom = recta.Bottom - m_cyBottomBorder
        
        If ((m_dwStyle And (CBRS_GRIPPER Or CBRS_FLOATING)) = CBRS_GRIPPER) Then
            recta.Left = recta.Left + CX_BORDER_GRIPPER + CX_GRIPPER + CX_BORDER_GRIPPER
        End If
    Else
        recta.Left = recta.Left + m_cyTopBorder
        recta.Top = recta.Top + m_cxLeftBorder
        recta.Right = recta.Right - m_cyBottomBorder
        recta.Bottom = recta.Bottom - m_cxRightBorder

        If ((m_dwStyle And (CBRS_GRIPPER Or CBRS_FLOATING)) = CBRS_GRIPPER) Then
            recta.Top = recta.Top + CY_BORDER_GRIPPER + CY_GRIPPER + CY_BORDER_GRIPPER
        End If
    End If

End Function


Friend Function CalcFixedLayout(bStretch As Boolean, bHorz As Boolean) As Size
    Dim sizea As Size
    If (bStretch And bHorz) Then
        sizea.cx = 32767
    Else
        sizea.cx = 0
    End If
    
    If (bStretch And (Not bHorz)) Then
        sizea.cy = 32767
    Else
        sizea.cy = 0
    End If

    CalcFixedLayout = sizea
End Function


Public Function IsVisible() As Boolean
    If (m_nStateFlags And delayHide) <> 0 Then
        IsVisible = False
        Exit Function
    End If

    If (((m_nStateFlags And delayShow) <> 0) Or ((GetStyle(m_hWnd) And WS_VISIBLE) <> 0)) Then
        IsVisible = True
        Exit Function
    End If

    IsVisible = False
End Function


Private Property Let m_dwStyle(ByVal RHS As Long)
    dwStyle = RHS
End Property

Private Property Get m_dwStyle() As Long
    m_dwStyle = dwStyle
End Property



Public Function SetBarStyle(dwStyle_param As Long)
    If (m_dwStyle <> dwStyle_param) Then
        Dim dwOldStyle As Long
        dwOldStyle = m_dwStyle
        m_dwStyle = dwStyle_param
        Call OnBarStyleChange(dwOldStyle, dwStyle_param)
    End If
End Function

Private Function OnBarStyleChange(dwOldStyle As Long, dwNewStyle As Long)
'
End Function

Private Function Initialize()
    ' no elements contained in the control bar yet
    m_nCount = 0
    'm_pData = NULL;

    ' set up some default border spacings
    m_cxLeftBorder = 6
    m_cxRightBorder = 6
    m_cxDefaultGap = 2
    m_cyTopBorder = 1
    m_cyBottomBorder = 1
    'm_bAutoDelete = FALSE;
    'm_hWndOwner = NULL;
    m_nStateFlags = 0
    Set m_pDockSite = Nothing
    Set m_pDockBar = Nothing
    Set m_pDockContext = Nothing
    m_dwStyle = 0
    m_dwDockStyle = 0
    m_nMRUWidth = 32767
End Function

Private Property Let CWnd_hWndOwner(ByVal RHS As Long)
    Wnd.hWndOwner = RHS
End Property

Private Property Get CWnd_hWndOwner() As Long
    CWnd_hWndOwner = Wnd.hWndOwner
End Property

Public Property Let hWndOwner(ByVal RHS As Long)
    CWnd_hWndOwner = RHS
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = CWnd_hWndOwner
End Property


Friend Function OnSizeParent(wParam As Long, ByRef lParam As AFX_SIZEPARENTPARAMS, Optional pBar As Variant) As Long
    Dim lpLayout As AFX_SIZEPARENTPARAMS
    lpLayout = lParam
    
    Dim dwStyle As Long
    dwStyle = RecalcDelayShow(lpLayout, m_hWnd)
    
    dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    If (((dwStyle And WS_VISIBLE) <> 0) And ((dwStyle And CBRS_ALIGN_ANY) <> 0)) <> 0 Then
        ' align the control bar
        Dim recta As RECT
        Call CopyRect(recta, lpLayout.recta)

        Dim sizeAvail As Size
        sizeAvail = CSize(recta)      ' maximum size available

        ' get maximum requested size
        Dim dwMode As Long

        If (lpLayout.bStretch) <> 0 Then dwMode = LM_STRETCH Else dwMode = 0
        
        If ((((GetWindowLong(m_hWnd, GWL_STYLE) And CBRS_ALL) And CBRS_SIZE_DYNAMIC) <> 0) And (((GetWindowLong(m_hWnd, GWL_STYLE) And CBRS_ALL) And CBRS_FLOATING) <> 0)) <> 0 Then
            dwMode = dwMode Or (LM_HORZ Or LM_MRUWIDTH)
        ElseIf (dwStyle And CBRS_ORIENT_HORZ) <> 0 Then
            dwMode = dwMode Or (LM_HORZ Or LM_HORZDOCK)
        Else
            dwMode = dwMode Or LM_VERTDOCK
        End If

        Dim sizea As Size

        If (IsObject(pBar)) Then
                sizea = CalcDynamicLayout(-1, dwMode, pBar)
        Else
                sizea = CalcDynamicLayout(-1, dwMode)
        End If

        sizea.cx = GetMin(sizea.cx, sizeAvail.cx)
        sizea.cy = GetMin(sizea.cy, sizeAvail.cy)

        If (dwStyle And CBRS_ORIENT_HORZ) <> 0 Then
            lpLayout.sizeTotal.cy = lpLayout.sizeTotal.cy + sizea.cy
            lpLayout.sizeTotal.cx = GetMax(lpLayout.sizeTotal.cx, sizea.cx)
            
            If (dwStyle And CBRS_ALIGN_TOP) <> 0 Then
                lpLayout.recta.Top = lpLayout.recta.Top + sizea.cy
            ElseIf (dwStyle And CBRS_ALIGN_BOTTOM) <> 0 Then
                recta.Top = recta.Bottom - sizea.cy
                lpLayout.recta.Bottom = lpLayout.recta.Bottom - sizea.cy
            End If
        ElseIf (dwStyle And CBRS_ORIENT_VERT) <> 0 Then
            lpLayout.sizeTotal.cx = lpLayout.sizeTotal.cx + sizea.cx
            lpLayout.sizeTotal.cy = GetMax(lpLayout.sizeTotal.cy, sizea.cy)

            If (dwStyle And CBRS_ALIGN_LEFT) <> 0 Then
                lpLayout.recta.Left = lpLayout.recta.Left + sizea.cx
            ElseIf (dwStyle And CBRS_ALIGN_RIGHT) <> 0 Then
                recta.Left = recta.Right - sizea.cx
                lpLayout.recta.Right = lpLayout.recta.Right - sizea.cx
            End If
        Else
            Assert ("Assertion failed!")    ' can never happen
        End If

        recta.Right = recta.Left + sizea.cx
        recta.Bottom = recta.Top + sizea.cy

        ' only resize the window if doing layout and not just rect query
        If (lpLayout.hDWP <> 0) Then
            If (IsObject(pBar)) Then
                If (pBar.IsDockBar()) Then
                        Call AfxRepositionWindow(lpLayout, m_hWnd, recta)
                End If
            End If
        End If

    End If

    OnSizeParent = 0
    lParam = lpLayout
End Function


Friend Property Get m_hWnd() As Long
    m_hWnd = Wnd.hWnd
End Property


Friend Function CalcDynamicLayout(x As Integer, nMode As Long, Optional pBar As Variant) As Size
    If IsObject(pBar) Then
        Dim tmp As Size
        Dim db As CDockBar
        Set db = pBar
        CalcDynamicLayout = db.CalcFixedLayout(nMode And LM_STRETCH, nMode And LM_HORZ)
        Set db = Nothing
    Else
        CalcDynamicLayout = CalcFixedLayout(nMode And LM_STRETCH, nMode And LM_HORZ)
    End If
End Function




Private Function RecalcDelayShow(lpLayout As AFX_SIZEPARENTPARAMS, Optional m_hWnd As Long = 0) As Long
    ' resize and reposition this control bar based on styles
    Dim dwStyle As Long
    dwStyle = (m_dwStyle And (CBRS_ALIGN_ANY Or CBRS_BORDER_ANY)) Or _
        (GetWindowLong(m_hWnd, GWL_STYLE) And WS_VISIBLE)
        
    dwStyle = GetWindowLong(m_hWnd, GWL_STYLE)

    ' handle delay hide/show
    If (m_nStateFlags And (delayHide Or delayShow)) = 0 Then
        Dim swpFlags As Long
        swpFlags = 0

        If (m_nStateFlags And delayHide) <> 0 Then
            'ASSERT((m_nStateFlags & delayShow) == 0);
            If (dwStyle And WS_VISIBLE) <> 0 Then
                swpFlags = SWP_HIDEWINDOW
            End If
        Else
            'ASSERT(m_nStateFlags & delayShow);
            If ((dwStyle And WS_VISIBLE) = 0) <> 0 Then
                swpFlags = SWP_SHOWWINDOW
            End If
        End If

        If (swpFlags <> 0) Then
            ' make the window seem visible/hidden
            dwStyle = dwStyle Xor WS_VISIBLE

            If (lpLayout.hDWP <> 0) Then
                ' clear delay flags
                m_nStateFlags = m_nStateFlags And (Not (delayShow Or delayHide))
                ' hide/show the window if actually doing layout
                lpLayout.hDWP = DeferWindowPos(lpLayout.hDWP, m_hWnd, 0, _
                    0, 0, 0, 0, swpFlags Or _
                    SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
            End If
        Else
            ' clear delay flags -- window is already in correct state
            m_nStateFlags = m_nStateFlags And (Not (delayShow Or delayHide)) '**CHECK AGAIN**'
        End If
    End If

    RecalcDelayShow = dwStyle           ' return new style
End Function

Friend Function OnCreate(obj As Object) As Integer
    Dim pFrameWnd As CFrame
    Set pFrameWnd = GetParentFrame(m_hWnd)

    If (pFrameWnd.IsFrameWnd()) Then
        Set m_pDockSite = pFrameWnd
        Call m_pDockSite.AddControlBar(obj)
    End If

    OnCreate = 0
End Function


Private Property Get this() As CControlBar
   Set this = Me
End Property

Friend Function GetBarStyle() As Long
    GetBarStyle = m_dwStyle
End Function

Friend Function IsWindowVisible() As Boolean
    IsWindowVisible = CBool(Globals.IsWindowVisible(m_hWnd))
End Function

Friend Function SetWindowPos(hWndInsertAfter As Long, x As Long, y As Long, cx As Long, cy As Long, wFlags As Long) As Long
    SetWindowPos = Globals.SetWindowPos(m_hWnd, hWndInsertAfter, x, y, cx, cy, wFlags)
End Function

Friend Function GetDockedControlBar(nPos As Long) As Object
    Dim pResult As Object
    Set pResult = arrBars.GetItem(nPos)

    If IsObject(pResult) Then
        Set GetDockedControlBar = pResult
    Else
        Set GetDockedControlBar = Nothing
    End If
End Function

Friend Property Get GetParent() As Long
    GetParent = Globals.GetParent(m_hWnd)
End Property

Friend Function SetParent(newParent As Long)
    Call Globals.SetParent(m_hWnd, newParent)
End Function

Public Property Let bAutoDelete(ByVal RHS As Boolean)
    m_bAutoDelete = RHS
End Property
Public Property Get bAutoDelete() As Boolean
    bAutoDelete = m_bAutoDelete
End Property

Friend Function PreCreateWindow(ByRef cs As CREATESTRUCT) As Boolean
    ' force clipsliblings (otherwise will cause repaint problems)
    cs.style = cs.style Or WS_CLIPSIBLINGS

    ' default border style translation for Win4
    '  (you can turn off this translation by setting CBRS_BORDER_3D)

    If ((m_dwStyle And CBRS_BORDER_3D) = 0) Then
        Dim dwNewStyle As Long
        dwNewStyle = 0
        
        Select Case (m_dwStyle And (CBRS_BORDER_ANY Or CBRS_ALIGN_ANY))
            Case CBRS_LEFT:
                dwNewStyle = CBRS_BORDER_TOP Or CBRS_BORDER_BOTTOM
            Case CBRS_TOP:
                dwNewStyle = CBRS_BORDER_TOP
            Case CBRS_RIGHT:
                dwNewStyle = CBRS_BORDER_TOP Or CBRS_BORDER_BOTTOM
            Case CBRS_BOTTOM:
                dwNewStyle = CBRS_BORDER_BOTTOM
        End Select

        ' set new style if it matched one of the predefined border types
        If (dwNewStyle <> 0) Then
            m_dwStyle = m_dwStyle And (Not (CBRS_BORDER_ANY))
            m_dwStyle = m_dwStyle Or (dwNewStyle Or CBRS_BORDER_3D)
        End If
    End If

    PreCreateWindow = True
End Function

Private Property Let CWnd_nFlags(ByVal RHS As Long)
'
End Property

Private Property Get CWnd_nFlags() As Long
'
End Property

Friend Function OnMouseActivate(pDesktopWnd As Long, nHitTest As Integer, nMsg As Integer) As Long
    ' call default when toolbar is not floating
    If (Not IsFloating()) Then
        OnMouseActivate = Wnd.OnMouseActivate(pDesktopWnd, nHitTest, nMsg)
    End If

    ' special behavior when floating
    Call ActivateTopParent

    OnMouseActivate = MA_NOACTIVATE     ' activation already done
End Function


Private Function IsFloating() As Boolean
    If (IsDockBar()) Then
        IsFloating = bFloating
    Else
        IsFloating = (Not (m_pDockBar Is Nothing)) And (m_pDockBar.m_bFloating)
    End If
End Function

Private Function IsDockBar() As Boolean
    IsDockBar = False
End Function

Private Function ActivateTopParent()
    Call Wnd.ActivateTopParent
End Function


Friend Function EraseNonClient(Optional obj As Object = Nothing)
    ' get window DC that is clipped to the non-client area
    Dim dc As Long
    dc = GetWindowDC(hWnd)
    
    Dim rectClient As RECT
    Call GetClientRect(hWnd, rectClient)

    Dim rectWindow As RECT
    Call GetWindowRect(hWnd, rectWindow)
    
    Call MyScreenToClient(hWnd, rectWindow)
    Call OffsetRect(rectClient, -rectWindow.Left, -rectWindow.Top)  'rectClient.
    Call ExcludeClipRect(dc, rectClient.Left, rectClient.Top, rectClient.Right, rectClient.Bottom)       'dc.

    ' draw borders in non-client area
    Call OffsetRect(rectWindow, -rectWindow.Left, -rectWindow.Top) 'rectWindow.

    Call DrawBorders(ByVal dc, rectWindow, obj)

    Call IntersectClipRect(ByVal dc, rectWindow.Left, rectWindow.Top, rectWindow.Right, rectWindow.Bottom)    'dc.
    Call SendMessage(hWnd, WM_ERASEBKGND, ByVal dc, 0)
    
    ' draw gripper in non-client area
    Call DrawGripper(dc, rectWindow)

    Call ReleaseDC(hWnd, dc)
End Function


Private Function DrawBorders(ByVal pDC As Long, ByRef recta As RECT, Optional obj As Object = Nothing)
    'ASSERT_VALID(this);
    'ASSERT_VALID(pDC);

    Dim dwStyle As Long
    dwStyle = m_dwStyle
    
    If Not ((dwStyle And CBRS_BORDER_ANY) <> 0) Then
        Exit Function
    End If

    ' prepare for dark lines
    'ASSERT(rect.top == 0 && rect.left == 0);
    Dim rect1 As RECT, rect2 As RECT
    rect1 = recta
    rect2 = recta
        
    'COLORREF clr = afxData.bWin4 ? afxData.clrBtnShadow : afxData.clrWindowFrame;
    Dim clr As Long
    clr = clrBtnShadow

    ' draw dark line one pixel back/up
    If (dwStyle And CBRS_BORDER_3D) <> 0 Then
        rect1.Right = rect1.Right - CX_BORDER
        rect1.Bottom = rect1.Bottom - CY_BORDER
    End If
    
    If (dwStyle And CBRS_BORDER_TOP) <> 0 Then
        rect2.Top = rect2.Top + cyBorder2
    End If
    If (dwStyle And CBRS_BORDER_BOTTOM) <> 0 Then
        rect2.Bottom = rect2.Bottom - cyBorder2
    End If

    ' draw left and top
    If (dwStyle And CBRS_BORDER_LEFT) <> 0 Then
        Call FillSolidRect(pDC, 0, rect2.Top, CX_BORDER, (rect2.Bottom - rect2.Top), clr)
    End If
    
    If (dwStyle And CBRS_BORDER_TOP) <> 0 Then
        Call FillSolidRect(pDC, 0, 0, recta.Right, CY_BORDER, clr)
    End If

    ' draw right and bottom
    If (dwStyle And CBRS_BORDER_RIGHT) <> 0 Then
        Call FillSolidRect(pDC, rect1.Right, rect2.Top, -CX_BORDER, (rect2.Bottom - rect2.Top), clr)
    End If
    
    If (dwStyle And CBRS_BORDER_BOTTOM) <> 0 Then
        Call FillSolidRect(pDC, 0, rect1.Bottom, recta.Right, -CY_BORDER, clr)
    End If

    If (dwStyle And CBRS_BORDER_3D) <> 0 Then
        ' prepare for hilite lines
        clr = clrBtnHilite

        ' draw left and top
        If (dwStyle And CBRS_BORDER_LEFT) <> 0 Then
            Call FillSolidRect(pDC, 1, rect2.Top, CX_BORDER, (rect2.Bottom - rect2.Top), clr)
        End If
        If (dwStyle And CBRS_BORDER_TOP) <> 0 Then
            Call FillSolidRect(pDC, 0, 1, recta.Right, CY_BORDER, clr)
        End If

        ' draw right and bottom
        If (dwStyle And CBRS_BORDER_RIGHT) <> 0 Then
            Call FillSolidRect(pDC, recta.Right, rect2.Top, -CX_BORDER, (rect2.Bottom - rect2.Top), clr)
        End If
        
        If (dwStyle And CBRS_BORDER_BOTTOM) <> 0 Then
            Call FillSolidRect(pDC, 0, recta.Bottom, recta.Right, -CY_BORDER, clr)
        End If
    End If

    If (dwStyle And CBRS_BORDER_LEFT) <> 0 Then
        recta.Left = recta.Left + cxBorder2
    End If
    If (dwStyle And CBRS_BORDER_TOP) <> 0 Then
        recta.Top = recta.Top + cyBorder2
    End If
    If (dwStyle And CBRS_BORDER_RIGHT) <> 0 Then
        recta.Right = recta.Right - cxBorder2
    End If
    If (dwStyle And CBRS_BORDER_BOTTOM) <> 0 Then
        recta.Bottom = recta.Bottom - cyBorder2
    End If
End Function



Private Function DrawGripper(ByVal pDC As Long, ByRef recta As RECT)
    ' only draw the gripper if not floating and gripper is specified
    If ((m_dwStyle And (CBRS_GRIPPER Or CBRS_FLOATING)) = CBRS_GRIPPER) Then

        ' draw the gripper in the border
        If (m_dwStyle And CBRS_ORIENT_HORZ) <> 0 Then
            Call Draw3dRect(pDC, recta.Left + CX_BORDER_GRIPPER, _
                recta.Top + m_cyTopBorder, _
                CX_GRIPPER, (recta.Bottom - recta.Top) - m_cyTopBorder - m_cyBottomBorder, _
                clrBtnHilite, clrBtnShadow)
        Else
            Call Draw3dRect(pDC, recta.Left + m_cyTopBorder, _
                recta.Top + CY_BORDER_GRIPPER, _
                (recta.Right - recta.Left) - m_cyTopBorder - m_cyBottomBorder, CY_GRIPPER, _
                clrBtnHilite, clrBtnShadow)
        End If
    End If
End Function


Friend Function OnWindowPosChanging(lpWndPos As WINDOWPOS)
    ' WINBUG: We call DefWindowProc here instead of CWnd::OnWindowPosChanging
    '  (which calls CWnd::Default, which calls through the super wndproc)
    '  because certain control bars that are system implemented (such as
    '  CToolBar with TBSTYLE_FLAT) do not implement WM_WINDOWPOSCHANGING
    '  correctly, causing repaint problems.  This code bypasses that whole
    '  mess.

    Call DefWindowProc(m_hWnd, WM_WINDOWPOSCHANGING, 0, VarPtr(lpWndPos))

    If (lpWndPos.flags And SWP_NOSIZE) <> 0 Then
        Exit Function
    End If

    ' invalidate borders on the right
    Dim recta As RECT
    
    Call GetWindowRect(m_hWnd, recta)
    
    Dim sizePrev  As Size
    sizePrev = CRECTtoSize(recta)
    
    Dim cx As Integer
    cx = lpWndPos.cx
    Dim cy As Integer
    cy = lpWndPos.cy
    
    If ((cx <> sizePrev.cx) And ((m_dwStyle And CBRS_BORDER_RIGHT) <> 0)) Then
        Call SetRect(recta, cx - cxBorder2, 0, cx, cy)
        Call InvalidateRect(m_hWnd, recta, True)
        Call SetRect(recta, sizePrev.cx - cxBorder2, 0, sizePrev.cx, cy)
        Call InvalidateRect(m_hWnd, recta, True)
    End If

    ' invalidate borders on the bottom
    If ((cy <> sizePrev.cy) And ((m_dwStyle And CBRS_BORDER_BOTTOM) <> 0)) Then
        Call SetRect(recta, 0, cy - cyBorder2, cx, cy)
        Call InvalidateRect(m_hWnd, recta, True)
        Call SetRect(recta, 0, sizePrev.cy - cyBorder2, cx, sizePrev.cy)
        Call InvalidateRect(m_hWnd, recta, True)
    End If
End Function


Friend Function OnDestroy()
    If Not (m_pDockSite Is Nothing) Then
        Call m_pDockSite.RemoveControlBar(this)
        Set m_pDockSite = Nothing
    End If
End Function


Friend Function OnLButtonDblClk(nFlags As Long, pt As POINTAPI, Optional obj As CToolbar = Nothing)
    ' only toggle docking if clicked in "void" space
    Dim tmp As tagTOOLINFO
    If (Not (m_pDockBar Is Nothing) And (obj.OnToolHitTest(pt, False, tmp) = -1)) Then
        ' start the drag
        'ASSERT(m_pDockContext != NULL);
        Call m_pDockContext.ToggleDocking
    Else
        'CWnd::OnLButtonDblClk(nFlags, pt);
    End If
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CFrame.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CFrame class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CWnd

Public Event OnCommand(ID As Long)

'****IMPORTANT : SUBCLASSING DATA****
'***********DO NOT MODIFY!!**********
Private Wnd As New CWnd
Private mWndProcOrg As Long
Private mHWndSubClassed As Long
'******END OF SUBCLASSING DATA*******

Private bInRecalcLayout As Boolean              ' avoid recursion in RecalcLayout
Private m_nIdleFlags As Long                    ' set of bit flags for idle processing
Private m_listControlBars As New Collection
Private m_rectBorder As RECT                    ' for OLE border space negotiation

Private dwDockBarMap(0 To 3, 0 To 1) As Long    ' initialized in Class_Initialize

Private Enum IdleFlags
    idleMenu = 1
    idleTitle = 2
    idleNotify = 4
    idleLayout = 8
End Enum


Friend Property Get m_bInRecalcLayout() As Boolean
    m_bInRecalcLayout = bInRecalcLayout
End Property

Friend Property Let m_bInRecalcLayout(RHS As Boolean)
    bInRecalcLayout = RHS
End Property

Private Sub Class_Initialize()
dwDockBarMap(0, 0) = AFX_IDW_DOCKBAR_TOP
dwDockBarMap(0, 1) = CBRS_TOP

dwDockBarMap(1, 0) = AFX_IDW_DOCKBAR_BOTTOM
dwDockBarMap(1, 1) = CBRS_BOTTOM

dwDockBarMap(2, 0) = AFX_IDW_DOCKBAR_LEFT
dwDockBarMap(2, 1) = CBRS_LEFT

dwDockBarMap(3, 0) = AFX_IDW_DOCKBAR_RIGHT
dwDockBarMap(3, 1) = CBRS_RIGHT
End Sub

Friend Function AddControlBar(pBar As Object)
    m_listControlBars.Add pBar
End Function

Private Function CWnd_CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long, Optional obj As Object = Nothing) As Boolean
    CWnd_CreateEx = Wnd.CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam, obj)
End Function

Private Property Let CWnd_dwStyle(ByVal RHS As Long)
    Wnd.dwStyle = RHS
End Property

Private Property Get CWnd_dwStyle() As Long
    CWnd_dwStyle = Wnd.dwStyle
End Property

Private Property Let CWnd_hWnd(ByVal RHS As Long)
    Wnd.hWnd = RHS
End Property

Private Property Get CWnd_hWnd() As Long
    CWnd_hWnd = Wnd.hWnd
End Property

Public Function CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long, Optional obj As Object = Nothing) As Boolean
    CreateEx = CWnd_CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam, obj)
End Function

Public Property Get hWnd() As Long
    hWnd = Wnd.hWnd
End Property

Public Property Let hWnd(RHS As Long)
    Wnd.hWnd = RHS
End Property

Friend Function Create(lpszClassName As String, lpszWindowName As String, dwStyle As Long, recta As RECT, pParentWnd As Long, nID As Long, pContext As Long) As Boolean
    Create = Wnd.Create(lpszClassName, lpszWindowName, dwStyle, recta, pParentWnd, nID, pContext)
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


Friend Property Get m_hWnd() As Long
    m_hWnd = Wnd.hWnd
End Property

Public Function EnableDocking(dwDockStyle As Long)
    ' must be CBRS_ALIGN_XXX or CBRS_FLOAT_MULTI only
    'ASSERT((dwDockStyle & ~(CBRS_ALIGN_ANY|CBRS_FLOAT_MULTI)) == 0);

    'm_pFloatingFrameClass = RUNTIME_CLASS(CMiniDockFrameWnd);

    Dim i As Integer
    For i = 0 To 3
        If CBool(dwDockBarMap(i, 1) And dwDockStyle And CBRS_ALIGN_ANY) Then
            Dim pDock As CDockBar
            Set pDock = GetControlBar(dwDockBarMap(i, 0))

            If (pDock Is Nothing) Then
                Set pDock = New CDockBar
                If (Not pDock.Create(this.m_hWnd, _
                    WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_CHILD Or WS_VISIBLE Or _
                        dwDockBarMap(i, 1), dwDockBarMap(i, 0))) Then
                    MsgBox "Low system resources. Could not create dockbar", vbCritical             'AfxThrowResourceException();
                End If

            End If
        End If
    Next i

End Function


Friend Function GetControlBar(nID As Long) As CControlBar
    If CBool(nID = 0) Then
        Set GetControlBar = Nothing
        Exit Function
    End If

    Dim pos As Integer
    pos = 1

    Do While (pos <= m_listControlBars.Count)
        Dim pBar As CControlBar
        Set pBar = m_listControlBars.item(pos)

        'ASSERT(pBar != NULL);

        If (AfxGetDlgCtrlID(pBar.hWnd) = nID) Then
            'ASSERT_KINDOF(CControlBar, pBar);
            Set GetControlBar = pBar
            Exit Function
        End If
        pos = pos + 1
    Loop

    Set GetControlBar = Nothing
End Function


Private Property Get this() As CFrame
   Set this = Me
End Property

Public Function IsFrameWnd() As Boolean
    IsFrameWnd = True
End Function


Friend Function WindowPROC(ByVal hWnd As Long, _
         ByVal uMsg As Long, ByVal wParam As Long, _
         lParam As Long) As Long

    Select Case uMsg

        Case WM_NCACTIVATE
            WindowPROC = OnNcActivate(CBool(wParam))

        Case WM_ACTIVATE
            Dim tmp As Long
            Call CopyMemory(tmp, VarPtr(lParam), Len(tmp))
            Call OnActivate(GetLoWord(wParam), tmp, CBool(GetHiWord(wParam)))

        Case WM_SIZE
            WindowPROC = OnSize(wParam, 0, 0) 'TEMP'

        Case WM_COMMAND
            RaiseEvent OnCommand(wParam)

        Case Else
            WindowPROC = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, lParam)

    End Select

End Function


Friend Sub Subclass()
    If mWndProcOrg Then Exit Sub
    mWndProcOrg = SetWindowLong(m_hWnd, GWL_WNDPROC, _
                        AddressOf FM_SubWndProc)
    mHWndSubClassed = m_hWnd
    Call SetWindowLong(m_hWnd, GWL_USERDATA, ObjPtr(Me))
End Sub


Private Sub UnSubClass()
    If mWndProcOrg = 0 Then Exit Sub
    SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg
    mWndProcOrg = 0
End Sub


Private Sub Class_Terminate()
    Call UnSubClass
End Sub


Private Function OnSize(nType As Long, cx As Long, cy As Long)
    Call Wnd.OnSize(nType, cx, cy)          ' important for MDI Children

    If (nType <> SIZE_MINIMIZED) Then
        Call RecalcLayout
    End If
End Function


Friend Function RecalcLayout(Optional bNotify As Boolean = True, Optional obj As CMiniDockFrameWnd = Nothing)
    If (m_bInRecalcLayout) Then
        Exit Function
    End If

    m_bInRecalcLayout = True

    ' clear idle flags for recalc layout if called elsewhere
    If (m_nIdleFlags And idleNotify) <> 0 Then
        bNotify = True
    End If

    m_nIdleFlags = m_nIdleFlags And (Not (idleLayout Or idleNotify))

'#ifndef _AFX_NO_OLE_SUPPORT
'    // call the layout hook -- OLE support uses this hook
'    if (bNotify && m_pNotifyHook != NULL)
'        m_pNotifyHook->OnRecalcLayout();
'#End If

    ' reposition all the child windows (regardless of ID)
    If (GetStyle(m_hWnd) And FWS_SNAPTOBARS) = FWS_SNAPTOBARS Then
        Dim recta As RECT
        recta.Left = 0
        recta.Top = 0
        recta.Right = 32767
        recta.Bottom = 32767

        Call Wnd.RepositionBars(0, &HFFFF&, AFX_IDW_PANE_FIRST, reposQuery, _
              recta, recta, False)
        Call Wnd.RepositionBars(0, &HFFFF&, AFX_IDW_PANE_FIRST, reposExtra, _
            m_rectBorder, recta, True)

        If (obj Is Nothing) Then
            Call Custom.CalcWindowRect(m_hWnd, recta)
        Else
            Call obj.CalcWindowRect(recta)
        End If

        Call SetWindowPos(hWnd, 0, 0, 0, GetWidth(recta), GetHeight(recta), _
            SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOZORDER)
    Else
        Dim emptyRECT As RECT           ' because optional parameters cannot be user defined types
        Call Wnd.RepositionBars(0, &HFFFF&, AFX_IDW_PANE_FIRST, reposExtra, m_rectBorder, emptyRECT)   '0xffff=65535
    End If

    m_bInRecalcLayout = False
End Function


Friend Function CanDock(recta As RECT, dwDockStyle As Long, Optional ByRef ppDockBar As CDockBar = Nothing) As Long
    ' dwDockStyle -- allowable styles of bar
    ' don't allow to dock to floating unless multi is specified
    dwDockStyle = dwDockStyle And (CBRS_ALIGN_ANY Or CBRS_FLOAT_MULTI)

    If IsObject(ppDockBar) Then
        Set ppDockBar = Nothing
    End If

    Dim pos As Integer
    pos = 1

    Dim pDockBar As Object
    Do While (pos <= m_listControlBars.Count)

        Set pDockBar = Nothing
        Set pDockBar = m_listControlBars(pos)

        If Not (pDockBar.IsDockBar) Then GoTo Continue

        If (pDockBar.IsDockBar() And pDockBar.IsWindowVisible() And _
            (pDockBar.m_dwStyle And dwDockStyle And CBRS_ALIGN_ANY) <> 0 And _
            (Not pDockBar.m_bFloating Or _
                (dwDockStyle And pDockBar.m_dwStyle And CBRS_FLOAT_MULTI))) Then

            Dim rectBar As RECT
            Call GetWindowRect(pDockBar.m_hWnd, rectBar)

            If (GetWidth(rectBar) = 0) Then
                rectBar.Right = rectBar.Right + 1
            End If

            If (GetHeight(rectBar) = 0) Then
                rectBar.Bottom = rectBar.Bottom + 1
            End If

            Call IntersectRect(rectBar, rectBar, recta)

            If (Not CBool(IsRectEmpty(rectBar))) Then
                If IsObject(ppDockBar) Then
                    Set ppDockBar = pDockBar
                End If

                CanDock = pDockBar.m_dwStyle And dwDockStyle
                Exit Function
            End If

        End If

Continue:
    pos = pos + 1
    Loop
    'return 0;
    CanDock = 0

End Function


Friend Function DockControlBar1(pBar As Object, nDockBarID As Long, lpRect As RECT)
    Dim pDockBar As CDockBar
    If (nDockBarID = 0) Then
        Set pDockBar = Nothing
    Else
        Set pDockBar = GetControlBar(nDockBarID)
    End If

    Call DockControlBar(pBar, pDockBar, lpRect)
End Function


Friend Function DockControlBar(pBar As Object, pDockBar As CDockBar, lpRect As RECT)
    'ASSERT(pBar != NULL);
    ' make sure CControlBar::EnableDocking has been called
    'ASSERT(pBar->m_pDockContext != NULL);

    If (pDockBar Is Nothing) Then
        Dim i As Integer
        For i = 0 To 3
            If ((dwDockBarMap(i, 1) And CBRS_ALIGN_ANY) = _
                    (pBar.m_dwStyle And CBRS_ALIGN_ANY)) Then

                Set pDockBar = GetControlBar(dwDockBarMap(i, 0))
                'ASSERT(pDockBar != NULL);
                ' assert fails when initial CBRS_ of bar does not
                ' match available docking sites, as set by EnableDocking()
                Exit For
            End If
        Next i
    End If

    'ASSERT(pDockBar != NULL);
    'ASSERT(m_listControlBars.Find(pBar) != NULL);
    'ASSERT(pBar->m_pDockSite == this);
    ' if this assertion occurred it is because the parent of pBar was not initially
    ' this CFrameWnd when pBar's OnCreate was called
    ' i.e. this control bar should have been created with a different parent initially

    Call pDockBar.DockControlBar(pBar, lpRect)
End Function


Friend Function FloatControlBar(pBar As CToolbar, point As POINTAPI, dwStyle As Long)
    'ASSERT(pBar != NULL);

    ' if the bar is already floating and the dock bar only contains this
    ' bar and same orientation then move the window rather than recreating
    ' the frame
    If ((Not (pBar.m_pDockSite Is Nothing)) And (Not (pBar.m_pDockBar Is Nothing))) Then
        Dim pDockBar As CDockBar
        Set pDockBar = pBar.m_pDockBar
        
        'ASSERT_KINDOF(CDockBar, pDockBar);
        If (pDockBar.m_bFloating And (pDockBar.GetDockedCount() = 1) And _
                (dwStyle And pDockBar.m_dwStyle And CBRS_ALIGN_ANY) <> 0) Then

            Dim pDockFrame As CMiniDockFrameWnd
            Set pDockFrame = GetParentBar(pDockBar.m_hWnd)

            'ASSERT(pDockFrame != NULL);
            'ASSERT_KINDOF(CMiniDockFrameWnd, pDockFrame);

            Call pDockFrame.SetWindowPos(0, point.x, point.y, 0, 0, _
                SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
            Call pDockFrame.RecalcLayout(True)
            Call pDockFrame.UpdateWindow

            Exit Function
        End If
    End If

    If (pBar.m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0 Then
        dwStyle = dwStyle Or CBRS_SIZE_DYNAMIC
        If (dwStyle And CBRS_ORIENT_VERT) <> 0 Then
            dwStyle = dwStyle And (Not CBRS_ALIGN_ANY)
            dwStyle = dwStyle Or CBRS_ALIGN_TOP
        End If
    End If

    Set pDockFrame = CreateFloatingFrame(dwStyle)
    Call AddBar(pDockFrame)

    'ASSERT(pDockFrame != NULL);
    Call pDockFrame.SetWindowPos(0, point.x, point.y, 0, 0, _
        SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)

    If (pDockFrame.m_hWndOwner = 0) Then
        pDockFrame.m_hWndOwner = pBar.m_hWnd
    End If

    Set pDockBar = pDockFrame.m_wndDockBar
    'ASSERT(pDockBar != NULL);
    'ASSERT_KINDOF(CDockBar, pDockBar);

    'ASSERT(pBar->m_pDockSite == this);
    ' if this assertion occurred it is because the parent of pBar was not
    ' initially this CFrameWnd when pBar's OnCreate was called
    ' (this control bar should have been created with a different
    ' parent initially)

    Dim emptyRECT As RECT
    Call SetRectEmpty(emptyRECT)
    Call pDockBar.DockControlBar(pBar, emptyRECT)
    Call pDockFrame.RecalcLayout(True)

    If (GetWindowLong(pBar.m_hWnd, GWL_STYLE) And WS_VISIBLE) <> 0 Then
        Call pDockFrame.ShowWindow(SW_SHOWNA)
        Call pDockFrame.UpdateWindow
    End If

End Function

Friend Function DestroyWindow()
    Call Globals.DestroyWindow(m_hWnd)
End Function

Friend Function ShowWindow()
'
End Function


Friend Function CreateFloatingFrame(dwStyle As Long) As CMiniDockFrameWnd
    Dim pFrame As CMiniDockFrameWnd
    'ASSERT(m_pFloatingFrameClass != NULL);

    Set pFrame = New CMiniDockFrameWnd

    'if (pFrame == NULL)
    '    AfxThrowMemoryException();
    'ASSERT_KINDOF(CMiniDockFrameWnd, pFrame);

    Call pFrame.Create(this, dwStyle)
    Set CreateFloatingFrame = pFrame
End Function


Friend Function CalcWindowRect(lpClientRect As RECT, nAdjustType As Long)
    Call Wnd.CalcWindowRect(lpClientRect, nAdjustType)
End Function


Friend Function ActivateTopParent()
    Call Wnd.ActivateTopParent
End Function


Friend Function OnActivate(nState As Long, pWndOther As Long, bMinimized As Boolean)
    ' get top level frame unless this is a child window
    ' determine if window should be active or not

    Dim pTopLevel As CFrame
    If (GetStyle(m_hWnd) And WS_CHILD) <> 0 Then
        Set pTopLevel = this
    Else
        Set pTopLevel = WindowExistsInList(GetTopLevelFrame(this))
    End If

    'ASSERT(pTopLevel != NULL);
    Dim pActive As Long
    If (nState = WA_INACTIVE) Then
        pActive = pWndOther
    Else
        pActive = this.m_hWnd
    End If
    
    Dim bStayActive As Boolean
    If (pTopLevel.m_hWnd = pActive) Then
        bStayActive = True
    Else
        If (pTopLevel.m_hWnd = GetParentFrameMFC(pActive)) Then
            If (pActive = pTopLevel.m_hWnd) Then
                bStayActive = True
            Else
                If CBool(SendMessage(pActive, WM_FLOATSTATUS, FS_SYNCACTIVE, 0&)) Then
                    bStayActive = True
                End If
            End If
        Else
            bStayActive = False
        End If
    End If

    pTopLevel.m_nFlags = pTopLevel.m_nFlags And (Not WF_STAYACTIVE)

    If bStayActive = True Then
        pTopLevel.m_nFlags = pTopLevel.m_nFlags Or WF_STAYACTIVE
    End If

    ' sync floating windows to the new state
    Call NotifyFloatingWindows(IIf(bStayActive, FS_ACTIVATE, FS_DEACTIVATE))

    ' get active view (use active frame if no active view)
    'CView* pActiveView = GetActiveView();
    'if (pActiveView == NULL)
    '    pActiveView = GetActiveFrame()->GetActiveView();

    ' when frame gets activated, re-activate current view
    'If (pActiveView <> 0) Then
    '    If (CBool(nState <> WA_INACTIVE) And (Not bMinimized)) Then
    '        Call OnActivateView(pActiveView, True, pActiveView, pActiveView)
    '    End If
    '
    '    ' always notify the view of frame activations
    '    Call OnActivateFrame(pActiveView, nState, this)
    'End If

End Function

Private Property Let CWnd_nFlags(ByVal RHS As Long)
'
End Property

Private Property Get CWnd_nFlags() As Long
'
End Property


Friend Property Get m_nFlags() As Long
    m_nFlags = Wnd.nFlags
End Property

Friend Property Let m_nFlags(RHS As Long)
    Wnd.nFlags = RHS
End Property




Private Function NotifyFloatingWindows(dwFlags As Long)
    'ASSERT_VALID(this);
    'ASSERT(m_hWnd != NULL);

    ' get top level parent frame window first unless this is a child window
    Dim pParent As Long
    If (GetStyle(m_hWnd) And WS_CHILD) <> 0 Then
        pParent = this.m_hWnd
    Else
        pParent = GetTopLevelFrame(this)
    End If

    'ASSERT(pParent != NULL);
    If (dwFlags And (FS_DEACTIVATE Or FS_ACTIVATE)) <> 0 Then
        ' update parent window activation state
        Dim bActivate As Boolean
        bActivate = Not ((dwFlags And FS_DEACTIVATE) <> 0)
        Dim bEnabled As Boolean

        bEnabled = True                 'bEnabled = pParent.IsWindowEnabled()

        If (bActivate And bEnabled And (pParent <> this.m_hWnd)) Then
            ' Excel will try to Activate itself when it receives a
            ' WM_NCACTIVATE so we need to keep it from doing that here.
            m_nFlags = m_nFlags Or WF_KEEPMINIACTIVE
            Call SendMessage(pParent, WM_NCACTIVATE, True, 0)
            m_nFlags = m_nFlags And (Not WF_KEEPMINIACTIVE)
        Else
            Call SendMessage(pParent, WM_NCACTIVATE, False, 0)
        End If
    End If

    ' then update the state of all floating windows owned by the parent
    Dim hWnd As Long
    hWnd = GetWindow(GetDesktopWindow(), GW_CHILD)
    Do While (hWnd <> 0)
        If (AfxIsDescendant(pParent, hWnd)) Then
            Call SendMessage(hWnd, WM_FLOATSTATUS, dwFlags, 0)
        End If
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop

End Function


Friend Function GetTopLevelFrame(obj As Object) As Long
    GetTopLevelFrame = Wnd.GetTopLevelFrame(obj)
End Function


Friend Function OnNcActivate(bActive As Boolean) As Boolean
    ' stay active if WF_STAYACTIVE bit is on
    If (m_nFlags And WF_STAYACTIVE) <> 0 Then
        bActive = True
    End If

    ' but do not stay active if the window is disabled
    If (Not IsWindowEnabled()) Then
        bActive = False
    End If

    ' do not call the base class because it will call Default()
    ' and we may have changed bActive.
    OnNcActivate = CBool(DefWindowProc(m_hWnd, WM_NCACTIVATE, bActive, 0&))
End Function

Friend Function OnNcCreate(lpcs As CREATESTRUCT) As Boolean '***TEMP NEW***'
    OnNcCreate = True
End Function


Private Function IsWindowEnabled() As Boolean
    IsWindowEnabled = Wnd.IsWindowEnabled()
End Function


Friend Function GetTopLevelParent() As Long '***NEW***'
    GetTopLevelParent = Wnd.GetTopLevelParent()
End Function


Friend Function DelayRecalcLayout(Optional bNotify As Boolean = True)
    m_nIdleFlags = m_nIdleFlags Or (idleLayout Or IIf(bNotify, idleNotify, 0))
    Call OnIdleUpdateCmdUI
End Function


Friend Function OnIdleUpdateCmdUI()
    ' update menu if necessary
    'if (m_nIdleFlags & idleMenu)
    '    OnUpdateFrameMenu(m_hMenuAlt);

    ' update title if necessary
    'if (m_nIdleFlags & idleTitle)
    '    OnUpdateFrameTitle(TRUE);

    ' recalc layout if necessary
    If (m_nIdleFlags And idleLayout) <> 0 Then
        Call RecalcLayout((m_nIdleFlags And idleNotify) <> 0)
        Call UpdateWindow(m_hWnd)
    End If

    ' set the current message string if necessary
    'if (m_nIDTracking != m_nIDLastMessage)
    '{
    '    SetMessageText(m_nIDTracking);
    '    ASSERT(m_nIDTracking == m_nIDLastMessage);
    '}
    m_nIdleFlags = 0

End Function


Friend Function RemoveControlBar(pBar As Object)
    Dim i As Integer
    i = 1
    
    Do While (i <= m_listControlBars.Count)
        If (m_listControlBars(i).m_hWnd = pBar.hWnd) Then
            m_listControlBars.Remove (i)
            Exit Function
        End If
        i = i + 1
    Loop
End Function


Friend Function ReDockControlBar(pBar As CToolbar, pDockBar As CDockBar, ByRef lpRect As RECT)
    'ASSERT(pBar != NULL);
    ' make sure CControlBar::EnableDocking has been called
    'ASSERT(pBar->m_pDockContext != NULL);

    If (pDockBar Is Nothing) Then
        ' Search for the place holder.

        ' In case we don't find a place holder, find a bar with the correct alignment
        ' and keep it in pPossibleBar.
        Dim pPossibleBar As CDockBar
        Set pPossibleBar = Nothing

        Dim i As Integer
        For i = 0 To 3
            Dim pTempBar As CDockBar
            Set pTempBar = GetControlBar(dwDockBarMap(i, 0))

            If Not (pTempBar Is Nothing) Then
                ' Is this the same bar we docked with before?
                If (pTempBar.FindBar(AfxGetDlgCtrlID(pBar.m_hWnd)) > 0) Then
                    Set pDockBar = pTempBar
                    Exit For
                End If
            End If

            If ((dwDockBarMap(i, 1) And CBRS_ALIGN_ANY) = _
                    (pBar.m_dwStyle And CBRS_ALIGN_ANY)) Then
                Set pPossibleBar = GetControlBar(dwDockBarMap(i, 0))

                'ASSERT(pPossibleBar != NULL);
                ' assert fails when initial CBRS_ of bar does not
                ' match available docking sites, as set by EnableDocking()
            End If
        Next i

        ' Did we find the place holder?
        If (pDockBar Is Nothing) Then
            Set pDockBar = pPossibleBar
        End If
    End If

    'ASSERT(pDockBar != NULL);
    'ASSERT(m_listControlBars.Find(pBar) != NULL);
    'ASSERT(pBar->m_pDockSite == this);
    ' if this assertion occurred it is because the parent of pBar was not initially
    ' this CFrameWnd when pBar's OnCreate was called
    ' i.e. this control bar should have been created with a different parent initially

    Call pDockBar.ReDockControlBar(pBar, lpRect)
End Function


'****************(START) DEBUG FUNCTIONS****************
Friend Property Get listControlBars() As Collection
    Set listControlBars = m_listControlBars
End Property
'****************DEBUG FUNCTIONS (END)******************

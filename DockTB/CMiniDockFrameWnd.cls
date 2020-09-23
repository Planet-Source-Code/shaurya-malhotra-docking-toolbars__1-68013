VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMiniDockFrameWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CMiniDockFrameWnd.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CMiniDockFrameWnd class and adapted to
' Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CMiniFrameWnd

Private MiniFrame As New CMiniFrameWnd
Private m_strCaption As String

'****IMPORTANT : SUBCLASSING DATA****
'***********DO NOT MODIFY!!**********
Private mWndProcOrg As Long
Private mHWndSubClassed As Long
'******END OF SUBCLASSING DATA*******

Public m_wndDockBar As New CDockBar

Public Property Get m_bInRecalcLayout() As Boolean
    m_bInRecalcLayout = MiniFrame.bInRecalcLayout
End Property

Public Property Let m_bInRecalcLayout(RHS As Boolean)
    MiniFrame.bInRecalcLayout = RHS
End Property

Public Property Get m_hWndOwner() As Long
    m_hWndOwner = MiniFrame.hWndOwner
End Property

Public Property Let m_hWndOwner(RHS As Long)
'
End Property

Public Function SetWindowPos(hWndInsertAfter As Long, x As Long, y As Long, cx As Long, cy As Long, wFlags As Long) As Long
    SetWindowPos = Globals.SetWindowPos(m_hWnd, hWndInsertAfter, x, y, cx, cy, wFlags)
End Function

Public Function RecalcLayout(bNotify As Boolean)
    If (Not m_bInRecalcLayout) Then
        'CMiniFrameWnd::RecalcLayout(bNotify);
        Call MiniFrame.RecalcLayout(bNotify, this)

        ' syncronize window text of frame window with dockbar itself
        Dim szTitle As String
        szTitle = Space(MAX_PATH_)
        Call GetWindowText(m_wndDockBar.m_hWnd, szTitle, MAX_PATH_)
        Call SetWindowText(m_hWnd, szTitle)
    End If
End Function

Public Function UpdateWindow()
    Call Globals.UpdateWindow(m_hWnd)
End Function

Public Function DestroyWindow()
    Call MiniFrame.DestroyWindow
End Function

Public Function DelayRecalcLayout(Optional bNotify As Boolean = True)
    Call MiniFrame.DelayRecalcLayout(bNotify)
End Function

Public Function GetDlgItem()
'
End Function

Public Function ShowWindow(nCmdShow As Long) As Boolean
    'ASSERT(::IsWindow(m_hWnd));

    'if (m_pCtrlSite == NULL)
    '    return ::ShowWindow(m_hWnd, nCmdShow);
    'else
    '    return m_pCtrlSite->ShowWindow(nCmdShow);

        ShowWindow = Globals.ShowWindow(m_hWnd, nCmdShow)
End Function


Friend Function Create(pParent As CFrame, dwBarStyle As Long) As Boolean
    ' set m_bInRecalcLayout to avoid flashing during creation
    ' RecalcLayout will be called once something is docked
    m_bInRecalcLayout = True

    Dim dwStyle As Long
    dwStyle = WS_POPUP Or WS_CAPTION Or WS_SYSMENU Or MFS_MOVEFRAME _
                Or MFS_4THICKFRAME Or MFS_SYNCACTIVE Or MFS_BLOCKSYSMENU _
                Or FWS_SNAPTOBARS

    If (dwBarStyle And CBRS_SIZE_DYNAMIC) <> 0 Then
        dwStyle = dwStyle And (Not MFS_MOVEFRAME)
    End If

    Dim dwExStyle As Long
    dwExStyle = 0

    Dim rectDefault As RECT
    If (Not MiniFrame.CreateEx(dwExStyle, "", "", dwStyle, rectDefault, pParent.m_hWnd)) Then
        m_bInRecalcLayout = False
        Create = False
        Exit Function
    End If

    If (dwBarStyle And (CBRS_ALIGN_LEFT Or CBRS_ALIGN_RIGHT)) <> 0 Then
        dwStyle = CBRS_ALIGN_LEFT
    Else
        dwStyle = CBRS_ALIGN_TOP
    End If

    dwStyle = dwStyle Or (dwBarStyle And CBRS_FLOAT_MULTI)

    Dim pSysMenu As Long
    pSysMenu = GetSystemMenu(m_hWnd, False)
    Call DeleteMenu(pSysMenu, SC_SIZE, MF_BYCOMMAND)
    Call DeleteMenu(pSysMenu, SC_MINIMIZE, MF_BYCOMMAND)
    Call DeleteMenu(pSysMenu, SC_MAXIMIZE, MF_BYCOMMAND)
    Call DeleteMenu(pSysMenu, SC_RESTORE, MF_BYCOMMAND)

    Dim strHide As String
    strHide = "Hide"
    If CBool(1) Then
        Call DeleteMenu(pSysMenu, SC_CLOSE, MF_BYCOMMAND)
        Call AppendMenu(pSysMenu, MF_STRING Or MF_ENABLED, SC_CLOSE, strHide)
    End If

    ' must initially create with parent frame as parent
    If (Not m_wndDockBar.Create(pParent.m_hWnd, WS_CHILD Or WS_VISIBLE Or dwStyle, _
                            AFX_IDW_DOCKBAR_FLOAT)) Then
        m_bInRecalcLayout = False
        Create = False
        Exit Function
    End If

    ' set parent to CMiniDockFrameWnd
    Call m_wndDockBar.SetParent(this.m_hWnd)
    m_bInRecalcLayout = False

    Create = True

    ' subclass the newly created frame
    If Create = True Then
        Call Subclass
        Dim lpcs As CREATESTRUCT
        Call MiniFrame.OnNcCreate(lpcs)
    End If

End Function


Private Property Get this() As CMiniDockFrameWnd
    Set this = Me
End Property

Public Property Get m_hWnd() As Long
    m_hWnd = MiniFrame.hWnd
End Property

Private Sub Class_Initialize()
    Call AddWindowInList(this)
    Call Initialize
End Sub

Private Sub Class_Terminate()
    Call UnSubClass
    Call RemoveWindowInList(this)
End Sub

Private Function Initialize()
    m_wndDockBar.m_bFloating = True
    m_wndDockBar.m_bAutoDelete = False
End Function


Private Sub Subclass()
    If mWndProcOrg Then Exit Sub
    mWndProcOrg = SetWindowLong(m_hWnd, GWL_WNDPROC, _
                        AddressOf MDF_SubWndProc)
    mHWndSubClassed = m_hWnd
    Call SetWindowLong(m_hWnd, GWL_USERDATA, ObjPtr(Me))
End Sub


Private Sub UnSubClass()
    If mWndProcOrg = 0 Then Exit Sub
    SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg
    mWndProcOrg = 0
End Sub


Friend Function WindowPROC(ByVal hWnd As Long, _
                    ByVal uMsg As Long, ByVal wParam As Long, _
                    lParam As Long) As Long

    Select Case uMsg
        Case WM_NCCREATE
            MsgBox "ABCD"
        Case WM_NCACTIVATE
            WindowPROC = OnNcActivate(CBool(wParam))

        Case WM_ACTIVATE
            Dim tmp As Long
            Call CopyMemory(tmp, VarPtr(lParam), Len(tmp))
            Call OnActivate(GetLoWord(wParam), tmp, CBool(GetHiWord(wParam)))

        Case WM_FLOATSTATUS
            WindowPROC = MiniFrame.OnFloatStatus(wParam, lParam)

        Case WM_NCLBUTTONDOWN
            Dim point As POINTAPI
            Call GetCursorPos(point)

            WindowPROC = OnNcLButtonDown(wParam, point)
            Exit Function

        Case WM_NCLBUTTONDBLCLK
            Dim pt As POINTAPI
            Call GetCursorPos(pt)
            Call ScreenToClient(hWnd, pt)
            Call OnNcLButtonDblClk(wParam, pt)

        Case Else
            WindowPROC = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, lParam)
    End Select

End Function


Friend Function CalcWindowRect(lpClientRect As RECT, Optional nAdjustType As Long = adjustBorder)
    Call MiniFrame.CalcWindowRect(lpClientRect, nAdjustType)
End Function


Private Function OnNcLButtonDown(nHitTest As Long, point As POINTAPI)
    If (nHitTest = HTCAPTION) Then
        ' special activation for floating toolbars
        Call MiniFrame.ActivateTopParent

        ' initiate toolbar drag for non-CBRS_FLOAT_MULTI toolbars
        If ((m_wndDockBar.m_dwStyle And CBRS_FLOAT_MULTI) = 0) Then
            Dim nPos As Long
            nPos = 1
            Dim pBar As CControlBar
            Set pBar = Nothing

            Do While ((pBar Is Nothing) And (nPos < m_wndDockBar.arrBars.GetSize))
                Set pBar = m_wndDockBar.GetDockedControlBar(nPos)
                nPos = nPos + 1
            Loop

            'ASSERT(pBar != NULL);
            'ASSERT_KINDOF(CControlBar, pBar);
            'ASSERT(pBar->m_pDockContext != NULL);
            Call pBar.pDockContext.StartDrag(point)
            Exit Function
        End If
    ElseIf ((nHitTest >= HTSIZEFIRST) And (nHitTest <= HTSIZELAST)) Then
        ' special activation for floating toolbars
        Call MiniFrame.ActivateTopParent

        nPos = 1
        Set pBar = Nothing

        Do While ((pBar Is Nothing) And (nPos < m_wndDockBar.arrBars.GetSize))
            Set pBar = m_wndDockBar.GetDockedControlBar(nPos)
            nPos = nPos + 1
        Loop

        'ASSERT(pBar != NULL);
        'ASSERT_KINDOF(CControlBar, pBar);
        'ASSERT(pBar->m_pDockContext != NULL);

        ' CBRS_SIZE_DYNAMIC toolbars cannot have the CBRS_FLOAT_MULTI style
        'ASSERT((m_wndDockBar.m_dwStyle & CBRS_FLOAT_MULTI) == 0);
        Call pBar.pDockContext.StartResize(nHitTest, point)
        Exit Function
    End If

    Call DefWindowProc(m_hWnd, WM_NCLBUTTONDOWN, nHitTest, 0)
End Function

Private Property Let CMiniFrameWnd_bActive(RHS As Boolean)
    MiniFrame.bActive = RHS
End Property

Private Property Get CMiniFrameWnd_bActive() As Boolean
    CMiniFrameWnd_bActive = MiniFrame.bActive
End Property

Private Property Let CMiniFrameWnd_bInRecalcLayout(RHS As Boolean)
    MiniFrame.bInRecalcLayout = RHS
End Property

Private Property Get CMiniFrameWnd_bInRecalcLayout() As Boolean
    CMiniFrameWnd_bInRecalcLayout = MiniFrame.bInRecalcLayout
End Property

Private Property Let CMiniFrameWnd_bInSys(RHS As Boolean)
    MiniFrame.bInSys = RHS
End Property

Private Property Get CMiniFrameWnd_bInSys() As Boolean
    CMiniFrameWnd_bInSys = MiniFrame.bInSys
End Property

Private Property Let CMiniFrameWnd_bSysTracking(RHS As Boolean)
    MiniFrame.bSysTracking = RHS
End Property

Private Property Get CMiniFrameWnd_bSysTracking() As Boolean
    CMiniFrameWnd_bSysTracking = MiniFrame.bSysTracking
End Property

Public Function IsFrameWnd() As Boolean
    IsFrameWnd = True
End Function

Private Property Get CMiniFrameWnd_hWnd() As Long
    CMiniFrameWnd_hWnd = MiniFrame.hWnd
End Property

Private Property Get CMiniFrameWnd_hWndOwner() As Long
    CMiniFrameWnd_hWndOwner = MiniFrame.hWndOwner
End Property

Private Property Let CMiniFrameWnd_strCaption(RHS As String)
    MiniFrame.strCaption = RHS
End Property

Private Property Get CMiniFrameWnd_strCaption() As String
    CMiniFrameWnd_strCaption = MiniFrame.strCaption
End Property

Friend Function OnActivate(nState As Long, pWndOther As Long, bMinimized As Boolean)
    OnActivate = MiniFrame.OnActivate(nState, pWndOther, bMinimized)
End Function

Friend Function OnNcActivate(bActive As Boolean) As Boolean
    OnNcActivate = MiniFrame.OnNcActivate(bActive)
End Function

Public Property Get hWnd() As Long
    hWnd = MiniFrame.hWnd
End Property


Private Function OnNcLButtonDblClk(nHitTest As Long, point As POINTAPI)
    If (nHitTest = HTCAPTION) Then
        ' special activation for floating toolbars
        Call MiniFrame.ActivateTopParent

        ' initiate toolbar toggle for non-CBRS_FLOAT_MULTI toolbars
        If ((m_wndDockBar.m_dwStyle And CBRS_FLOAT_MULTI) = 0) Then
            Dim nPos As Long
            nPos = 1
            Dim pBar As CControlBar

            Do While (pBar Is Nothing And nPos < m_wndDockBar.arrBars.GetSize())
                Set pBar = m_wndDockBar.GetDockedControlBar(nPos)
                nPos = nPos + 1
            Loop

            'ASSERT(pBar != NULL);
            'ASSERT_KINDOF(CControlBar, pBar);
            'ASSERT(pBar->m_pDockContext != NULL);
            Call pBar.pDockContext.ToggleDocking
            Exit Function
        End If
    End If

    'CMiniFrameWnd::OnNcLButtonDblClk(nHitTest, point);
End Function

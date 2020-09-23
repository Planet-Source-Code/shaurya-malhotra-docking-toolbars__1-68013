VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDockBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CDockBar.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CDockBar class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CControlBar
Private ControlBar As New CControlBar

Private m_rectLayout As RECT

'****IMPORTANT : SUBCLASSING DATA****
'***********DO NOT MODIFY!!**********
Private mWndProcOrg As Long
Private mHWndSubClassed As Long
'******END OF SUBCLASSING DATA*******

Friend Function CalcFixedLayout(bStretch As Boolean, bHorz As Boolean) As Size
    Dim sizeFixed As Size
    sizeFixed = ControlBar.CalcFixedLayout(bStretch, bHorz)
    
    ' get max size
    Dim sizeMax As Size
    
    If (Not (IsRectEmpty(m_rectLayout) <> 0)) Then
        sizeMax = CSize(m_rectLayout)
    Else
        Dim rectFrame As RECT
        Dim pFrame As CFrame
        Set pFrame = GetParentFrame(ControlBar.hWnd)
        Call GetClientRect(pFrame.m_hWnd, rectFrame)
        sizeMax = CSize(rectFrame)
    End If

    ' prepare for layout
    Dim layout As AFX_SIZEPARENTPARAMS
    
    If (m_bLayoutQuery <> 0) Then layout.hDWP = 0 Else _
            layout.hDWP = BeginDeferWindowPos(ControlBar.arrBars.GetSize)

    
    Dim pt As POINTAPI
    pt.x = -cxBorder2
    pt.y = -cyBorder2

    Dim nWidth As Long
    nWidth = 0

    Dim bWrapped As Boolean
    bWrapped = False

    ' layout all the control bars
    Dim nPos As Long
    
    For nPos = 0 To (ControlBar.arrBars.GetSize - 1)
        Dim pBar As CToolbar
        Set pBar = ControlBar.GetDockedControlBar(nPos)

        Dim pVoid As Object
        Set pVoid = ControlBar.arrBars.GetItem(nPos)
        If Not (pBar Is Nothing) Then
            If (pBar.IsVisible) Then
                ' get ideal rect for bar
                Dim dwMode As Long
                dwMode = 0

                If (((pBar.m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0) And _
                    ((pBar.m_dwStyle And CBRS_FLOATING) <> 0)) Then
                    dwMode = dwMode Or (LM_HORZ Or LM_MRUWIDTH)
                ElseIf (pBar.m_dwStyle And CBRS_ORIENT_HORZ) <> 0 Then
                    dwMode = dwMode Or (LM_HORZ Or LM_HORZDOCK)
                Else
                    dwMode = dwMode Or LM_VERTDOCK
                End If

                Dim sizeBar As Size
                sizeBar = pBar.CalcDynamicLayout(-1, dwMode)

                Dim recta As RECT
                recta = CPtRect(pt, sizeBar)

                ' get current rect for bar
                Dim rectBar As RECT
                Call GetWindowRect(pBar.m_hWnd, rectBar)
                Call ScreenToClientRect(m_hWnd, rectBar)

                If (bHorz) Then
                    ' Offset Calculated Rect out to Actual
                    If ((rectBar.Left > recta.Left) And ((Not m_bFloating) <> 0)) <> 0 Then
                        Call OffsetRect(recta, rectBar.Left - recta.Left, 0)
                    End If

                    ' If ControlBar goes off the right, then right justify
                    If ((recta.Right > sizeMax.cx) And ((Not m_bFloating) <> 0)) <> 0 Then
                        Dim x As Long
                        x = GetWidth(recta) - cxBorder2
                        x = GetMax(sizeMax.cx - x, pt.x)
                        Call OffsetRect(recta, x - recta.Left, 0)
                    End If

                    ' If ControlBar has been wrapped, then left justify
                    If (bWrapped) Then
                        bWrapped = False
                        Call OffsetRect(recta, -(recta.Left + cxBorder2), 0)

                    ' If ControlBar is completely invisible, then wrap it
                    ElseIf ((recta.Left >= (sizeMax.cx - cxBorder2)) And _
                        (nPos > 0) And (Not (ControlBar.arrBars.GetItem(nPos - 1) Is Nothing))) <> 0 Then
                        Call ControlBar.arrBars.InsertAt(nPos, Nothing)
                        Set pBar = Nothing
                        Set pVoid = Nothing
                        bWrapped = True
                    End If
                    
                    If (Not bWrapped) Then
                        If (Not IsRectEq(recta, rectBar)) Then
                            If ((Not m_bLayoutQuery) And (Not ((pBar.m_dwStyle And CBRS_FLOATING) <> 0))) Then
                                pBar.m_pDockContext.m_rectMRUDockPos = recta
                            End If

                            Call AfxRepositionWindow(layout, pBar.m_hWnd, recta)
                        End If

                        pt.x = recta.Left + sizeBar.cx - cxBorder2
                        nWidth = GetMax(nWidth, sizeBar.cy)
                    End If
                Else
                    ' Offset Calculated Rect out to Actual
                    If ((rectBar.Top > recta.Top) And (Not m_bFloating)) <> 0 Then
                        Call OffsetRect(recta, 0, rectBar.Top - recta.Top)
                    End If

                    ' If ControlBar goes off the bottom, then bottom justify
                    If ((recta.Bottom > sizeMax.cy) And (Not m_bFloating)) <> 0 Then
                        Dim y As Long
                        y = GetHeight(recta) - cyBorder2
                        y = GetMax(sizeMax.cy - y, pt.y)
                        Call OffsetRect(recta, 0, y - recta.Top)
                    End If

                    ' If ControlBar has been wrapped, then top justify
                    If (bWrapped) Then
                        bWrapped = False
                        Call OffsetRect(recta, 0, -(recta.Top + cyBorder2))
                    ' If ControlBar is completely invisible, then wrap it
                    ElseIf (nPos > 0) Then
                        If ((recta.Top >= (sizeMax.cy - cyBorder2)) And _
                            Not (ControlBar.arrBars.GetItem(nPos - 1) Is Nothing)) <> 0 Then

                            Call ControlBar.arrBars.InsertAt(nPos, Nothing)
                            Set pBar = Nothing
                            Set pVoid = Nothing
                            bWrapped = True
                        End If
                    End If
                    
                    If (Not bWrapped) Then
                        If (Not IsRectEq(recta, rectBar)) Then
                            If ((Not m_bLayoutQuery) And _
                                (Not ((pBar.m_dwStyle And CBRS_FLOATING) <> 0))) Then
                                pBar.m_pDockContext.m_rectMRUDockPos = recta
                            End If
                            Call AfxRepositionWindow(layout, pBar.m_hWnd, recta)
                        End If
                        pt.y = recta.Top + sizeBar.cy - cyBorder2
                        nWidth = GetMax(nWidth, sizeBar.cx)
                    End If
                End If
            End If

            If (Not bWrapped) Then
                ' handle any delay/show hide for the bar
                Call pBar.RecalcDelayShow(layout)
            End If
        End If
        
        If ((pBar Is Nothing) And (pVoid Is Nothing) And (nWidth <> 0)) Then
            ' end of row because pBar == NULL
            If (bHorz) Then
                pt.y = pt.y + (nWidth - cyBorder2)
                sizeFixed.cx = GetMax(sizeFixed.cx, pt.x)
                sizeFixed.cy = GetMax(sizeFixed.cy, pt.y)
                pt.x = -cxBorder2
            Else
                pt.x = pt.x + (nWidth - cxBorder2)
                sizeFixed.cx = GetMax(sizeFixed.cx, pt.x)
                sizeFixed.cy = GetMax(sizeFixed.cy, pt.y)
                pt.y = -cyBorder2
            End If
            nWidth = 0
        End If
    Next nPos

    
    If (Not m_bLayoutQuery) <> 0 Then
        ' move and resize all the windows at once!
        If ((layout.hDWP = 0) Or (Not (EndDeferWindowPos(layout.hDWP) <> 0))) Then
            Debug.Print "Warning: DeferWindowPos failed - low system resources."
        End If
    End If

    ' adjust size for borders on the dock bar itself
    Call SetRectEmpty(recta)
    Call CalcInsideRect(recta, bHorz)
    If (((Not bStretch) Or (Not bHorz)) And (sizeFixed.cx <> 0)) Then
        sizeFixed.cx = sizeFixed.cx + (-recta.Right + recta.Left)
    End If

    If (((Not bStretch) Or bHorz) And (sizeFixed.cy <> 0)) Then
        sizeFixed.cy = sizeFixed.cy + (-recta.Bottom + recta.Top)
    End If

    CalcFixedLayout = sizeFixed
End Function


Private Property Set CControlBar_arrBars(ByVal RHS As CPtrArray)
'
End Property

Private Property Get CControlBar_arrBars() As CPtrArray
'
End Property

Private Property Let CControlBar_bAutoDelete(ByVal RHS As Boolean)
'
End Property

Private Property Get CControlBar_bAutoDelete() As Boolean
'
End Property

Private Property Let CControlBar_bFloating(ByVal RHS As Boolean)
'
End Property

Private Property Get CControlBar_bFloating() As Boolean
'
End Property

Private Property Let CControlBar_bLayoutQuery(ByVal RHS As Boolean)
'
End Property

Private Property Get CControlBar_bLayoutQuery() As Boolean
'
End Property

Private Function CControlBar_CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long) As Boolean
'
End Function

Private Property Let CControlBar_cxDefaultGap(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cxDefaultGap() As Long
'
End Property

Private Property Get CControlBar_cxLeftBorder() As Long
'
End Property

Private Property Let CControlBar_cxLeftBorder(ByVal RHS As Long)
'
End Property

Private Property Let CControlBar_cxRightBorder(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cxRightBorder() As Long
'
End Property

Private Property Let CControlBar_cyBottomBorder(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cyBottomBorder() As Long
'
End Property

Private Property Let CControlBar_cyTopBorder(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cyTopBorder() As Long
'
End Property

Private Property Let CControlBar_dwDockStyle(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_dwDockStyle() As Long
'
End Property

Private Property Let CControlBar_dwStyle(ByVal RHS As Long)
    m_dwStyle = RHS
End Property

Private Property Let m_dwStyle(ByVal RHS As Long)
    ControlBar.dwStyle = RHS
End Property

Public Property Get m_dwStyle() As Long
    m_dwStyle = ControlBar.dwStyle
End Property

Private Property Get CControlBar_dwStyle() As Long
    CControlBar_dwStyle = m_dwStyle()
End Property

Public Property Get m_hWnd() As Long
    m_hWnd = ControlBar.hWnd
End Property

Private Property Get CControlBar_hWnd() As Long
    CControlBar_hWnd = ControlBar.hWnd
End Property

Public Function CControlBar_IsVisible() As Boolean
    CControlBar_IsVisible = ControlBar.IsVisible
End Function

Private Property Let CControlBar_nCount(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_nCount() As Long
'
End Property

Private Property Let CControlBar_nMRUWidth(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_nMRUWidth() As Long
'
End Property

Private Property Let CControlBar_nStateFlags(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_nStateFlags() As Long
'
End Property

Private Property Let CControlBar_pDockContext(ByVal RHS As CDockContext)
'
End Property

Private Property Get CControlBar_pDockContext() As CDockContext
'
End Property


Private Function CalcInsideRect(ByRef recta As RECT, bHorz As Boolean)
    Call ControlBar.CalcInsideRect(recta, bHorz)
End Function


Public Function SetBarStyle(dwStyle_param As Long)
    If (m_dwStyle <> dwStyle_param) Then
        Dim dwOldStyle As Long
        dwOldStyle = m_dwStyle
        m_dwStyle = dwStyle_param
        Call OnBarStyleChange(dwOldStyle, dwStyle_param)
    End If
End Function

Private Function CControlBar_SetBarStyle(dwStyle_param As Long)
    Call SetBarStyle(dwStyle_param)
End Function


Private Function OnBarStyleChange(dwOldStyle As Long, dwNewStyle As Long)
'
End Function


Private Property Let CControlBar_hWndOwner(ByVal RHS As Long)
    ControlBar.hWndOwner = RHS
End Property

Private Property Get CControlBar_hWndOwner() As Long
    CControlBar_hWndOwner = ControlBar.hWndOwner
End Property

Public Property Let m_hWndOwner(ByVal RHS As Long)
    CControlBar_hWndOwner = RHS
End Property

Public Property Get m_hWndOwner() As Long
    CControlBar_hWndOwner = CControlBar_hWndOwner
End Property


Private Property Let CControlBar_pDockSite(ByVal RHS As CFrame)
    ControlBar.pDockSite = RHS
End Property

Private Property Get CControlBar_pDockSite() As CFrame
    Set CControlBar_pDockSite = ControlBar.pDockSite
End Property

Public Property Let m_pDockSite(RHS As CFrame)
    CControlBar_pDockSite = RHS
End Property
Public Property Get m_pDockSite() As CFrame
    Set m_pDockSite = CControlBar_pDockSite
End Property

Private Property Let CControlBar_pDockBar(ByVal RHS As CDockBar)
    ControlBar.pDockBar = RHS
End Property

Private Property Get CControlBar_pDockBar() As CDockBar
    Set CControlBar_pDockBar = ControlBar.pDockBar
End Property

Public Property Let m_pDockBar(RHS As CDockBar)
    CControlBar_pDockBar = RHS
End Property

Public Property Get m_pDockBar() As CDockBar
    Set m_pDockBar = CControlBar_pDockBar
End Property

Public Function Create(pParentWnd As Long, dwStyle As Long, nID As Long) As Boolean
    ' save the style
    m_dwStyle = (dwStyle And CBRS_ALL)
    
    'VERIFY(AfxDeferRegisterClass(AFX_WNDCONTROLBAR_REG));

    ' create the HWND
    Dim recta As RECT
    Call SetRectEmpty(recta)

    ' Note: Parent must resize itself for control bar to be resized
    Create = ControlBar.Create(afxWndControlBar, "", dwStyle, recta, pParentWnd, nID)

    If Create = True Then
        Call Subclass
        Call ControlBar.OnCreate(this)
    End If
End Function



Friend Function WindowPROC(ByVal hWnd As Long, _
         ByVal uMsg As Long, ByVal wParam As Long, _
         lParam As Long) As Long

    Select Case uMsg
        Case WM_DESTROY
            Call ControlBar.OnDestroy
            WindowPROC = 0

        Case WM_NCPAINT
            Call EraseNonClient
            WindowPROC = 0

        Case WM_NCCALCSIZE
            Dim b As Boolean
            If wParam <> 0 Then b = True Else b = False
            Call OnNcCalcSize(b, lParam)

        ' No point putting WM_CREATE here, because this control was
        ' subclassed AFTER it was created
        ' Instead call OnCreate from the function 'Create'
        'Case WM_CREATE
        '    WindowPROC = OnCreate

        Case WM_SIZEPARENT
            Dim tmp As AFX_SIZEPARENTPARAMS
            Call CopyMemory(tmp, lParam, Len(tmp))
            WindowPROC = OnSizeParent(wParam, tmp)
            Call CopyMemory(lParam, tmp, Len(tmp))

        Case Else
            WindowPROC = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End Select

End Function


Private Sub Subclass()
    If mWndProcOrg Then Exit Sub

    mWndProcOrg = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf DB_SubWndProc)
    mHWndSubClassed = m_hWnd

    Call SetWindowLong(m_hWnd, GWL_USERDATA, ObjPtr(Me))
End Sub


Private Sub UnSubClass()
    If mWndProcOrg = 0 Then Exit Sub
    SetWindowLong mHWndSubClassed, GWL_WNDPROC, mWndProcOrg
    mWndProcOrg = 0
End Sub
   

Private Sub Class_Initialize()
    Call Initialize
End Sub

Private Sub Class_Terminate()
    Call UnSubClass
End Sub

Private Function OnSizeParent(wParam As Long, ByRef lParam As AFX_SIZEPARENTPARAMS) As Long
    Dim lpLayout As AFX_SIZEPARENTPARAMS
    lpLayout = lParam

    ' set m_bLayoutQuery to TRUE if lpLayout->hDWP == NULL
    Dim bLayoutQuery As Boolean
    bLayoutQuery = m_bLayoutQuery

    Dim rectLayout As RECT
    rectLayout = m_rectLayout

    m_bLayoutQuery = CBool(lpLayout.hDWP = 0)

    m_rectLayout = lpLayout.recta

    Dim lResult As Long
    lResult = ControlBar.OnSizeParent(wParam, lParam, this)
    ' restore m_bLayoutQuery
    m_bLayoutQuery = bLayoutQuery

    m_rectLayout = rectLayout

    OnSizeParent = lResult
End Function


Private Property Get this() As CDockBar
   Set this = Me
End Property

Public Property Get IsDockBar() As Boolean
    IsDockBar = True
End Property

Public Property Get IsWindowVisible() As Boolean
    IsWindowVisible = CBool(Globals.IsWindowVisible(m_hWnd))
End Property

Public Property Get m_bFloating() As Boolean
    m_bFloating = ControlBar.bFloating
End Property

Public Property Let m_bFloating(RHS As Boolean)
    ControlBar.bFloating = RHS
End Property

Public Property Get m_bLayoutQuery() As Boolean
    m_bLayoutQuery = ControlBar.bLayoutQuery
End Property
Public Property Let m_bLayoutQuery(RHS As Boolean)
    ControlBar.bLayoutQuery = RHS
End Property

Friend Function DockControlBar(pBar As CToolbar, lpRect As RECT)
    'ASSERT_VALID(this);
    'ASSERT_VALID(pBar);
    'ASSERT_KINDOF(CControlBar, pBar);

    Dim rectBar As RECT
    Call GetWindowRect(pBar.m_hWnd, rectBar)
    
    If Not (pBar.m_pDockBar Is Nothing) Then
        If ((pBar.m_pDockBar.m_hWnd = this.m_hWnd) And CBool(IsRectEmpty(lpRect)) Or IsRectEq(rectBar, lpRect)) Then
            ' already docked and no change in position
            Exit Function
        End If
    End If

    ' set CBRS_FLOAT_MULTI style if docking bar has it
    If (m_bFloating And ((pBar.m_dwDockStyle And CBRS_FLOAT_MULTI) <> 0)) Then
        m_dwStyle = m_dwStyle Or CBRS_FLOAT_MULTI
    End If

    m_dwStyle = m_dwStyle And (Not (CBRS_SIZE_FIXED Or CBRS_SIZE_DYNAMIC))
    
    m_dwStyle = m_dwStyle Or pBar.m_dwStyle And (CBRS_SIZE_FIXED Or CBRS_SIZE_DYNAMIC)

    If Not ((m_dwStyle And CBRS_FLOAT_MULTI) <> 0) Then
        'TCHAR szTitle[_MAX_PATH];
        'pBar->GetWindowText(szTitle, _countof(szTitle));
        'AfxSetWindowText(m_hWnd, szTitle);
        Dim szTitle As String
        szTitle = Space(260)
        Call GetWindowText(pBar.m_hWnd, szTitle, MAX_PATH_)
        Call SetWindowText(m_hWnd, szTitle)
    End If

    ' align correctly and turn on all borders
    Dim dwStyle As Long
    dwStyle = pBar.GetBarStyle()
    dwStyle = dwStyle And (Not (CBRS_ALIGN_ANY))
    dwStyle = dwStyle Or (m_dwStyle And CBRS_ALIGN_ANY) Or CBRS_BORDER_ANY

    If (m_bFloating) Then
        dwStyle = dwStyle Or CBRS_FLOATING
    Else
        dwStyle = dwStyle And (Not CBRS_FLOATING)
    End If

    Call pBar.SetBarStyle(dwStyle)

    ' hide first if changing to a new docking site to avoid flashing
    Dim bShow As Boolean
    bShow = False

    If Not (pBar.m_pDockBar Is Nothing) Then
        If ((pBar.m_pDockBar.m_hWnd <> this.m_hWnd) And pBar.IsWindowVisible()) Then
            Call pBar.SetWindowPos(0, 0, 0, 0, 0, _
                SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_HIDEWINDOW)
            bShow = True
        End If
    End If

    Dim nPos As Long
    nPos = -1
    If Not (CBool(IsRectEmpty(lpRect))) Then
        ' insert into appropriate row
        Dim recta As RECT
        Call CopyRect(recta, lpRect)
        Call ScreenToClientRect(Me.m_hWnd, recta)
        Dim ptMid As POINTAPI
        ptMid.x = (recta.Left + (GetWidth(recta) / 2))
        ptMid.y = (recta.Top + (GetHeight(recta) / 2))
        
        nPos = Insert(pBar, recta, ptMid)

        ' position at requested position
        Call pBar.SetWindowPos(0, recta.Left, recta.Top, GetWidth(recta), _
            GetHeight(recta), SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOCOPYBITS)
    Else
        ' always add on current row, then create new one
        ControlBar.arrBars.Add pBar
        ControlBar.arrBars.Add Nothing

        ' align off the edge initially
        Call pBar.SetWindowPos(0, -cxBorder2, -cyBorder2, 0, 0, _
            SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOCOPYBITS)
    End If

    ' attach it to the docking site
    If (pBar.GetParent() <> this.m_hWnd) Then
        Call pBar.SetParent(this.m_hWnd)
    End If
    
    If Not (pBar.m_pDockBar Is Nothing) Then
        If (pBar.m_pDockBar.m_hWnd = this.m_hWnd) Then
            Call pBar.m_pDockBar.RemoveControlBar(pBar, nPos)
        ElseIf Not (pBar.m_pDockBar Is Nothing) Then
            Call pBar.m_pDockBar.RemoveControlBar(pBar, -1, IIf((m_bFloating And (Not (pBar.m_pDockBar.m_bFloating))), 1, 0))
        End If
    End If

    pBar.m_pDockBar = this

    If (bShow) Then
        'ASSERT(!pBar->IsWindowVisible());
        Call pBar.SetWindowPos(0, 0, 0, 0, 0, _
            SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW)
    End If

    ' remove any place holder for pBar in this dockbar
    Call RemovePlaceHolder(pBar)

    ' get parent frame for recalc layout
    Dim pFrameWnd As Object
    Set pFrameWnd = GetDockingFrame()

    If TypeName(pFrameWnd) = "CFrame" Then
        Dim t1 As CFrame
        Set t1 = pFrameWnd
        Call t1.DelayRecalcLayout
    ElseIf TypeName(pFrameWnd) = "CMiniDockFrameWnd" Then
        Dim t2 As CMiniDockFrameWnd
        Set t2 = pFrameWnd
        Call t2.DelayRecalcLayout
    End If
End Function


Friend Function GetDockedCount() As Integer
    Dim nCount As Integer
    nCount = 0
    
    Dim i As Long
    For i = 0 To ControlBar.arrBars.GetSize - 1
        If (Not (ControlBar.GetDockedControlBar(i) Is Nothing)) Then
            nCount = nCount + 1
        End If
    Next i
    
    GetDockedCount = nCount
End Function


Public Property Get GetParent()
'
End Property





Private Function Insert(pBarIns As Object, recta As RECT, ptMid As POINTAPI) As Long
    'ASSERT_VALID(this);
    'ASSERT(pBarIns != NULL);

    Dim nPos As Long
    nPos = 0
    
    Dim nPosInsAfter As Long
    nPosInsAfter = 0
    
    Dim nWidth As Long
    nWidth = 0
    
    Dim nTotalWidth As Long
    nTotalWidth = 0
    
    Dim bHorz As Boolean
    bHorz = CBool((m_dwStyle And CBRS_ORIENT_HORZ) <> 0) 'MAJOR DOCKING PROBLEM SOLVED BY CHANGING THE ABOVE STATEMENT WITH THIS
    
    For nPos = 0 To ControlBar.arrBars.GetSize - 1
        Dim pBar As CToolbar
        Set pBar = ControlBar.GetDockedControlBar(nPos)

        Dim bTmp As Boolean
        If (Not (pBar Is Nothing)) Then
            If (pBar.IsVisible()) Then
                bTmp = True
            Else
                bTmp = False
            End If
        Else
            bTmp = False
        End If

        If bTmp Then
            Dim rectBar As RECT
            Call GetWindowRect(pBar.m_hWnd, rectBar)
            Call ScreenToClientRect(m_hWnd, rectBar)
            
            nWidth = GetMax(nWidth, _
                 IIf(bHorz, GetHeight(rectBar), GetWidth(rectBar) - 1))
            If bHorz Then
                If (recta.Left > rectBar.Left) Then
                    nPosInsAfter = nPos
                End If
            Else
                If (recta.Top > rectBar.Top) Then
                    nPosInsAfter = nPos
                End If
            End If
        Else ' end of row because pBar == NULL
            nTotalWidth = nTotalWidth + (nWidth - cyBorder2)
            
            nWidth = 0
            
            If (IIf(bHorz, ptMid.y, ptMid.x) < nTotalWidth) Then
                If (nPos = 0) Then ' first section
                    Call ControlBar.arrBars.InsertAt(nPosInsAfter + 1, Nothing)
                End If
                Call ControlBar.arrBars.InsertAt(nPosInsAfter + 1, pBarIns)
                
                Insert = nPosInsAfter + 1
                Exit Function
            End If
            nPosInsAfter = nPos
        End If
    Next nPos

    ' create a new row
    Call ControlBar.arrBars.InsertAt(nPosInsAfter + 1, Nothing)
    
    Call ControlBar.arrBars.InsertAt(nPosInsAfter + 1, pBarIns)

    Insert = nPosInsAfter + 1
End Function



Public Function RemoveControlBar(pBar As CToolbar, Optional nPosExclude As Long = -1, Optional nAddPlaceHolder As Long = 0) As Boolean
    'ASSERT(nAddPlaceHolder == 1 || nAddPlaceHolder == 0 || nAddPlaceHolder == -1);
    'ASSERT_VALID(this);
    'ASSERT(pBar != NULL);

    Dim nPos As Long
    nPos = FindBar(pBar, nPosExclude)
    Debug.Assert nPos > 0

    If (nAddPlaceHolder = 1) Then
        Call ControlBar.arrBars.SetAtPos(nPos, AfxGetDlgCtrlID(pBar.m_hWnd))      'Set ControlBar.arrBars(nPos + 1) = pBar

        ' check for already existing place holder
        Dim nPosOld As Integer
        nPosOld = FindBar(ControlBar.arrBars.GetItem(nPos), nPos)

        If (nPosOld > 0) Then
            Call ControlBar.arrBars.RemoveAt(nPos)

            ' remove section indicator (NULL) if nothing else in section
            If ((ControlBar.arrBars.GetItem(nPos - 1) Is Nothing) And (ControlBar.arrBars.GetItem(nPos) Is Nothing)) Then
                Call ControlBar.arrBars.RemoveAt(nPos)
            End If
        End If
    Else
        Call ControlBar.arrBars.RemoveAt(nPos)

        If ((ControlBar.arrBars.GetItem(nPos - 1) Is Nothing) And (ControlBar.arrBars.GetItem(nPos) Is Nothing)) Then
            Call ControlBar.arrBars.RemoveAt(nPos)
        End If

        ' Remove any pre-existing place holders.
        If (nAddPlaceHolder <> -1) Then
            Call RemovePlaceHolder(pBar)
        End If
    End If

    ' don't do anything more in the shutdown case!
    If (pBar.m_pDockContext Is Nothing) Then
        RemoveControlBar = False
        Exit Function
    End If

    ' get parent frame for recalc layout/frame destroy
    Dim pFrameWnd As Object
    Set pFrameWnd = GetDockingFrame()

    If (m_bFloating And (GetDockedVisibleCount() = 0)) Then
        If (GetDockedCount() = 0) Then
            Call pFrameWnd.DestroyWindow
            RemoveControlBar = True     ' Self-Destruct
            Exit Function
        Else
            Call pFrameWnd.ShowWindow(SW_HIDE)
        End If
    Else
        If TypeName(pFrameWnd) = "CFrame" Then
            Dim t1 As CFrame
            Set t1 = pFrameWnd
            Call t1.DelayRecalcLayout
        ElseIf TypeName(pFrameWnd) = "CMiniDockFrameWnd" Then
            Dim t2 As CMiniDockFrameWnd
            Set t2 = pFrameWnd
            Call t2.DelayRecalcLayout
        End If
        'Call pFrameWnd.DelayRecalcLayout
    End If

    RemoveControlBar = False
End Function

Private Function RemovePlaceHolder(pBar As CControlBar)
    ' remove remembered docking position
    Dim nOldPos As Long
    nOldPos = FindBar(AfxGetDlgCtrlID(pBar.hWnd))

    If (nOldPos > 0) Then
        Call ControlBar.arrBars.RemoveAt(nOldPos)

        ' remove section indicator (NULL) if nothing else in section
        If ((ControlBar.arrBars.GetItem(nOldPos - 1) Is Nothing) And (ControlBar.arrBars.GetItem(nOldPos) Is Nothing)) Then
            Call ControlBar.arrBars.RemoveAt(nOldPos)
        End If
    End If
End Function

Private Function GetDockingFrame() As Object
    Dim p_hWnd As Long
    p_hWnd = GetParentFrameMFC(m_hWnd)
    Set GetDockingFrame = WindowExistsInList(p_hWnd)
End Function

Friend Function FindBar(pBar As Variant, Optional nPosExclude As Long = -1) As Long

    Dim comp As Long
    If IsObject(pBar) Then
        If (Not (pBar Is Nothing)) Then
            Dim tmp As CToolbar
            Set tmp = pBar
            comp = tmp.m_hWnd
        Else
            comp = -1
        End If
    Else
        comp = pBar
    End If

    Dim nPos As Long
    For nPos = 0 To ControlBar.arrBars.GetSize - 1
            If ((nPos <> nPosExclude) And (ControlBar.arrBars.GetItemID(nPos) = comp)) Then
                FindBar = nPos
                Exit Function
            End If
    Next nPos
    FindBar = -1
End Function



Public Function GetDockedVisibleCount() As Long
    Dim nCount As Long
    nCount = 0
    
    Dim i As Long
    For i = 0 To ControlBar.arrBars.GetSize - 1
        Dim pBar As CControlBar
        Set pBar = ControlBar.GetDockedControlBar(i)
        If (Not (pBar Is Nothing)) Then
             If (pBar.IsVisible()) Then
                nCount = nCount + 1
            End If
        End If
    Next i
    GetDockedVisibleCount = nCount
End Function


Private Function Initialize(Optional bFloating As Boolean = False)
    m_bFloating = bFloating
    m_bAutoDelete = True

    ControlBar.arrBars.Add Nothing
    
    m_bLayoutQuery = False
    Call SetRectEmpty(m_rectLayout)

    ' assume no margins
    m_cxLeftBorder = 0
    m_cxRightBorder = 0
    m_cyBottomBorder = 0
    m_cyTopBorder = 0
End Function

Public Property Get m_cxLeftBorder() As Long
    m_cxLeftBorder = ControlBar.cxLeftBorder
End Property

Public Property Let m_cxLeftBorder(RHS As Long)
    ControlBar.cxLeftBorder = RHS
End Property

Public Property Get m_cxRightBorder() As Long
    m_cxRightBorder = ControlBar.cxRightBorder
End Property

Public Property Let m_cxRightBorder(RHS As Long)
    ControlBar.cxRightBorder = RHS
End Property

Public Property Get m_cyBottomBorder() As Long
    m_cyBottomBorder = ControlBar.cyBottomBorder
End Property

Public Property Let m_cyBottomBorder(RHS As Long)
    ControlBar.cyBottomBorder = RHS
End Property

Public Property Get m_cyTopBorder() As Long
    m_cyTopBorder = ControlBar.cyTopBorder
End Property

Public Property Let m_cyTopBorder(RHS As Long)
    ControlBar.cyTopBorder = RHS
End Property

Public Property Get m_bAutoDelete() As Boolean
    m_bAutoDelete = ControlBar.bAutoDelete
End Property
Public Property Let m_bAutoDelete(RHS As Boolean)
    ControlBar.bAutoDelete = RHS
End Property


Private Function OnNcCalcSize(bCalcValidRects As Boolean, lpncs As Long)
    Dim lpncsp As NCCALCSIZE_PARAMS
    Call CopyMemory(lpncsp, lpncs, Len(lpncsp))

    ' calculate border space (will add to top/bottom, subtract from right/bottom)
    Dim recta As RECT
    Call SetRectEmpty(recta)        ' not required in VB, but still..

    Call CalcInsideRect(recta, (m_dwStyle And CBRS_ORIENT_HORZ) <> 0)

    ' adjust non-client area for border space
    lpncsp.rgrc(0).Left = lpncsp.rgrc(0).Left + recta.Left
    lpncsp.rgrc(0).Top = lpncsp.rgrc(0).Top + recta.Top
    lpncsp.rgrc(0).Right = lpncsp.rgrc(0).Right + recta.Right
    lpncsp.rgrc(0).Bottom = lpncsp.rgrc(0).Bottom + recta.Bottom

    Call CopyMemory(lpncs, lpncsp, Len(lpncsp))
End Function

Friend Function SetParent(newParent As Long)
    Call Globals.SetParent(m_hWnd, newParent)
End Function

Friend Function GetDockedControlBar(nPos As Long) As Object
    Set GetDockedControlBar = ControlBar.GetDockedControlBar(nPos)
End Function

Public Property Get hWnd() As Long
    hWnd = ControlBar.hWnd
End Property

Friend Function OnMouseActivate(pDesktopWnd As Long, nHitTest As Integer, nMsg As Integer) As Long
    OnMouseActivate = ControlBar.OnMouseActivate(pDesktopWnd, nHitTest, nMsg)
End Function

Friend Function EraseNonClient()
    Call ControlBar.EraseNonClient(this)
End Function

Private Function OnWindowPosChanging(lpWndPos As WINDOWPOS)
    ' not necessary to invalidate the borders
    Dim dwStyle As Long
    dwStyle = m_dwStyle
    
    m_dwStyle = m_dwStyle And (Not (CBRS_BORDER_ANY))
    
    Call ControlBar.OnWindowPosChanging(lpWndPos)
    m_dwStyle = dwStyle

    ' If we can resize while floating
    If (dwStyle And CBRS_SIZE_DYNAMIC) <> 0 Then
        ' And we are resizing
        If (lpWndPos.flags And SWP_NOSIZE) <> 0 Then
            Exit Function
        End If

        ' Then redraw the buttons
        Call ControlBar.Invalidate
    End If
End Function



Friend Property Get arrBars() As CPtrArray
    Set arrBars = ControlBar.arrBars
End Property

Friend Function ReDockControlBar(pBar As CToolbar, ByRef lpRect As RECT)
    'ASSERT_VALID(this);
    'ASSERT_VALID(pBar);

    'ASSERT_KINDOF(CControlBar, pBar);
    'ASSERT(pBar->m_pDockBar != this); // can't redock here if already docked here

    Dim rectBar As RECT
    Call GetWindowRect(pBar.m_hWnd, rectBar)
    
    If ((pBar.m_pDockBar.m_hWnd = this.m_hWnd) And (CBool(IsRectEmpty(lpRect)) Or IsRectEq(rectBar, lpRect))) Then
        ' already docked and no change in position
        Exit Function
    End If

    ' set CBRS_FLOAT_MULTI style if docking bar has it
    If (m_bFloating And (pBar.m_dwDockStyle And CBRS_FLOAT_MULTI) <> 0) Then
        m_dwStyle = m_dwStyle Or CBRS_FLOAT_MULTI
    End If

    m_dwStyle = m_dwStyle And (Not (CBRS_SIZE_FIXED Or CBRS_SIZE_DYNAMIC))
    m_dwStyle = m_dwStyle Or (pBar.m_dwStyle And (CBRS_SIZE_FIXED Or CBRS_SIZE_DYNAMIC))

    If (Not ((m_dwStyle And CBRS_FLOAT_MULTI) <> 0)) Then
        Dim szTitle As String
        szTitle = Space(260)
        Call GetWindowText(pBar.m_hWnd, szTitle, MAX_PATH_)
        Call SetWindowText(m_hWnd, szTitle)
    End If

    ' align correctly and turn on all borders
    Dim dwStyle As Long
    dwStyle = pBar.GetBarStyle()
    
    dwStyle = dwStyle And (Not (CBRS_ALIGN_ANY))
    
    dwStyle = dwStyle Or ((m_dwStyle And CBRS_ALIGN_ANY) Or CBRS_BORDER_ANY)

    If (m_bFloating) Then
        dwStyle = dwStyle Or CBRS_FLOATING
    Else
        dwStyle = dwStyle And (Not CBRS_FLOATING)
    End If

    Call pBar.SetBarStyle(dwStyle)

    Dim nPos As Long
    nPos = FindBar(AfxGetDlgCtrlID(pBar.m_hWnd))

    If (nPos > 0) Then
        Call arrBars.SetAtPos(nPos, pBar)
    End If

    If Not (IsRectEmpty(lpRect)) Then
        Dim recta As RECT
        recta = lpRect
        
        Call ScreenToClientRect(Me.m_hWnd, recta)

        If (nPos < 1) Then
            Dim ptMid As POINTAPI
            ptMid.x = (recta.Left + (GetWidth(recta) / 2))
            ptMid.x = (recta.Top + (GetHeight(recta) / 2))
            nPos = Insert(pBar, recta, ptMid)
        End If

        ' position at requested position
        Call pBar.SetWindowPos(0, recta.Left, recta.Top, GetWidth(recta), _
            GetHeight(recta), SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOCOPYBITS)
    Else
        If (nPos < 1) Then
            ' always add on current row, then create new one
            Call ControlBar.arrBars.Add(pBar)
            Call ControlBar.arrBars.Add(Nothing)
        End If

        ' align off the edge initially
        Call pBar.SetWindowPos(0, -cxBorder2, -cyBorder2, 0, 0, _
            SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOCOPYBITS)
    End If

    ' attach it to the docking site
    If (pBar.GetParent <> Me.hWnd) Then
        Call pBar.SetParent(Me.m_hWnd)
    End If
    
    If Not (pBar.m_pDockBar Is Nothing) Then
        Call pBar.m_pDockBar.RemoveControlBar(pBar)
    End If

    pBar.m_pDockBar = this

    ' get parent frame for recalc layout
    Dim pFrameWnd As CFrame
    Set pFrameWnd = GetDockingFrame()
    Call pFrameWnd.DelayRecalcLayout
End Function


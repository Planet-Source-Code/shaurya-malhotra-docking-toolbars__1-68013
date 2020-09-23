VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDockContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CDockContext.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CDockContext class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Private m_ptLast As POINTAPI            ' last mouse position during drag
Private m_rectLast As RECT
Private m_sizeLast As Size
Private m_bDitherLast As Boolean

' Rectangles used during dragging or resizing
Private m_rectDragHorz As RECT
Private m_rectDragVert As RECT
Private m_rectFrameDragHorz As RECT
Private m_rectFrameDragVert As RECT

Private m_pBar As CToolbar              ' the toolbar that created this context
Private m_pDockSite As CFrame           ' the controlling frame of the CControlBar
Private m_dwDockStyle As Long           ' allowable dock styles for bar
Private m_dwOverDockStyle As Long       ' style of dock that rect is over
Private m_dwStyle As Long               ' style of control bar
Private m_bFlip As Boolean              ' if shift key is down
Private m_bForceFrame As Boolean        ' if ctrl key is down

Private m_pDC As Long                   ' where to draw during drag
Public m_bDragging As Boolean
Private m_nHitTest As Long

Private m_uMRUDockID As Long
Private rectMRUDockPos As RECT

Private m_dwMRUFloatStyle As Long
Private m_ptMRUFloatPos As POINTAPI

Private m_rectRequestedSize     As RECT     ' #define m_rectRequestedSize     m_rectDragHorz
Private m_rectActualSize        As RECT     ' #define m_rectActualSize        m_rectDragVert
Private m_rectActualFrameSize   As RECT     ' #define m_rectActualFrameSize   m_rectFrameDragHorz
Private m_rectFrameBorders      As RECT     ' #define m_rectFrameBorders      m_rectFrameDragVert


Friend Function StartDrag(pt As POINTAPI)
    Dim recta As RECT
    Dim sizeHorz As Size, sizeVert As Size

    m_bDragging = True

    Call InitLoop

    ' GetWindowRect returns screen coordinates (not mirrored),
    ' so if the desktop is mirrored then turn off mirroring
    ' for the desktop dc so that we get correct focus rect drawn.
    ' This layout change should be remembered, just in case ...

    If (m_pBar.m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0 Then
        ' get true bar size (including borders)
        Call GetWindowRect(m_pBar.m_hWnd, recta)
        m_ptLast = pt
        
        Dim sizeFloat As Size
        
        sizeHorz = m_pBar.CalcDynamicLayout(0, LM_HORZ Or LM_HORZDOCK)
        sizeVert = m_pBar.CalcDynamicLayout(0, LM_VERTDOCK)
        sizeFloat = m_pBar.CalcDynamicLayout(0, LM_HORZ Or LM_MRUWIDTH)

        m_rectDragHorz = CRectNEW(recta.Left, recta.Top, sizeHorz)
        m_rectDragVert = CRectNEW(recta.Left, recta.Top, sizeVert)


        ' calculate frame dragging rectangle
        m_rectFrameDragHorz = CRectNEW(recta.Left, recta.Top, sizeFloat)
        m_rectFrameDragVert = CRectNEW(recta.Left, recta.Top, sizeFloat)

        Call CalcBorders(m_rectFrameDragHorz)
        Call CalcBorders(m_rectFrameDragVert)

        Call InflateRect(m_rectFrameDragHorz, -cxBorder2, -cyBorder2)
        Call InflateRect(m_rectFrameDragVert, -cxBorder2, -cyBorder2)
    ElseIf (GetWindowLong(m_pBar.m_hWnd, GWL_STYLE) And CBRS_SIZE_FIXED) <> 0 Then
        ' get true bar size (including borders)
        Call GetWindowRect(m_pBar.m_hWnd, recta)
        m_ptLast = pt

        sizeHorz = m_pBar.CalcDynamicLayout(-1, LM_HORZ Or LM_HORZDOCK)
        sizeVert = m_pBar.CalcDynamicLayout(-1, LM_VERTDOCK)

        ' calculate frame dragging rectangle
        m_rectFrameDragHorz = CRectNEW(recta.Left, recta.Top, sizeHorz)
        m_rectDragHorz = m_rectFrameDragHorz

        m_rectFrameDragVert = CRectNEW(recta.Left, recta.Top, sizeVert)
        m_rectDragVert = m_rectFrameDragVert

        Call CalcBorders(m_rectFrameDragHorz)
        Call CalcBorders(m_rectFrameDragVert)

        Call InflateRect(m_rectFrameDragHorz, -cxBorder2, -cyBorder2)
        Call InflateRect(m_rectFrameDragVert, -cxBorder2, -cyBorder2)
    Else
        ' get true bar size (including borders)

        Call GetWindowRect(m_pBar.m_hWnd, recta)
        m_ptLast = pt
        
        Dim bHorz  As Boolean
        bHorz = HORZF(m_dwStyle)
        
        Dim dwMode As Long
        If (Not bHorz) Then dwMode = (LM_HORZ Or LM_HORZDOCK) Else dwMode = LM_VERTDOCK
        
        Dim sizea As Size
        sizea = m_pBar.CalcDynamicLayout(-1, dwMode)

        ' calculate inverted dragging rect
        If (bHorz) Then
            m_rectDragHorz = recta
            m_rectDragVert = CRectNEW((pt.x - ((recta.Bottom - recta.Top) / 2)), recta.Top, sizea)
        Else ' vertical orientation
            m_rectDragVert = recta
            m_rectDragHorz = CRectNEW(recta.Left, (pt.y - ((recta.Right - recta.Left) / 2)), sizea)
        End If

        ' calculate frame dragging rectangle
        m_rectFrameDragHorz = m_rectDragHorz
        m_rectFrameDragVert = m_rectDragVert

        Call CalcBorders(m_rectFrameDragHorz)
        Call CalcBorders(m_rectFrameDragVert)
        
        Call InflateRect(m_rectFrameDragHorz, -cxBorder2, -cyBorder2)
        Call InflateRect(m_rectFrameDragVert, -cxBorder2, -cyBorder2)
    End If

    ' adjust rectangles so that point is inside
    Call AfxAdjustRectangle(m_rectDragHorz, pt)
    Call AfxAdjustRectangle(m_rectDragVert, pt)
    Call AfxAdjustRectangle(m_rectFrameDragHorz, pt)
    Call AfxAdjustRectangle(m_rectFrameDragVert, pt)

    ' initialize tracking state and enter tracking loop
    m_dwOverDockStyle = CanDock()

    Call Move(pt)                ' call it here to handle special keys
    Call Track
End Function

Private Function HORZF(i As Long) As Long
    If (i And CBRS_ORIENT_HORZ) <> 0 Then HORZF = 1 Else HORZF = 0
End Function

Private Function VERTF(i As Long) As Long
    If (i And CBRS_ORIENT_VERT) <> 0 Then VERTF = 1 Else VERTF = 0
End Function


Private Function CanDock() As Long
    Dim bStyleHorz As Boolean
    Dim dwDock As Long  ' Dock Canidate
    dwDock = 0

    Dim dwCurr As Long  ' Current Orientation
    dwCurr = 0

    ' let's check for something in our current orientation first
    ' then if the shift key is not forcing our orientation then
    ' check for horizontal or vertical orientations as long
    ' as we are close enough

    'ASSERT(m_dwStyle != 0);

    bStyleHorz = HORZF(m_dwStyle)
    If (m_bFlip <> 0) Then bStyleHorz = (Not bStyleHorz) Else bStyleHorz = bStyleHorz

    If (bStyleHorz And HORZF(m_dwDockStyle)) <> 0 Then
        dwDock = m_pDockSite.CanDock(m_rectDragHorz, _
                                      (m_dwDockStyle And (Not CBRS_ORIENT_VERT)))
    ElseIf (VERTF(m_dwDockStyle)) <> 0 Then
        dwDock = m_pDockSite.CanDock(m_rectDragVert, _
                                      (m_dwDockStyle And (Not CBRS_ORIENT_HORZ)))
    End If


    If (Not m_bFlip) <> 0 Then
        If (dwDock = 0 And (HORZF(m_dwDockStyle) <> 0)) Then
            dwCurr = m_pDockSite.CanDock(m_rectDragVert, _
                                          (m_dwDockStyle And (Not CBRS_ORIENT_VERT)))
            dwDock = m_pDockSite.CanDock(m_rectDragHorz, _
                                          (m_dwDockStyle And (Not CBRS_ORIENT_VERT)))
            If (dwDock = dwCurr) Then dwDock = dwDock Else dwDock = 0
        End If
        
        If (dwDock = 0 And (VERTF(m_dwDockStyle) <> 0)) Then
            dwCurr = m_pDockSite.CanDock(m_rectDragHorz, _
                                          (m_dwDockStyle And (Not CBRS_ORIENT_HORZ)))
            dwDock = m_pDockSite.CanDock(m_rectDragVert, _
                                          (m_dwDockStyle And (Not CBRS_ORIENT_HORZ)))
            If (dwDock = dwCurr) Then dwDock = dwDock Else dwDock = 0
        End If
    End If

    CanDock = dwDock
End Function


Private Function Move(pt As POINTAPI)
    Dim ptOffset As POINTAPI
    ptOffset.x = pt.x - m_ptLast.x
    ptOffset.y = pt.y - m_ptLast.y

    ' offset all drag rects to new position
    Call OffsetRect(m_rectDragHorz, ptOffset.x, ptOffset.y)
    Call OffsetRect(m_rectFrameDragHorz, ptOffset.x, ptOffset.y)
    Call OffsetRect(m_rectDragVert, ptOffset.x, ptOffset.y)
    Call OffsetRect(m_rectFrameDragVert, ptOffset.x, ptOffset.y)

    m_ptLast = pt

    ' if control key is down don't dock

    If (m_bForceFrame <> 0) Then
        m_dwOverDockStyle = 0
    Else
        m_dwOverDockStyle = CanDock()
    End If

    ' update feedback
    Call DrawFocusRect
End Function


Public Function DrawFocusRect(Optional bRemoveRect As Boolean = False)
    'ASSERT(m_pDC != NULL);

    ' default to thin frame
    Dim sizea As Size
    sizea.cx = CX_BORDER
    sizea.cy = CY_BORDER

    ' determine new rect and size
    Dim recta As RECT

    Dim pWhiteBrush As Long, pDitherBrush As Long, pBrush As Long
    pWhiteBrush = GetStockObject(WHITE_BRUSH)
    pDitherBrush = GetHalftoneBrush()
    pBrush = pWhiteBrush

    If (HORZF(m_dwOverDockStyle)) Then
        recta = m_rectDragHorz
    ElseIf (VERTF(m_dwOverDockStyle)) Then
        recta = m_rectDragVert
    Else
        ' use thick frame instead
        sizea.cx = GetSystemMetrics(SM_CXFRAME) - CX_BORDER
        sizea.cy = GetSystemMetrics(SM_CYFRAME) - CY_BORDER

        If ((HORZF(m_dwStyle) And (Not m_bFlip)) Or (VERTF(m_dwStyle) And m_bFlip)) Then
            recta = m_rectFrameDragHorz
        Else
            recta = m_rectFrameDragVert
        End If
        pBrush = pDitherBrush
    End If
    
    If (bRemoveRect) Then
        sizea.cx = 0
        sizea.cy = 0
    End If

'    if (afxData.bWin4 &&
'        (HORZF(m_dwOverDockStyle) || VERTF(m_dwOverDockStyle)))
    If ((HORZF(m_dwOverDockStyle) Or VERTF(m_dwOverDockStyle))) Then
        ' looks better one pixel in (makes the bar look pushed down)
        Call InflateRect(recta, -CX_BORDER, -CY_BORDER)
    End If

    ' draw it and remember last size
    If (m_bDitherLast <> 0) Then
        Call DrawDragRect(recta, sizea, m_rectLast, m_sizeLast, _
            pBrush, pDitherBrush, m_pDC)
    Else
        Call DrawDragRect(recta, sizea, m_rectLast, m_sizeLast, _
            pBrush, pWhiteBrush, m_pDC)
    End If

    m_rectLast = recta
    m_sizeLast = sizea

    If (pBrush = pDitherBrush) Then m_bDitherLast = True Else m_bDitherLast = False
End Function


Friend Property Get m_rectMRUDockPos() As RECT
    m_rectMRUDockPos = rectMRUDockPos
End Property

Friend Property Let m_rectMRUDockPos(RHS As RECT)
    rectMRUDockPos = RHS
End Property


Public Function Initialize(pBar As CToolbar)
    'ASSERT(pBar != NULL);
    'ASSERT(pBar->m_pDockSite != NULL);

    Set m_pBar = pBar
    Set m_pDockSite = pBar.m_pDockSite

    m_uMRUDockID = 0
    m_rectMRUDockPos.Left = 0
    m_rectMRUDockPos.Top = 0
    If ((pBar.m_dwStyle And CBRS_SIZE_DYNAMIC) = CBRS_SIZE_DYNAMIC) Then
        m_dwMRUFloatStyle = pBar.m_dwStyle And (CBRS_ALIGN_TOP Or CBRS_SIZE_DYNAMIC)
    ElseIf ((pBar.m_dwStyle And CBRS_ORIENT_HORZ) = CBRS_ORIENT_HORZ) Then
        m_dwMRUFloatStyle = CBRS_ALIGN_TOP Or (pBar.m_dwStyle And CBRS_FLOAT_MULTI)
    Else
        m_dwMRUFloatStyle = CBRS_ALIGN_LEFT Or (pBar.m_dwStyle And CBRS_FLOAT_MULTI)
    End If
    m_ptMRUFloatPos.x = CW_USEDEFAULT

    'ASSERT(m_pDockSite->IsFrameWnd());
    m_pDC = 0
End Function


Private Function Track() As Boolean
    ' don't handle if capture already set
    If (GetCapture() <> 0) Then
        Track = False
        Exit Function
    End If

    ' set capture to the window which received this message
    Call m_pBar.SetCapture
    'ASSERT(m_pBar == CWnd::GetCapture());

    ' get messages until capture lost or cancelled/accepted
    Do While (GetCapture() = m_pBar.m_hWnd)
        Dim msgs As Msg
        If (Not CBool(GetMessage(msgs, 0, 0, 0))) Then
            Call PostQuitMessage(msgs.wParam)
            Exit Do
        End If

        Select Case (msgs.message)
            Case WM_LBUTTONUP
                If (m_bDragging) Then
                    Call EndDrag
                Else
                    Call EndResize
                End If

                Track = True
                Exit Function

            Case WM_MOUSEMOVE
                If (m_bDragging) Then
                    Call Move(msgs.pt)
                Else
                    Call Stretch(msgs.pt)
                End If

            Case WM_KEYUP
                If (m_bDragging) Then
                    Call OnKey(CInt(msgs.wParam), False)
                End If

            Case WM_KEYDOWN
                If (m_bDragging) Then
                    Call OnKey(CInt(msgs.wParam), True)
                End If
                If (msgs.wParam = VK_ESCAPE) Then
                    Call CancelLoop
                    Track = False
                    Exit Function
                End If

            Case WM_RBUTTONDOWN
                Call CancelLoop
                Track = False
                Exit Function
    
            ' just dispatch rest of the messages
            Case Else
                Call DispatchMessage(msgs)

        End Select
    Loop

    Call CancelLoop

    Track = False
End Function


Public Function EndResize()
    Call CancelLoop

    Dim sz As Size
    sz = m_pBar.CalcDynamicLayout(GetWidth(m_rectActualSize), LM_HORZ Or LM_COMMIT)
    
    Call m_pDockSite.FloatControlBar(m_pBar, GetTopLeft(m_rectActualFrameSize), _
        CBRS_ALIGN_TOP Or (m_dwDockStyle And CBRS_FLOAT_MULTI) Or CBRS_SIZE_DYNAMIC)

End Function



Private Function EndDrag()
    Call CancelLoop

    If (m_dwOverDockStyle <> 0) Then
        Dim pDockBar As CDockBar
        Set pDockBar = GetDockBar(m_dwOverDockStyle)

        'Debug.Assert Not (pDockBar Is Nothing)         'ASSERT(pDockBar != NULL);


        Dim recta As RECT
        If (m_dwOverDockStyle And CBRS_ORIENT_VERT) <> 0 Then
            recta = m_rectDragVert
        Else
            recta = m_rectDragHorz
        End If

        Dim uId As Long
        uId = AfxGetDlgCtrlID(pDockBar.m_hWnd)
        
        If (uId >= AFX_IDW_DOCKBAR_TOP And _
                uId <= AFX_IDW_DOCKBAR_BOTTOM) Then
            m_uMRUDockID = uId
            m_rectMRUDockPos = recta

            Call ScreenToClientRect(pDockBar.m_hWnd, rectMRUDockPos)
        End If

        ' dock it at the specified position, RecalcLayout will snap
        Call m_pDockSite.DockControlBar(m_pBar, pDockBar, recta)
        Call m_pDockSite.RecalcLayout

    ElseIf (((m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0) Or (CBool(HORZF(m_dwStyle)) And Not m_bFlip) Or _
            CBool((VERTF(m_dwStyle)) And m_bFlip)) Then
        m_dwMRUFloatStyle = CBRS_ALIGN_TOP Or (m_dwDockStyle And CBRS_FLOAT_MULTI)
        m_ptMRUFloatPos = GetTopLeft(m_rectFrameDragHorz)
        Call m_pDockSite.FloatControlBar(m_pBar, m_ptMRUFloatPos, m_dwMRUFloatStyle)

    Else ' vertical float
        m_dwMRUFloatStyle = CBRS_ALIGN_LEFT Or (m_dwDockStyle And CBRS_FLOAT_MULTI)
        m_ptMRUFloatPos = GetTopLeft(m_rectFrameDragVert)
        Call m_pDockSite.FloatControlBar(m_pBar, m_ptMRUFloatPos, m_dwMRUFloatStyle)
    End If

End Function


Private Function Stretch(pt As POINTAPI) ', m_nHitTest As Long)
    '*******************************************
    m_rectRequestedSize = m_rectDragHorz
    m_rectActualSize = m_rectDragVert
    m_rectActualFrameSize = m_rectFrameDragHorz
    m_rectFrameBorders = m_rectFrameDragVert
    '*******************************************

    Dim ptOffset As POINTAPI
    ptOffset.x = pt.x - m_ptLast.x
    ptOffset.y = pt.y - m_ptLast.y

    ' offset all drag rects to new position
    Dim nLength As Long
    nLength = 0

    Dim dwMode As Long
    dwMode = LM_HORZ

    If ((m_nHitTest = HTLEFT) Or (m_nHitTest = HTRIGHT)) Then
        If (m_nHitTest = HTLEFT) Then
            m_rectRequestedSize.Left = m_rectRequestedSize.Left + ptOffset.x
        Else
            m_rectRequestedSize.Right = m_rectRequestedSize.Right + ptOffset.x
        End If

        nLength = m_rectRequestedSize.Right - m_rectRequestedSize.Left
    Else
        dwMode = dwMode Or LM_LENGTHY

        If (m_nHitTest = HTTOP) Then
            m_rectRequestedSize.Top = m_rectRequestedSize.Top + ptOffset.y
        Else
            m_rectRequestedSize.Bottom = m_rectRequestedSize.Bottom + ptOffset.y
        End If

        nLength = m_rectRequestedSize.Bottom - m_rectRequestedSize.Top
    End If

    If (nLength >= 0) Then nLength = nLength Else nLength = 0

    Dim sizea As Size
    sizea = m_pBar.CalcDynamicLayout(nLength, dwMode)

    Dim rectDesk As RECT
    Dim hWndDesk As Long
    hWndDesk = GetDesktopWindow()
    Call GetWindowRect(hWndDesk, rectDesk)

    Dim rectTemp As RECT
    rectTemp = m_rectActualFrameSize

    Dim recta As RECT
    If ((m_nHitTest = HTLEFT) Or (m_nHitTest = HTTOP)) Then
        rectTemp.Left = rectTemp.Right - _
                (sizea.cx + (m_rectFrameBorders.Right - m_rectFrameBorders.Left))
        rectTemp.Top = rectTemp.Bottom - _
                (sizea.cy + (m_rectFrameBorders.Bottom - m_rectFrameBorders.Top))
        If (IntersectRect(recta, rectDesk, rectTemp)) <> 0 Then
            m_rectActualSize.Left = m_rectActualSize.Right - sizea.cx
            m_rectActualSize.Top = m_rectActualSize.Bottom - sizea.cy
            m_rectActualFrameSize.Left = rectTemp.Left
            m_rectActualFrameSize.Top = rectTemp.Top
        End If
    Else
        rectTemp.Right = rectTemp.Left + _
            (sizea.cx + (m_rectFrameBorders.Right - m_rectFrameBorders.Left))
        rectTemp.Bottom = rectTemp.Top + _
            (sizea.cy + (m_rectFrameBorders.Bottom - m_rectFrameBorders.Top))

        If (IntersectRect(recta, rectDesk, rectTemp)) <> 0 Then
            m_rectActualSize.Right = m_rectActualSize.Left + sizea.cx
            m_rectActualSize.Bottom = m_rectActualSize.Top + sizea.cy
            m_rectActualFrameSize.Right = rectTemp.Right
            m_rectActualFrameSize.Bottom = rectTemp.Bottom
        End If
    End If
    m_ptLast = pt

    '*******************************************
    m_rectDragHorz = m_rectRequestedSize
    m_rectDragVert = m_rectActualSize
    m_rectFrameDragHorz = m_rectActualFrameSize
    m_rectFrameDragVert = m_rectFrameBorders
    '*******************************************

    ' update feedback
    Dim sz As Size
    sz.cx = 3
    sz.cy = 3
    Dim re As RECT
    Call DrawFocusRect

End Function


Private Function CancelLoop()

    Call DrawFocusRect(True)    ' gets rid of focus rect
    Call ReleaseCapture

    Dim pWnd As Long
    pWnd = GetDesktopWindow()
    Call LockWindowUpdate(0)    ' to unlock, pass NULL to LockWindowUpdate
    If (m_pDC <> 0) Then
        Call ReleaseDC(pWnd, m_pDC)
        m_pDC = 0
    End If

End Function


Private Function InitLoop()
    ' handle pending WM_PAINT messages
    Dim msgs As Msg
    Do While (CBool(PeekMessage(msgs, 0, WM_PAINT, WM_PAINT, PM_NOREMOVE)))
        If (Not CBool(GetMessage(msgs, 0, WM_PAINT, WM_PAINT))) Then
            Exit Function
        End If
        Call DispatchMessage(msgs)
    Loop

    ' get styles from bar
    m_dwDockStyle = m_pBar.m_dwDockStyle
    m_dwStyle = m_pBar.m_dwStyle And CBRS_ALIGN_ANY

    'ASSERT(m_dwStyle != 0);

    ' initialize state
    Call SetRectEmpty(m_rectLast)
    m_sizeLast.cx = 0
    m_sizeLast.cy = 0
    
    m_bForceFrame = False
    m_bFlip = False
    m_bDitherLast = False

    ' lock window update while dragging

    'ASSERT(m_pDC == NULL);

    Dim pWnd As Long
    pWnd = GetDesktopWindow()

    If (CBool(LockWindowUpdate(pWnd))) Then
        m_pDC = GetDCEx(pWnd, 0, DCX_WINDOW Or DCX_CACHE Or DCX_LOCKWINDOWUPDATE)
    Else
        m_pDC = GetDCEx(pWnd, 0, DCX_WINDOW Or DCX_CACHE)
    End If

    'ASSERT(m_pDC != NULL);
End Function



Private Function GetDockBar(dwOverDockStyle As Long) As CDockBar
    Dim dw As Long
    dw = 0
    
    Dim pBar As CDockBar
    
    If CBool(HORZF(dwOverDockStyle)) Then
        dw = m_pDockSite.CanDock(m_rectDragHorz, _
            dwOverDockStyle And (Not CBRS_ORIENT_VERT), pBar)
        Set GetDockBar = pBar
        Exit Function
    End If

    If CBool(VERTF(dwOverDockStyle)) Then
        dw = m_pDockSite.CanDock(m_rectDragVert, _
            dwOverDockStyle And (Not CBRS_ORIENT_HORZ), pBar)

        'ASSERT(dw != 0);
        'ASSERT(pBar != NULL);

        Set GetDockBar = pBar
        Exit Function
    End If

    Set GetDockBar = Nothing
End Function



Friend Function StartResize(nHitTest As Long, pt As POINTAPI)
    'ASSERT_VALID(m_pBar);
    'ASSERT(m_pBar->m_dwStyle & CBRS_SIZE_DYNAMIC);

    m_bDragging = False

    Call InitLoop

    ' GetWindowRect returns screen coordinates(not mirrored)
    ' So if the desktop is mirrored then turn off mirroring
    ' for the desktop dc so that we draw correct focus rect

    'if (m_pDC->GetLayout() & LAYOUT_RTL)
    '    m_pDC->SetLayout(LAYOUT_LTR);

    ' get true bar size (including borders)
    Dim recta As RECT
    Call GetWindowRect(m_pBar.m_hWnd, recta)

    m_ptLast = pt
    m_nHitTest = nHitTest

    Dim sizea As Size
    sizea = m_pBar.CalcDynamicLayout(0, LM_HORZ Or LM_MRUWIDTH)
    m_rectRequestedSize = PtSizeToRECT(GetTopLeft(recta), sizea)

    m_rectActualSize = PtSizeToRECT(GetTopLeft(recta), sizea)
    m_rectActualFrameSize = PtSizeToRECT(GetTopLeft(recta), sizea)

    ' calculate frame rectangle
    Call CalcBorders(m_rectActualFrameSize)
    Call InflateRect(m_rectActualFrameSize, -cxBorder2, -cyBorder2)

    '*************************************************
    Dim tmpPt As POINTAPI
    tmpPt.x = 0
    tmpPt.y = 0
    Dim tmpSize As Size
    Dim tmpRECTSize1 As Size, tmpRECTSize2 As Size

    tmpRECTSize1 = CRECTtoSize(m_rectActualFrameSize)
    tmpRECTSize2 = CRECTtoSize(m_rectActualSize)

    tmpSize.cx = tmpRECTSize1.cx - tmpRECTSize2.cx
    tmpSize.cy = tmpRECTSize1.cy - tmpRECTSize2.cy
    m_rectFrameBorders = PtSizeToRECT(tmpPt, tmpSize)
    '*************************************************

    ' initialize tracking state and enter tracking loop
    m_dwOverDockStyle = 0

    '******************************************
    m_rectDragHorz = m_rectRequestedSize
    m_rectDragVert = m_rectActualSize
    m_rectFrameDragHorz = m_rectActualFrameSize
    m_rectFrameDragVert = m_rectFrameBorders
    '******************************************

    Call Stretch(pt)   ' call it here to handle special keys
    Call Track
End Function


' CDockContext Double Click Operations
Friend Function ToggleDocking()

    If (m_pBar.IsFloating()) Then
        ' Dock it only if is allowed to be docked

        If (m_pBar.m_dwDockStyle And CBRS_ALIGN_ANY) <> 0 Then
            'ASSERT((m_uMRUDockID >= AFX_IDW_DOCKBAR_TOP &&
            '    m_uMRUDockID <= AFX_IDW_DOCKBAR_BOTTOM) ||
            '    m_uMRUDockID == 0);

            Dim recta As RECT
            recta = m_rectMRUDockPos

            Dim pDockBar As CDockBar
            Set pDockBar = Nothing

            If (m_uMRUDockID <> 0) Then
                Set pDockBar = m_pDockSite.GetControlBar(m_uMRUDockID)
                Call ClientToScreenRect(pDockBar.m_hWnd, recta)
            End If

            ' dock it at the specified position, RecalcLayout will snap
            Call m_pDockSite.ReDockControlBar(m_pBar, pDockBar, recta)
            Call m_pDockSite.RecalcLayout
        End If
    Else
        Dim ptFloat As POINTAPI
        ptFloat = m_ptMRUFloatPos
        
        If (ptFloat.x < 0 Or ptFloat.y < 0) Then
            ptFloat = GetTopLeft(m_rectMRUDockPos)
            Call ClientToScreen(m_pBar.GetParent, ptFloat)
        End If
        Call m_pDockSite.FloatControlBar(m_pBar, ptFloat, m_dwMRUFloatStyle)
    End If

End Function


Private Function OnKey(nChar As Integer, bDown As Boolean)
    If (nChar = VK_CONTROL) Then _
        Call UpdateState(m_bForceFrame, bDown)
    If (nChar = VK_SHIFT) Then _
        Call UpdateState(m_bFlip, bDown)
End Function


Private Function UpdateState(pFlag As Boolean, bNewValue As Boolean)
    If (pFlag <> bNewValue) Then
        pFlag = bNewValue
        m_bFlip = (HORZF(m_dwDockStyle) And VERTF(m_dwDockStyle) And m_bFlip)  '// shift key
        m_dwOverDockStyle = IIf(m_bForceFrame, 0, CanDock())
        Call DrawFocusRect
    End If
End Function


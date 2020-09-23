VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CToolbar.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CToolbar class and adapted to Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CControlBar

Private Const CX_OVERLAP = 0
Private Const afxDropDownWidth = 13
Private Const TOOLBARCLASSNAME As String = "ToolbarWindow32"

Private Type AFX_COLORMAP
    rgbqFrom As Long
    iSysColorTo As Long
End Type

Private ControlBar As New CControlBar

'****IMPORTANT : SUBCLASSING DATA****
'***********DO NOT MODIFY!!**********
Private mWndProcOrg As Long
Private mHWndSubClassed As Long
'******END OF SUBCLASSING DATA*******

Private afxSysColorMap_(0 To 3) As AFX_COLORMAP
Dim m_sizeButton As Size, m_sizeImage As Size

Private m_hRsrcImageWell As Long            ' handle to loaded resource for image well
Private m_hInstImageWell As Long            ' instance handle to load image well from
Private m_hbmImageWell As Long              ' contains color mapped button images
Private m_bDelayedButtonLayout As Boolean   ' used to manage when button layout should be done




Public Function Create(ByVal pParentWnd As Long, ByVal dwStyle As Long, ByVal nID As Long) As Boolean
    Create = CreateEx(pParentWnd, 0, dwStyle, _
            CRect(m_cxLeftBorder, m_cyTopBorder, m_cxRightBorder, m_cyBottomBorder), nID)

    If Create = True Then
        Call Subclass
        Call ControlBar.OnCreate(this)
    End If
End Function

Private Sub SetBorders(rcBorders As RECT)
    Call ControlBar.SetBorders(rcBorders.Left, rcBorders.Top, rcBorders.Right, rcBorders.Bottom)
End Sub

Private Function CreateEx(ByVal pParentWnd As Long, ByVal dwCtrlStyle As Long, ByVal dwStyle As Long, rcBorders As RECT, ByVal nID As Long) As Boolean
    'ASSERT_VALID(pParentWnd);   // must have a parent
    'ASSERT (!((dwStyle & CBRS_SIZE_FIXED) && (dwStyle & CBRS_SIZE_DYNAMIC)));

    Call SetBorders(rcBorders)

    ' save the style
    m_dwStyle = (dwStyle And CBRS_ALL)
    dwStyle = dwStyle Or (CCS_NOPARENTALIGN Or CCS_NOMOVEY Or CCS_NODIVIDER Or CCS_NORESIZE)
    dwStyle = dwStyle Or dwCtrlStyle

    ' initialize common controls
    'VERIFY(AfxDeferRegisterClass(AFX_WNDCOMMCTL_BAR_REG));
    '_AfxGetComCtlVersion();
    'ASSERT(_afxComCtlVersion != -1);
    '_AfxGetDropDownWidth();
    'ASSERT(_afxDropDownWidth != -1);

    ' create the HWND
    Dim recta As RECT
    Call SetRectEmpty(recta)
    If (Not ControlBar.Create(TOOLBARCLASSNAME, 0, dwStyle, recta, pParentWnd, nID, 0)) Then
        CreateEx = False
        Exit Function
    End If

    ' sync up the sizes
    Call SetSizes(m_sizeButton, m_sizeImage)

    ' Note: Parent must resize itself for control bar to be resized

    CreateEx = True

End Function


Private Function SetSizes(sizeButton As Size, sizeImage As Size)
    'ASSERT_VALID(this);

    ' sizes must be non-zero and positive
    'ASSERT(sizeButton.cx > 0 && sizeButton.cy > 0);
    'ASSERT(sizeImage.cx > 0 && sizeImage.cy > 0);

    ' button must be big enough to hold image
    '    + 7 pixels on x
    '    + 6 pixels on y
    'ASSERT(sizeButton.cx >= sizeImage.cx + 7);
    'ASSERT(sizeButton.cy >= sizeImage.cy + 6);
    If (IsWindow(m_hWnd)) Then
        ' set the sizes via TB_SETBITMAPSIZE and TB_SETBUTTONSIZE
        Call SendMessageLong(m_hWnd, TB_SETBITMAPSIZE, 0, MAKELONG(CInt(sizeImage.cx), CInt(sizeImage.cy)))
        Call SendMessageLong(m_hWnd, TB_SETBUTTONSIZE, 0, MAKELONG(CInt(sizeButton.cx), CInt(sizeButton.cy)))

        Call Invalidate             ' just to be nice if called when toolbar is visible
    Else
        ' just set our internal values for later
        m_sizeButton = sizeButton
        m_sizeImage = sizeImage
    End If
End Function


Friend Function CalcDynamicLayout(nLength As Long, dwMode As Long) As Size
    If ((nLength = -1) And (Not ((dwMode And LM_MRUWIDTH) <> 0)) And _
                        (Not ((dwMode And LM_COMMIT) <> 0)) And _
                        (((dwMode And LM_HORZDOCK) <> 0) Or ((dwMode And LM_VERTDOCK) <> 0))) <> 0 Then
        CalcDynamicLayout = CalcFixedLayout((dwMode And LM_STRETCH) <> 0, (dwMode And LM_HORZDOCK) <> 0)
        Exit Function
    End If

    CalcDynamicLayout = CalcLayout(dwMode, nLength)
End Function


Private Function CalcFixedLayout(bStretch As Boolean, bHorz As Boolean) As Size
    Dim dwMode As Long
    dwMode = IIf(bStretch, LM_STRETCH, 0)

    dwMode = dwMode Or IIf(bHorz, LM_HORZ, 0)

    CalcFixedLayout = CalcLayout(dwMode)
End Function


Friend Function CalcLayout(dwMode As Long, Optional nLength As Long = -1) As Size
    'ASSERT_VALID(this);
    'ASSERT(::IsWindow(m_hWnd));
    
    'If (dwMode And LM_HORZDOCK) <> 0 Then
        'ASSERT(dwMode & LM_HORZ);
    'End If

    Dim nCount As Long
    Dim pData() As TBBUTTON

    Dim sizeResult As Size
    sizeResult.cx = 0
    sizeResult.cy = 0


    ' BLOCK: Load Buttons
    nCount = SendMessage(m_hWnd, TB_BUTTONCOUNT, 0, 0)
    If (nCount <> 0) Then
        Dim i As Long
        ReDim pData(nCount) As TBBUTTON

        For i = 0 To nCount
            Call GetButton(i, pData(i))
        Next i
    End If

    If (nCount > 0) Then
        If (Not ((m_dwStyle And CBRS_SIZE_FIXED) <> 0)) Then
            Dim bDynamic As Boolean
            bDynamic = (m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0

            If (bDynamic And ((dwMode And LM_MRUWIDTH) <> 0)) <> 0 Then
                Call SizeToolBar(pData, nCount, m_nMRUWidth)

            ElseIf (bDynamic And ((dwMode And LM_HORZDOCK) <> 0)) <> 0 Then
                Call SizeToolBar(pData, nCount, 32767)

            ElseIf (bDynamic And ((dwMode And LM_VERTDOCK) <> 0)) <> 0 Then
                Call SizeToolBar(pData, nCount, 0)

            ElseIf (bDynamic And (nLength <> -1)) <> 0 Then
                Dim recta As RECT
                Call SetRectEmpty(recta)

                Call ControlBar.CalcInsideRect(recta, (dwMode And LM_HORZ))  '**CHECK AGAIN**'

                Dim bVert As Boolean
                bVert = ((dwMode And LM_LENGTHY) <> 0)

                Dim nLen As Long
                If bVert = True Then nLen = nLength + (recta.Bottom - recta.Top) Else nLen = nLength + (recta.Right - recta.Left)

                Call SizeToolBar(pData, nCount, nLen, bVert)

            ElseIf (bDynamic And ((m_dwStyle And CBRS_FLOATING) <> 0)) <> 0 Then
                Call SizeToolBar(pData, nCount, m_nMRUWidth)

            Else
                If (dwMode And LM_HORZ) = LM_HORZ Then
                    Call SizeToolBar(pData, nCount, 32767)
                Else
                    Call SizeToolBar(pData, nCount, 0)
                End If
            End If
        End If

        sizeResult = CalcSize(pData, nCount)

        If (dwMode And LM_COMMIT) <> 0 Then
            Dim pControl() As AFX_CONTROLPOS

            Dim nControlCount As Integer
            nControlCount = 0

            Dim bIsDelayed As Boolean
            bIsDelayed = m_bDelayedButtonLayout
            m_bDelayedButtonLayout = False

            For i = 0 To (nCount - 1)
                If (((pData(i).fsStyle And TBSTYLE_SEP) <> 0) And (pData(i).idCommand <> 0)) <> 0 Then
                    nControlCount = nControlCount + 1
                End If
            Next i

            If (nControlCount > 0) Then
                ReDim pControl(nControlCount) As AFX_CONTROLPOS
                nControlCount = 0

                For i = 0 To (nCount - 1)
                    If (((pData(i).fsStyle And TBSTYLE_SEP) <> 0) And (pData(i).idCommand <> 0)) <> 0 Then
                        pControl(nControlCount).nIndex = i
                        pControl(nControlCount).nID = pData(i).idCommand

                        Dim recta1 As RECT
                        Call GetItemRect(i, recta1)
                        Call ClientToScreenRect(m_hWnd, recta1)
                        pControl(nControlCount).rectOldPos = recta1

                        nControlCount = nControlCount + 1
                    End If
                Next i
            End If

            If (((m_dwStyle And CBRS_FLOATING) = CBRS_FLOATING) And ((m_dwStyle And CBRS_SIZE_DYNAMIC) = CBRS_SIZE_DYNAMIC)) <> 0 Then
                m_nMRUWidth = sizeResult.cx
            End If

            For i = 0 To (nCount - 1)
                Call SetButton(i, pData(i))
            Next i

            If (nControlCount > 0) Then
                For i = 0 To (nControlCount - 1)
                    Dim pWnd As Long
                    pWnd = GetDlgItemTB(pControl(i).nID)

                    If (pWnd <> 0) Then
                        Dim recta2 As RECT
                        Call GetWindowRect(pWnd, recta2)

                        Dim pt As POINTAPI

                        pt.x = recta2.Left - pControl(i).rectOldPos.Left
                        pt.y = recta2.Top - pControl(i).rectOldPos.Top

                        Call GetItemRect(pControl(i).nIndex, recta2)

                        pt.x = recta2.Left + pt.x
                        pt.y = recta2.Top + pt.y

                        Call Globals.SetWindowPos(pWnd, 0, pt.x, pt.y, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER)
                    End If

                Next i

            End If

            m_bDelayedButtonLayout = bIsDelayed
        End If
    End If


    ' BLOCK: Adjust Margins
    Dim recta3 As RECT
    Call SetRectEmpty(recta3)

    Call ControlBar.CalcInsideRect(recta3, (dwMode And LM_HORZ))  '**CHECK AGAIN**'

    sizeResult.cy = sizeResult.cy - (recta3.Bottom - recta3.Top)
    sizeResult.cx = sizeResult.cx - (recta3.Right - recta3.Left)

    Dim sizea As Size
    sizea = ControlBar.CalcFixedLayout((dwMode And LM_STRETCH), (dwMode And LM_HORZ)) '**CHECK AGAIN**'

    sizeResult.cx = GetMax(sizeResult.cx, sizea.cx)
    sizeResult.cy = GetMax(sizeResult.cy, sizea.cy)

    CalcLayout = sizeResult

End Function


Private Function CalcSize(ByRef pData() As TBBUTTON, nCount As Long) As Size
        'ASSERT(pData != NULL && nCount > 0);

        Dim cur As POINTAPI
        cur.x = 0
        cur.y = 0
        
        Dim sizeResult As Size
        sizeResult.cx = 0
        sizeResult.cy = 0

        Dim dwExtendedStyle As Long
        dwExtendedStyle = SendMessage(m_hWnd, TB_GETEXTENDEDSTYLE, 0, 0)

        Dim i As Long
        For i = 0 To (nCount - 1)
                ' WINBUG: The IE4 version of COMCTL32.DLL calculates the separation
                '   on a TBSTYLE_WRAP button as 100% of the value in iBitmap compared
                '   to the other versions which calculate it at 2/3 of that value.
                '   This is actually a bug which should be fixed in IE 4.01, so we
                '   only do the 100% calculation specifically for IE4.

                Dim cySep As Long
                cySep = pData(i).iBitmap

                'ASSERT(_afxComCtlVersion != -1);

                If (Not ((GetStyle(m_hWnd) And TBSTYLE_FLAT) <> 0)) <> 0 Then  '&& _afxComCtlVersion != VERSION_IE4)
                        cySep = cySep * 2 / 3
                End If

                If (pData(i).fsState And TBSTATE_HIDDEN) <> 0 Then
                        GoTo ContinueWithFor
                End If

                Dim cx As Long
                cx = m_sizeButton.cx

                If (pData(i).fsStyle And TBSTYLE_SEP) <> 0 Then
                        ' a separator represents either a height or width
                        If (pData(i).fsState And TBSTATE_WRAP) <> 0 Then
                                sizeResult.cy = GetMax(cur.y + m_sizeButton.cy + cySep, sizeResult.cy)
                        Else
                                sizeResult.cx = GetMax(cur.x + pData(i).iBitmap, sizeResult.cx)
                        End If
                Else
                        ' check for dropdown style, but only if the buttons are being drawn
                        If (((pData(i).fsStyle And TBSTYLE_DROPDOWN) <> 0) And _
                                ((dwExtendedStyle And TBSTYLE_EX_DRAWDDARROWS) <> 0)) <> 0 Then
                            ' add size of drop down

                            'ASSERT(_afxDropDownWidth != -1);
                            cx = cx + afxDropDownWidth
                        End If
                        sizeResult.cx = GetMax(cur.x + cx, sizeResult.cx)
                        sizeResult.cy = GetMax(cur.y + m_sizeButton.cy, sizeResult.cy)

                End If

                If (pData(i).fsStyle And TBSTYLE_SEP) <> 0 Then
                        cur.x = cur.x + pData(i).iBitmap
                Else
                        cur.x = cur.x + (cx - CX_OVERLAP)
                End If

                If (pData(i).fsState And TBSTATE_WRAP) <> 0 Then
                        cur.x = 0
                        cur.y = cur.y + m_sizeButton.cy
                        If (pData(i).fsStyle And TBSTYLE_SEP) <> 0 Then
                                cur.y = cur.y + cySep
                        End If
                End If

ContinueWithFor:
        Next i

        CalcSize = sizeResult
End Function



Private Function GetButton(nIndex As Long, ByRef pButton As TBBUTTON)
    Dim pBar As Long
    pBar = Me.m_hWnd

    Call SendMessage(pBar, TB_GETBUTTON, nIndex, pButton)

    ' TBSTATE_ENABLED == TBBS_DISABLED so invert it
    pButton.fsState = pButton.fsState Xor TBSTATE_ENABLED
End Function


Private Function SizeToolBar(pData() As TBBUTTON, nCount As Long, nLength As Long, Optional bVert As Boolean = False)
    'ASSERT(pData != NULL && nCount > 0);

    If (bVert = False) Then
        Dim nMin As Long, nMax As Long, nTarget As Long, nCurrent As Long, nMid As Long

        ' Wrap ToolBar as specified
        nMax = nLength

        nTarget = WrapToolBar(pData, nCount, nMax)

        ' Wrap ToolBar vertically
        nMin = 0

        nCurrent = WrapToolBar(pData, nCount, nMin)

        If (nCurrent <> nTarget) Then
            Do While (nMin < nMax)
                nMid = Fix((nMin + nMax) / 2)

                nCurrent = WrapToolBar(pData, nCount, nMid)

                If (nCurrent = nTarget) Then
                    nMax = nMid
                Else
                    If (nMin = nMid) Then
                        Call WrapToolBar(pData, nCount, nMax)
                        Exit Do
                    End If

                    nMin = nMid
                End If
            Loop
        End If

        Dim sizea As Size
        sizea = CalcSize(pData, nCount)

        Call WrapToolBar(pData, nCount, sizea.cx)

    Else
        Dim sizeMax As Size, sizeMin As Size, sizeMid As Size

        ' Wrap ToolBar vertically
        Call WrapToolBar(pData, nCount, 0)

        sizeMin = CalcSize(pData, nCount)

        ' Wrap ToolBar horizontally
        Call WrapToolBar(pData, nCount, 32767)

        sizeMax = CalcSize(pData, nCount)

        Do While (sizeMin.cx < sizeMax.cx)
            sizeMid.cx = Fix((sizeMin.cx + sizeMax.cx) / 2)

            Call WrapToolBar(pData, nCount, sizeMid.cx)

            sizeMid = CalcSize(pData, nCount)

            If (nLength < sizeMid.cy) Then
                If ((sizeMin.cx = sizeMid.cx) And (sizeMin.cy = sizeMid.cy)) Then '**CHECK AGAIN**'
                    Call WrapToolBar(pData, nCount, sizeMax.cx)
                    Exit Function
                End If

                sizeMin = sizeMid

            ElseIf (nLength > sizeMid.cy) Then
                sizeMax = sizeMid
            Else
                Exit Function
            End If

        Loop

    End If

End Function


Private Function GetItemRect(nIndex As Long, ByRef lpRect As RECT)
    'ASSERT_VALID(this);
    'ASSERT(::IsWindow(m_hWnd));

    ' handle any delayed layout
    If (m_bDelayedButtonLayout) Then
        Call layout
    End If

    ' now it is safe to get the item rectangle
    If (Not SendMessage(m_hWnd, TB_GETITEMRECT, nIndex, lpRect)) <> 0 Then
        Call SetRectEmpty(lpRect)
    End If

End Function


Private Function WrapToolBar(pData() As TBBUTTON, nCount As Long, nWidth As Long) As Integer
    'ASSERT(pData != NULL && nCount > 0);

    Dim nResult As Integer
    nResult = 0

    Dim x As Integer
    x = 0

    Dim i As Integer
    For i = 0 To (nCount - 1)
        pData(i).fsState = pData(i).fsState And (Not TBSTATE_WRAP)

        If (pData(i).fsState And TBSTATE_HIDDEN) <> 0 Then GoTo Continue

        Dim dx As Integer, dxNext As Integer

        If (pData(i).fsStyle And TBSTYLE_SEP) <> 0 Then
            dx = pData(i).iBitmap
            dxNext = dx
        Else
            dx = m_sizeButton.cx
            dxNext = dx - CX_OVERLAP
        End If

        If ((x + dx) > nWidth) Then
            Dim bFound As Boolean
            bFound = False

            Dim j As Integer
            j = i

            Do While (j >= 0)
            If Not ((j >= 0) And (Not ((pData(j).fsState And TBSTATE_WRAP) <> 0))) Then Exit Do
                ' Find last separator that isn't hidden
                ' a separator that has a command ID is not
                ' a separator, but a custom control.
                If (((pData(j).fsStyle And TBSTYLE_SEP) <> 0) And _
                    (pData(j).idCommand = 0) And _
                    Not ((pData(j).fsState And TBSTATE_HIDDEN) <> 0)) Then
                    bFound = True
                    i = j
                    x = 0

                    pData(j).fsState = pData(j).fsState Or TBSTATE_WRAP

                    nResult = nResult + 1

                    Exit Do
                End If

            j = j - 1
            Loop

            If Not (bFound) Then
                For j = (i - 1) To 0 Step -1
                    If Not ((j >= 0) And (Not ((pData(j).fsState And TBSTATE_WRAP) <> 0))) Then Exit For

                    ' Never wrap anything that is hidden,
                    ' or any custom controls
                    If (((pData(j).fsState And TBSTATE_HIDDEN) <> 0) Or _
                        (((pData(j).fsStyle And TBSTYLE_SEP) <> 0) And _
                        (pData(j).idCommand <> 0))) Then
                        GoTo ContinueWithFor
                    End If

                    bFound = True
                    i = j
                    x = 0

                    pData(j).fsState = pData(j).fsState Or TBSTATE_WRAP

                    nResult = nResult + 1

                    Exit For
ContinueWithFor:
                Next j

                If Not (bFound) Then
                    x = x + dxNext
                End If

            End If

        Else
            x = x + dxNext
        End If

Continue:
    Next i

    WrapToolBar = nResult + 1

End Function


Private Function SetButton(nIndex As Long, ByRef pButton As TBBUTTON)
    ' get original button state
    Dim button As TBBUTTON

    Call SendMessage(m_hWnd, TB_GETBUTTON, nIndex, button)   '**CHECK AGAIN**'

    ' prepare for old/new button comparsion
    button.bReserved(1) = 0
    button.bReserved(2) = 0

    ' TBSTATE_ENABLED == TBBS_DISABLED so invert it
    pButton.fsState = pButton.fsState Xor TBSTATE_ENABLED

    pButton.bReserved(1) = 0
    pButton.bReserved(2) = 0

    ' nothing to do if they are the same
    If (CmpButtons(pButton, button) = False) Then
        ' don't redraw everything while setting the button
        Dim dwStyle As Long
        dwStyle = GetStyle(m_hWnd)

        Call ModifyStyle(m_hWnd, WS_VISIBLE, 0)

        Call SendMessage(m_hWnd, TB_DELETEBUTTON, nIndex, 0)
        Call SendMessage(m_hWnd, TB_INSERTBUTTON, nIndex, pButton) '**CHECK AGAIN**'

        Call ModifyStyle(m_hWnd, 0, dwStyle And WS_VISIBLE)

        ' invalidate appropriate parts
        If ((((pButton.fsStyle Xor button.fsStyle) And TBSTYLE_SEP) <> 0) Or _
                (((pButton.fsStyle And TBSTYLE_SEP) <> 0) And (pButton.iBitmap <> button.iBitmap))) Then
            ' changing a separator
            MsgBox "Changing a separator..."
            'Invalidate                     'Invalidate();
        Else
            ' invalidate just the button
            Dim recta As RECT
            If (SendMessage(m_hWnd, TB_GETITEMRECT, nIndex, recta)) Then
                Call InvalidateRect(m_hWnd, recta, True)
            End If

        End If
    End If

End Function


Private Function CmpButtons(b1 As TBBUTTON, b2 As TBBUTTON) As Boolean
    If (b1.bReserved(1) = b2.bReserved(1)) _
        And (b1.bReserved(2) = b2.bReserved(2)) _
        And (b1.dwData = b2.dwData) _
        And (b1.fsState = b2.fsState) _
        And (b1.fsStyle = b2.fsStyle) _
        And (b1.iBitmap = b2.iBitmap) _
        And (b1.idCommand = b2.idCommand) _
        And (b1.iString = b2.iString) _
    Then _
        CmpButtons = True _
    Else _
        CmpButtons = False
End Function


Public Function GetDlgItemTB(nID As Long) As Long
    Call GetDlgItem(m_hWnd, nID)
End Function

Public Property Get m_hWnd() As Long
   m_hWnd = ControlBar.hWnd
End Property

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
    'CWnd_CreateEx = CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam)
End Function

Friend Function Invalidate(Optional bErase As Boolean = True) As Variant
    ControlBar.Invalidate (bErase)
End Function

Private Property Let CControlBar_cxDefaultGap(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cxDefaultGap() As Long
'
End Property

Private Property Let CControlBar_cxLeftBorder(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_cxLeftBorder() As Long
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

Private Property Get CControlBar_dwDockStyle() As Long
    CControlBar_dwDockStyle = ControlBar.dwDockStyle
End Property

Private Property Let CControlBar_dwDockStyle(ByVal RHS As Long)
    ControlBar.dwDockStyle = RHS
End Property

Private Property Let CControlBar_dwStyle(ByVal RHS As Long)
    ControlBar.dwStyle = RHS
End Property

Private Property Get CControlBar_dwStyle() As Long
    CControlBar_dwStyle = ControlBar.dwStyle
End Property

Private Property Get CControlBar_hWnd() As Long
    CControlBar_hWnd = ControlBar.hWnd
End Property

Private Function CControlBar_IsVisible() As Boolean
    CControlBar_IsVisible = ControlBar.IsVisible
End Function

Public Function IsVisible() As Boolean
    IsVisible = CControlBar_IsVisible
End Function

Private Property Get CControlBar_nCount() As Long
'
End Property

Private Property Let CControlBar_nCount(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_nMRUWidth() As Long
'
End Property

Private Property Let CControlBar_nMRUWidth(ByVal RHS As Long)
'
End Property

Private Property Get CControlBar_nStateFlags() As Long
'
End Property

Private Property Let CControlBar_nStateFlags(ByVal RHS As Long)
'
End Property

Public Property Let m_cxDefaultGap(ByVal RHS As Long)
    ControlBar.cxDefaultGap = RHS
End Property

Public Property Get m_cxDefaultGap() As Long
    m_cxDefaultGap = ControlBar.cxDefaultGap
End Property

Public Property Let m_cxLeftBorder(ByVal RHS As Long)
    ControlBar.cxLeftBorder = RHS
End Property

Public Property Get m_cxLeftBorder() As Long
    m_cxLeftBorder = ControlBar.cxLeftBorder
End Property

Public Property Let m_cxRightBorder(ByVal RHS As Long)
    ControlBar.cxRightBorder = RHS
End Property

Public Property Get m_cxRightBorder() As Long
    m_cxRightBorder = ControlBar.cxRightBorder
End Property

Public Property Let m_cyBottomBorder(ByVal RHS As Long)
    ControlBar.cyBottomBorder = RHS
End Property

Public Property Get m_cyBottomBorder() As Long
    m_cyBottomBorder = ControlBar.cyBottomBorder
End Property

Public Property Let m_cyTopBorder(ByVal RHS As Long)
    ControlBar.cyTopBorder = RHS
End Property

Public Property Get m_cyTopBorder() As Long
    m_cyTopBorder = ControlBar.cyTopBorder
End Property

Public Property Let m_dwDockStyle(ByVal RHS As Long)
    ControlBar.dwDockStyle = RHS
End Property

Public Property Get m_dwDockStyle() As Long
    m_dwDockStyle = ControlBar.dwDockStyle
End Property

Public Property Let m_dwStyle(ByVal RHS As Long)
    ControlBar.dwStyle = RHS
End Property

Public Property Get m_dwStyle() As Long
    m_dwStyle = ControlBar.dwStyle
End Property

Public Property Let m_nCount(ByVal RHS As Long)
    ControlBar.nCount = RHS
End Property

Public Property Get m_nCount() As Long
    m_nCount = ControlBar.nCount
End Property

Public Property Let m_nMRUWidth(ByVal RHS As Long)
    ControlBar.nMRUWidth = RHS
End Property

Public Property Get m_nMRUWidth() As Long
    m_nMRUWidth = ControlBar.nMRUWidth
End Property

Public Property Let m_nStateFlags(ByVal RHS As Long)
    ControlBar.nStateFlags = RHS
End Property

Public Property Get m_nStateFlags() As Long
    m_nStateFlags = ControlBar.nStateFlags
End Property


Public Property Let m_bFloating(ByVal RHS As Boolean)
    ControlBar.bFloating = RHS
End Property

Public Property Get m_bFloating() As Boolean
    m_bFloating = ControlBar.bFloating
End Property

Public Property Let m_bLayoutQuery(ByVal RHS As Boolean)
    ControlBar.bLayoutQuery = RHS
End Property

Public Property Get m_bLayoutQuery() As Boolean
    m_bLayoutQuery = ControlBar.bLayoutQuery
End Property

Private Property Let m_pDockContext(RHS As CDockContext)
    ControlBar.pDockContext = RHS
End Property
Friend Property Get m_pDockContext() As CDockContext
    Set m_pDockContext = ControlBar.pDockContext
End Property

Private Property Let CControlBar_pDockContext(ByVal RHS As CDockContext)
    ControlBar.pDockContext = RHS
End Property

Private Property Get CControlBar_pDockContext() As CDockContext
    Set CControlBar_pDockContext = ControlBar.pDockContext
End Property

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






Friend Function RecalcDelayShow(lpLayout As AFX_SIZEPARENTPARAMS) As Long
    'ASSERT(lpLayout != NULL);

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
                m_nStateFlags = m_nStateFlags And (Not (delayShow Or delayHide)) '**CHECK AGAIN**'
                ' hide/show the window if actually doing layout
                lpLayout.hDWP = DeferWindowPos(lpLayout.hDWP, m_hWnd, 0, _
                        0, 0, 0, 0, swpFlags Or _
                        SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
            End If
        Else
            ' clear delay flags -- window is already in correct state
            m_nStateFlags = m_nStateFlags And (Not (delayShow Or delayHide))
        End If
    End If

    RecalcDelayShow = dwStyle
End Function


Private Function Initialize()
    ' initialize state

    'm_pStringMap = NULL;
    'm_hRsrcImageWell = NULL;
    'm_hInstImageWell = NULL;
    m_hbmImageWell = 0
    m_bDelayedButtonLayout = True

    ' default image sizes
    m_sizeImage.cx = 16
    m_sizeImage.cy = 15

    ' default button sizes
    m_sizeButton.cx = 23
    m_sizeButton.cy = 22

    ' top and bottom borders are 1 larger than default for ease of grabbing
    m_cyTopBorder = 3
    m_cyBottomBorder = 3
End Function

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

Private Sub Class_Initialize()
    Call Initialize
End Sub


Public Function SetButtons(lpIDArray() As Long, nIDCount As Integer) As Boolean
    Dim i As Integer        ' used as counter in for loops

    'ASSERT_VALID(this);
    'ASSERT(nIDCount >= 1);  // must be at least one of them
    'ASSERT(lpIDArray == NULL ||
    '    AfxIsValidAddress(lpIDArray, sizeof(UINT) * nIDCount, FALSE));

    '// delete all existing buttons
    'int nCount = (int)DefWindowProc(TB_BUTTONCOUNT, 0, 0);
    'while (nCount--)
    '    VERIFY(DefWindowProc(TB_DELETEBUTTON, 0, 0));

    'TBBUTTON button; memset(&button, 0, sizeof(TBBUTTON));
    Dim button As TBBUTTON
    button.iString = -1

    If (Not IsEmpty(lpIDArray)) Then
        ' add new buttons to the common control
        Dim iImage As Integer
        iImage = 0

        For i = 0 To nIDCount - 1
            button.fsState = TBSTATE_ENABLED
            button.idCommand = lpIDArray(i)
            If (button.idCommand = 0) Then
                ' separator
                button.fsStyle = TBSTYLE_SEP
                ' width of separator includes 8 pixel overlap
                'ASSERT(_afxComCtlVersion != -1);
                If (((GetStyle(m_hWnd) And TBSTYLE_FLAT) <> 0) Or (afxComCtlVersion = VERSION_IE4)) Then
                    button.iBitmap = 6
                Else
                    button.iBitmap = 8
                End If
            Else
                ' a command button with image
                button.fsStyle = TBSTYLE_BUTTON
                button.iBitmap = iImage
                iImage = iImage + 1
            End If

            If (SendMessage(m_hWnd, TB_ADDBUTTONS, 1, button) = 0) Then
                SetButtons = False
                Exit Function
            End If

        Next i
    Else
        ' add 'blank' buttons
        button.fsState = TBSTATE_ENABLED

        For i = 0 To nIDCount - 1
            'ASSERT(button.fsStyle == TBSTYLE_BUTTON);
            If (SendMessage(m_hWnd, TB_ADDBUTTONS, 1, button) = 0) Then
                SetButtons = False
                Exit Function
            End If
        Next i
    End If

    m_nCount = SendMessage(m_hWnd, TB_BUTTONCOUNT, 0, 0)

    m_bDelayedButtonLayout = True
    SetButtons = True
End Function


Public Function AddReplaceBitmap(hbmImageWell As Long) As Boolean
    ' need complete bitmap size to determine number of images
    Dim bitmapa As BITMAP

    Call GetObjectA(hbmImageWell, Len(bitmapa), bitmapa)

    ' add the bitmap to the common control toolbar
    Dim bResult As Boolean

    If (m_hbmImageWell = 0) Then
        Dim addBitmap As TBADDBITMAP

        addBitmap.hInst = 0 ' makes TBADDBITMAP::nID behave a HBITMAP
        addBitmap.nID = hbmImageWell

        bResult = (SendMessage(m_hWnd, TB_ADDBITMAP, _
            bitmapa.bmWidth / m_sizeImage.cx, addBitmap) = 0)
    Else
        Dim replaceBitmap As TBREPLACEBITMAP
        replaceBitmap.hInstOld = 0
        replaceBitmap.nIDOld = m_hbmImageWell
        replaceBitmap.hInstNew = 0
        replaceBitmap.nIDNew = hbmImageWell
        replaceBitmap.nButtons = bitmapa.bmWidth / m_sizeImage.cx
        bResult = CBool(SendMessage(m_hWnd, TB_REPLACEBITMAP, 0, _
            replaceBitmap))
    End If

    ' remove old bitmap, if present
    If (bResult <> 0) Then
        Call DeleteObject(m_hbmImageWell)
        m_hbmImageWell = hbmImageWell
    End If

    AddReplaceBitmap = bResult

End Function


Private Sub Subclass()
    If mWndProcOrg Then Exit Sub
    mWndProcOrg = SetWindowLong(m_hWnd, GWL_WNDPROC, _
                        AddressOf TB_SubWndProc)
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


Public Function OnNcCalcSize(bCalcValidRects As Boolean, lpncs As Long)
    Dim lpncsp As NCCALCSIZE_PARAMS
    Call CopyMemory(lpncsp, lpncs, Len(lpncsp))

    ' calculate border space (will add to top/bottom, subtract from right/bottom)
    Dim recta As RECT
    Call SetRectEmpty(recta)

    Dim bHorz As Boolean
    bHorz = CBool((m_dwStyle And CBRS_ORIENT_HORZ) <> 0)

    Call ControlBar.CalcInsideRect(recta, bHorz)

    'ASSERT(_afxComCtlVersion != -1);
    'ASSERT(_afxComCtlVersion >= VERSION_IE4 || rect.top >= 2);

    ' adjust non-client area for border space
    lpncsp.rgrc(0).Left = lpncsp.rgrc(0).Left + recta.Left
    lpncsp.rgrc(0).Top = lpncsp.rgrc(0).Top + recta.Top

    ' previous versions of COMCTL32.DLL had a built-in 2 pixel border
    'If (afxComCtlVersion < VERSION_IE4) Then
    '    lpncsp.rgrc(0).Top = lpncsp.rgrc(0).Top - 2
    'End If

    lpncsp.rgrc(0).Right = lpncsp.rgrc(0).Right + recta.Right
    lpncsp.rgrc(0).Bottom = lpncsp.rgrc(0).Bottom + recta.Bottom

    Call CopyMemory(lpncs, lpncsp, Len(lpncsp))

End Function


Friend Function WindowPROC(ByVal hWnd As Long, _
         ByVal uMsg As Long, ByVal wParam As Long, _
         lParam As Long) As Long

    Select Case uMsg
        Case WM_WINDOWPOSCHANGING
            Dim lpwp As WINDOWPOS
            Call CopyMemory(lpwp, lParam, Len(lpwp))
            Call OnWindowPosChanging(lpwp)
            Call CopyMemory(lParam, lpwp, Len(lpwp))
            WindowPROC = 0
            Exit Function

        Case WM_MOUSEACTIVATE
            Dim lp As Long
            Call CopyMemory(lp, VarPtr(lParam), Len(lp))
            WindowPROC = OnMouseActivate(wParam, GetLoWord(lp), GetHiWord(lp))

        Case WM_PAINT
            Call OnPaint
            WindowPROC = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, lParam)

        Case WM_NCCALCSIZE
            Dim b As Boolean
            If wParam <> 0 Then b = True Else b = False
            Call OnNcCalcSize(b, lParam)
            Exit Function

        Case WM_NCPAINT
            Call EraseNonClient
            WindowPROC = 0
            Exit Function

        Case WM_LBUTTONDBLCLK
            Dim pt As POINTAPI
            Call GetCursorPos(pt)
            Call ScreenToClient(hWnd, pt)
            Call ControlBar.OnLButtonDblClk(wParam, pt, this)

        Case WM_LBUTTONDOWN
            ' only start dragging if clicked in "void" space
            Dim tmp As tagTOOLINFO

            Call GetCursorPos(pt)
            Call ScreenToClient(hWnd, pt)

            If ((Not (m_pDockBar Is Nothing)) And (OnToolHitTest(pt, False, tmp) = -1)) Then
                ' start the drag
                'ASSERT(m_pDockContext != NULL);
                Call ClientToScreen(hWnd, pt)
                Call m_pDockContext.StartDrag(pt)
            Else
                'CWnd::OnLButtonDown(nFlags, pt);
                WindowPROC = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, lParam)
            End If

            Exit Function

        Case Else
            WindowPROC = CallWindowProc(mWndProcOrg, hWnd, uMsg, wParam, lParam)
    End Select

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
    ' a dynamically resizeable toolbar can not have the CBRS_FLOAT_MULTI
    'ASSERT(!((dwNewStyle & CBRS_SIZE_DYNAMIC) &&
    '        (m_dwDockStyle & CBRS_FLOAT_MULTI)));

    ' a toolbar can not be both dynamic and fixed in size
    'ASSERT (!((dwNewStyle & CBRS_SIZE_FIXED) &&
    '    (dwNewStyle & CBRS_SIZE_DYNAMIC)));

    ' CBRS_SIZE_DYNAMIC can not be disabled once it has been enabled
    'ASSERT (((dwOldStyle & CBRS_SIZE_DYNAMIC) == 0) ||
    '    ((dwNewStyle & CBRS_SIZE_DYNAMIC) != 0));

    If ((m_hWnd <> 0) And ( _
            ((dwOldStyle And CBRS_BORDER_ANY) <> 0) <> _
            ((dwNewStyle And CBRS_BORDER_ANY) <> 0))) Then
        '// recalc non-client area when border styles change
        Call SetWindowPos(0, 0, 0, 0, 0, _
            SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE Or SWP_DRAWFRAME)
    End If
    m_bDelayedButtonLayout = True

End Function


Public Function EnableDocking(dwDockStyle As Long)
    ' must be CBRS_ALIGN_XXX or CBRS_FLOAT_MULTI only
    'ASSERT((dwDockStyle & ~(CBRS_ALIGN_ANY|CBRS_FLOAT_MULTI)) == 0);
    ' CBRS_SIZE_DYNAMIC toolbar cannot have the CBRS_FLOAT_MULTI style
    'ASSERT(((dwDockStyle & CBRS_FLOAT_MULTI) == 0) || ((m_dwStyle & CBRS_SIZE_DYNAMIC) == 0));

    m_dwDockStyle = dwDockStyle
    If (m_pDockContext Is Nothing) Then
        m_pDockContext = New CDockContext
        Call m_pDockContext.Initialize(this)
    End If

    ' permanently wire the bar's owner to its current parent
    If (m_hWndOwner = Null) Then
        m_hWndOwner = Globals.GetParent(m_hWnd)
    End If

End Function


Private Property Get this() As CToolbar
   Set this = Me
End Property


Friend Function OnToolHitTest(point As POINTAPI, bool_pTI As Boolean, pTI As tagTOOLINFO) As Long
    'ASSERT_VALID(this);
    'ASSERT(::IsWindow(m_hWnd));

    '// check child windows first by calling CControlBar
    'int nHit = CControlBar::OnToolHitTest(point, pTI);
    'if (nHit != -1)
    '    return nHit;

    ' now hit test against CToolBar buttons
    Dim nButtons As Long
    nButtons = SendMessage(m_hWnd, TB_BUTTONCOUNT, 0, 0)

    Dim i As Long
    For i = 0 To (nButtons - 1)
        Dim recta As RECT
        Dim button As TBBUTTON

        If SendMessage(m_hWnd, TB_GETITEMRECT, i, recta) <> 0 Then

            recta.Bottom = recta.Bottom + 1
            recta.Right = recta.Right + 1

            Call SendMessage(m_hWnd, TB_GETBUTTON, i, button)

            If (PtInRectNEW(recta, point) And _
                    (Not ((button.fsStyle And TBSTYLE_SEP) <> 0))) Then
                Dim nHit As Long
                nHit = GetItemID(i)

                Dim tmp As tagAFX_OLDTOOLINFO
                If ((bool_pTI <> False) And pTI.cbSize >= Len(tmp)) <> 0 Then

                    pTI.hWnd = m_hWnd
                    pTI.recta = recta
                    pTI.uId = nHit
                    pTI.lpszText = 0 'pTI.lpszText = LPSTR_TEXTCALLBACK
                End If

                ' found matching rect, return the ID of the button
                If (nHit <> 0) Then OnToolHitTest = nHit Else OnToolHitTest = -1
                Exit Function

            End If
        End If

    Next i

    OnToolHitTest = -1
End Function


Friend Function SetCapture() As Long
    SetCapture = Globals.SetCapture(m_hWnd)
End Function


Public Function GetItemID(nIndex As Long) As Long
    'ASSERT_VALID(this);
    'ASSERT(::IsWindow(m_hWnd));

    Dim button As TBBUTTON
    Call GetButton(nIndex, button)

    GetItemID = button.idCommand
End Function


Public Property Get IsDockBar() As Boolean
    IsDockBar = False
End Property


Public Property Get IsWindowVisible() As Boolean
    IsWindowVisible = CBool(Globals.IsWindowVisible(m_hWnd))
End Property


Public Function GetBarStyle() As Long
    GetBarStyle = m_dwStyle
End Function


Public Function SetWindowPos(hWndInsertAfter As Long, x As Long, y As Long, cx As Long, cy As Long, wFlags As Long) As Long
    SetWindowPos = Globals.SetWindowPos(m_hWnd, hWndInsertAfter, x, y, cx, cy, wFlags)
End Function


Public Property Get GetParent() As Long
    GetParent = Globals.GetParent(m_hWnd)
End Property


Public Function SetParent(newParent As Long)
    Call Globals.SetParent(m_hWnd, newParent)
End Function


Private Function layout()
    'ASSERT(m_bDelayedButtonLayout);

    m_bDelayedButtonLayout = False

    Dim bHorz As Boolean
    bHorz = (m_dwStyle And CBRS_ORIENT_HORZ) <> 0

    If (((m_dwStyle And CBRS_FLOATING) <> 0) And ((m_dwStyle And CBRS_SIZE_DYNAMIC) <> 0)) Then
        Call this.CalcDynamicLayout(0, LM_HORZ Or LM_MRUWIDTH Or LM_COMMIT)
    ElseIf (bHorz) Then
        Call this.CalcDynamicLayout(0, LM_HORZ Or LM_HORZDOCK Or LM_COMMIT)
    Else
        Call this.CalcDynamicLayout(0, LM_VERTDOCK Or LM_COMMIT)
    End If

End Function


Private Function OnPaint()

    If (m_bDelayedButtonLayout) Then
        Call layout
    End If

    'Call Default
End Function


Friend Function OnMouseActivate(pDesktopWnd As Long, nHitTest As Integer, nMsg As Integer) As Long
    OnMouseActivate = ControlBar.OnMouseActivate(pDesktopWnd, nHitTest, nMsg)
End Function


Friend Function EraseNonClient()
    Call ControlBar.EraseNonClient(Nothing)
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

        Call Invalidate
    End If
End Function


' determine number of elements in an array
Private Function countof_(arr() As Variant) As Long
    countof_ = (UBound(arr) - LBound(arr))
End Function


Private Function GetRValue(colRGB As Long) As Integer
    Dim t As RGBComponents
    Dim a As AllLong
    a.a = colRGB
    LSet t = a
    GetRValue = t.r
End Function


Private Function GetGValue(colRGB As Long) As Integer
    Dim t As RGBComponents
    Dim a As AllLong
    a.a = colRGB
    LSet t = a
    GetGValue = t.g
End Function


Private Function GetBValue(colRGB As Long) As Integer
    Dim t As RGBComponents
    Dim a As AllLong
    a.a = colRGB
    LSet t = a
    GetBValue = t.b
End Function


Private Function RGB_TO_RGBQUAD(r As Integer, g As Integer, b As Integer) As COLORREF
    RGB_TO_RGBQUAD = RGB(b, g, r)
End Function


Private Function CLR_TO_RGBQUAD(clr As Long) As COLORREF
    CLR_TO_RGBQUAD = RGB(GetBValue(clr), GetGValue(clr), GetRValue(clr))
End Function


Private Function RGB(r As Integer, g As Integer, b As Integer) As COLORREF
    Dim t As RGBComponents
    Dim a As AllLong
    t.r = r
    t.g = g
    t.b = b
    t.unknown = 0
    LSet a = t

    RGB.x = a.a
End Function


Private Function Init()
    ' mapping from color in DIB to system color
    afxSysColorMap_(0).rgbqFrom = RGB_TO_RGBQUAD(&H0, &H0, &H0).x       ' black
    afxSysColorMap_(0).iSysColorTo = COLOR_BTNTEXT

    afxSysColorMap_(1).rgbqFrom = RGB_TO_RGBQUAD(&H80, &H80, &H80).x    ' dark gray
    afxSysColorMap_(1).iSysColorTo = COLOR_BTNSHADOW

    afxSysColorMap_(2).rgbqFrom = RGB_TO_RGBQUAD(&HC0, &HC0, &HC0).x    ' bright gray
    afxSysColorMap_(2).iSysColorTo = COLOR_BTNFACE

    afxSysColorMap_(3).rgbqFrom = RGB_TO_RGBQUAD(&HFF, &HFF, &HFF).x    ' white
    afxSysColorMap_(3).iSysColorTo = COLOR_BTNHIGHLIGHT

End Function


Private Function AfxLoadSysColorBitmap(hInst As Long, hRsrc As Long, Optional bMono As Boolean = False) As Long
    Call Init

    Dim hglb As Long
    hglb = LoadResource(hInst, hRsrc)
    If (hglb = 0) Then
        AfxLoadSysColorBitmap = 0
        Exit Function
    End If

    Dim lpBitmap As Long
    lpBitmap = LockResource(hglb)

    If (lpBitmap = 0) Then
        AfxLoadSysColorBitmap = 0
        Exit Function
    End If

    Dim lpBitmapa As BITMAPINFOHEADER
    Call CopyMemory(lpBitmapa, ByVal lpBitmap, Len(lpBitmapa))

    ' make copy of BITMAPINFOHEADER so we can modify the color table
    Const nColorTableSize As Long = 16

    Dim nSize As Long
    Dim rgbQ As RGBQUAD
    nSize = lpBitmapa.biSize + (nColorTableSize * Len(rgbQ))

    Dim lpBitmapInfo() As Byte
    ReDim lpBitmapInfo(0 To (nSize - 1)) As Byte

    Call CopyMemory(ByVal VarPtr(lpBitmapInfo(0)), ByVal lpBitmap, nSize)

    ' color table is in RGBQUAD DIB format

    Dim lpBitmapInfoa As BITMAPINFOHEADER
    Call CopyMemory(lpBitmapInfoa, ByVal lpBitmap, Len(lpBitmapInfoa))
    
    Dim pColorTable_Offset As Long
    pColorTable_Offset = VarPtr(lpBitmapInfo(0)) + lpBitmapa.biSize 'lpBitmapInfo.biSize

    Dim pColorTable() As Long

    ReDim pColorTable(0 To (nColorTableSize - 1)) As Long
    Call CopyMemory(ByVal VarPtr(pColorTable(0)), ByVal pColorTable_Offset, nColorTableSize * 4)  '**UNKNOWN - Don't know why 'nColorTableSize' copies only 4 elements, but 'nColorTableSize * 4' copies all 16 elements into the BYTE array. It seems like CopyMemory is copying LONG values to the BYTE array, as sizeof(LONG)=4

    Dim iColor As Integer
    For iColor = 0 To nColorTableSize - 1
        ' look for matching RGBQUAD color in original
        Dim i As Integer
        For i = 0 To (UBound(afxSysColorMap_) - LBound(afxSysColorMap_))
            If (pColorTable(iColor) = afxSysColorMap_(i).rgbqFrom) Then
                If (bMono) Then
                    ' all colors except text become white
                    If (afxSysColorMap_(i).iSysColorTo <> COLOR_BTNTEXT) Then
                        pColorTable(iColor) = RGB_TO_RGBQUAD(255, 255, 255).x
                    End If
                Else
                    pColorTable(iColor) = _
                        CLR_TO_RGBQUAD(GetSysColor(afxSysColorMap_(i).iSysColorTo)).x
                End If

                Exit For
            End If
        Next i
    Next iColor

    Call CopyMemory(ByVal pColorTable_Offset, ByVal VarPtr(pColorTable(0)), nColorTableSize * 4)  '**UNKNOWN - Don't know why 'nColorTableSize' copies only 4 elements, but 'nColorTableSize * 4' copies all 16 elements into the BYTE array. It seems like CopyMemory is copying LONG values to the BYTE array, as sizeof(LONG)=4

    Dim nWidth As Long
    nWidth = lpBitmapInfoa.biWidth
    Dim nHeight As Long
    nHeight = lpBitmapInfoa.biHeight

    Dim hDCScreen As Long
    hDCScreen = GetDC(0&)

    Dim hbm As Long
    hbm = CreateCompatibleBitmap(hDCScreen, nWidth, nHeight)

    If (hbm <> 0) Then
        Dim hDCGlyphs As Long
        hDCGlyphs = CreateCompatibleDC(hDCScreen)

        Dim hbmOld As Long
        hbmOld = SelectObject(hDCGlyphs, hbm)

        Dim lpBits As Long
        lpBits = (lpBitmap + Len(lpBitmapa)) '(lpBitmap + 1) in C/C++ is not the same as (lpBitmap + 1) in VB. Since 'lpBitmap' is a pointer, according to pointer arithmetic in C/C++, each increment/decrement by 1 MEANS incrementing/decrementing the pointer by the 'sizeof(the type pointed to by that pointer)'. Thus the VB equivalent is 'lpBits = (lpBitmap + Len(lpBitmapa))', where 'lpBitmapa' is a variable (structure in C/C++ terminology) of type BITMAPINFOHEADER

        lpBits = lpBits + (LShiftWord(1, lpBitmapInfoa.biBitCount)) * Len(rgbQ)
        Call StretchDIBits_LONG(hDCGlyphs, 0, 0, nWidth, nHeight, 0, 0, nWidth, nHeight, _
                ByVal lpBits, ByVal VarPtr(lpBitmapInfo(0)), DIB_RGB_COLORS, SRCCOPY)
        Call SelectObject(hDCGlyphs, hbmOld)
        Call DeleteDC(hDCGlyphs)
    End If

    Call ReleaseDC(0&, hDCScreen)

    ' free copy of bitmap info struct and resource itself
    Call FreeResource(hglb)

    AfxLoadSysColorBitmap = hbm
End Function


Public Function LoadBitmap(lpszResourceName As String) As Boolean
    'ASSERT_VALID(this);
    'ASSERT(lpszResourceName != NULL);

    ' determine location of the bitmap in resource fork
    Dim hInstImageWell As Long
    hInstImageWell = App.hInstance

    Dim hRsrcImageWell As Long
    hRsrcImageWell = FindResourceNEW(0, CLng(lpszResourceName), RT_BITMAP)

    If (hRsrcImageWell = 0) Then
        LoadBitmap = False
        Exit Function
    End If

    ' load the bitmap
    Dim hbmImageWell As Long
    hbmImageWell = AfxLoadSysColorBitmap(hInstImageWell, hRsrcImageWell)

    ' tell common control toolbar about the new bitmap
    If (Not AddReplaceBitmap(hbmImageWell)) Then
        LoadBitmap = False
        Exit Function
    End If

    ' remember the resource handles so the bitmap can be recolored if necessary
    m_hInstImageWell = hInstImageWell
    m_hRsrcImageWell = hRsrcImageWell

    LoadBitmap = True
End Function


Friend Function IsFloating() As Boolean
    'if (IsDockBar())
    '    return ((CDockBar*)this)->m_bFloating;
    'else
    '    return m_pDockBar != NULL && m_pDockBar->m_bFloating;
    IsFloating = Not (m_pDockBar Is Nothing) And m_pDockBar.m_bFloating
End Function


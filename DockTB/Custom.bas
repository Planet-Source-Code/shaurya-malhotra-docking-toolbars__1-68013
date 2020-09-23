Attribute VB_Name = "Custom"
' ----------------------------------------------------------------- '
' Filename: Custom.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Miscellaneous functions
' ----------------------------------------------------------------- '

Option Explicit

'-----------------------------------------------------------------
Public Const DIB_RGB_COLORS As Long = 0     ' color table in RGBs
Public Const DIB_PAL_COLORS As Long = 1     ' color table in palette indices

Public Const SRCCOPY = &HCC0020
'-----------------------------------------------------------------

Public Const CX_BORDER = 1
Public Const CY_BORDER = 1

Public Const cxBorder2 = CX_BORDER * 2
Public Const cyBorder2 = CY_BORDER * 2

Public Const CX_GRIPPER = 3
Public Const CY_GRIPPER = 3
Public Const CX_BORDER_GRIPPER = 2
Public Const CY_BORDER_GRIPPER = 2

Public Type AllLong
    a As Long
End Type

Private Type HiLoWord
    l As Integer
    h As Integer
End Type

Public Type RGBComponents
    r As Byte
    g As Byte
    b As Byte
    unknown As Byte
End Type

Enum afxData
    bSmCaption = 1
    m_bForceFrame = 0
End Enum

Public Type CollObject
    obj As Object
    ptr As Long
End Type

Public Enum AdjustType
    adjustBorder = 0
    adjustOutside = 1
End Enum

Public Enum RepositionFlags
    reposDefault = 0
    reposQuery = 1
    reposExtra = 2
End Enum


Public Function FromHandlePermanent(hWnd As Long) As Long
    FromHandlePermanent = hWnd
End Function


Public Function FromHandle(hWnd As Long) As Long
    FromHandle = hWnd
End Function


Public Function CalcWindowRect(hWnd As Long, lpClientRect As RECT, Optional nAdjustType As Long = adjustBorder)
    Dim dwExStyle As Long
    dwExStyle = GetExStyle(hWnd)

    If (nAdjustType = 0) Then
        dwExStyle = dwExStyle And (Not WS_EX_CLIENTEDGE)
    End If

    Call AdjustWindowRectEx(lpClientRect, GetStyle(hWnd), False, WS_EX_TOOLWINDOW)
End Function


Public Function GetExStyle(hWnd As Long)
    Dim ret As Long
    ret = GetWindowLong(hWnd, GWL_EXSTYLE)
    GetExStyle = ret
End Function


Public Function GetTopLeft(r As RECT) As POINTAPI
    Dim pt As POINTAPI
    pt.x = r.Left
    pt.y = r.Top

    GetTopLeft = pt
End Function

Public Function GetBottomRight(r As RECT) As POINTAPI
    Dim pt As POINTAPI
    pt.x = r.Right
    pt.y = r.Bottom

    GetBottomRight = pt
End Function


Public Function Assert(Msg As String)
    Beep
    MsgBox Msg, vbCritical, "Assertion"
End Function


Public Function GetParentFrameMFC(hWnd As Long) As Long
    Dim pParentWnd As Long
    pParentWnd = GetParent(hWnd)        ' start with one parent up

    Do While (pParentWnd <> 0)
        If Not (WindowExistsInList(pParentWnd) Is Nothing) Then
            If WindowExistsInList(pParentWnd).IsFrameWnd() Then
                GetParentFrameMFC = pParentWnd
                Exit Function
            End If
        Else
            GetParentFrameMFC = pParentWnd
            Exit Function
        End If

        pParentWnd = GetParent(pParentWnd)
    Loop

    GetParentFrameMFC = 0
End Function

Public Function CSize(r As RECT) As Size
    Dim s As Size
    s.cx = r.Right - r.Left
    s.cy = r.Bottom - r.Top

    CSize = s
End Function


Public Function AfxAdjustRectangle(ByRef recta As RECT, pt As POINTAPI)
    Dim nXOffset As Integer

    If (pt.x < recta.Left) Then
        nXOffset = (pt.x - recta.Left)
    Else
        If (pt.x > recta.Right) Then
            nXOffset = (pt.x - recta.Right)
        Else
            nXOffset = 0
        End If
    End If

    Dim nYOffset As Integer

    If (pt.y < recta.Top) Then
        nYOffset = (pt.y - recta.Top)
    Else
        If (pt.y > recta.Bottom) Then
            nYOffset = (pt.y - recta.Bottom)
        Else
            nYOffset = 0
        End If
    End If

    Call OffsetRect(recta, nXOffset, nYOffset)
End Function


Public Function CalcBorders(ByRef lpClientRect As RECT, Optional dwStyle As Long = WS_THICKFRAME Or WS_CAPTION, Optional dwExStyle As Long = 0)
    If (afxData.bSmCaption <> 0) Then
        Call AdjustWindowRectEx(lpClientRect, dwStyle, False, WS_EX_PALETTEWINDOW)
        Exit Function
    End If

    If (dwStyle And (MFS_4THICKFRAME Or WS_THICKFRAME Or MFS_THICKFRAME)) <> 0 Then
        Call InflateRect(lpClientRect, GetSystemMetrics(SM_CXFRAME), GetSystemMetrics(SM_CYFRAME))
    Else
        Call InflateRect(lpClientRect, GetSystemMetrics(SM_CXBORDER), GetSystemMetrics(SM_CYBORDER))
    End If

    If (dwStyle And WS_CAPTION) <> 0 Then
        'lpClientRect.Top = lpClientRect.Top - afx_sizeMiniSys.cy
    End If
End Function


Public Function CRectNEW(x As Long, y As Long, sizea As Size) As RECT
    Dim ret As RECT

    ret.Left = x
    ret.Right = ret.Left + sizea.cx

    ret.Top = y
    ret.Bottom = ret.Top + sizea.cy

    CRectNEW = ret
End Function


Public Function AfxModifyStyle(hWnd As Long, nStyleOffset As Long, _
                        dwRemove As Long, dwAdd As Long, nFlags As Long) As Boolean
    Dim dwStyle As Long
    dwStyle = GetWindowLong(hWnd, nStyleOffset)

    Dim dwNewStyle As Long
    dwNewStyle = (dwStyle And Not (dwRemove)) Or dwAdd

    If (dwStyle = dwNewStyle) Then
        AfxModifyStyle = False
        Exit Function
    End If

    Call SetWindowLong(hWnd, nStyleOffset, dwNewStyle)

    If (nFlags <> 0) Then
        Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or nFlags)
    End If

    AfxModifyStyle = True
End Function


Public Function ModifyStyle(hWnd As Long, dwRemove As Long, dwAdd As Long, Optional nFlags As Long = 0) As Boolean
    ModifyStyle = AfxModifyStyle(hWnd, GWL_STYLE, dwRemove, dwAdd, nFlags)
End Function


Public Function GetStyle(hWnd As Long)
    Dim ret As Long
    ret = GetWindowLong(hWnd, GWL_STYLE)
    GetStyle = ret 'Or FWS_SNAPTOBARS
End Function


Public Function GetMax(x As Long, y As Long) As Long
    If (x > y) Then GetMax = x Else GetMax = y
End Function

Public Function GetMin(x As Long, y As Long) As Long
    If (x < y) Then GetMin = x Else GetMin = y
End Function


Public Function ClientToScreenRect(hWnd As Long, ByRef recta As RECT)
    Dim s1 As POINTAPI, s2 As POINTAPI
    s1.x = recta.Left
    s1.y = recta.Top

    s2.x = recta.Right
    s2.y = recta.Bottom

    Call ClientToScreen(hWnd, s1)
    Call ClientToScreen(hWnd, s2)

    recta.Left = s1.x
    recta.Top = s1.y

    recta.Right = s2.x
    recta.Bottom = s2.y
End Function


Public Function MAKELONG(l As Integer, h As Integer) As Long
    Dim a As AllLong
    Dim hl As HiLoWord
    hl.h = h
    hl.l = l

    LSet a = hl
    MAKELONG = a.a
End Function


Public Function CRect(l As Long, t As Long, r As Long, b As Long) As RECT
    Dim tmp As RECT
    tmp.Left = l
    tmp.Top = t
    tmp.Right = r
    tmp.Bottom = b

    CRect = tmp
End Function


Public Function CPtRect(point As POINTAPI, sizea As Size) As RECT
    Dim ret As RECT

    ret.Left = point.x
    ret.Right = ret.Left + sizea.cx

    ret.Top = point.y
    ret.Bottom = ret.Top + sizea.cy

    CPtRect = ret
End Function


Public Function ScreenToClientRect(hWnd As Long, ByRef r As RECT)
    Dim p1 As POINTAPI, p2 As POINTAPI

    p1.x = r.Left
    p1.y = r.Top
    
    p2.x = r.Right
    p2.y = r.Bottom

    Call ScreenToClient(hWnd, p1)
    Call ScreenToClient(hWnd, p2)

    r.Left = p1.x
    r.Top = p1.y
    r.Right = p2.x
    r.Bottom = p2.y
End Function


Public Function GetWidth(r As RECT) As Long
    GetWidth = (r.Right - r.Left)
End Function

Public Function GetHeight(r As RECT) As Long
    GetHeight = (r.Bottom - r.Top)
End Function


Public Function IsRectEq(r1 As RECT, r2 As RECT) As Boolean
    If ((r1.Left = r2.Left) And (r1.Top = r2.Top) And (r1.Right = r2.Right) And (r1.Bottom = r2.Bottom)) Then
        IsRectEq = True
    Else
        IsRectEq = False
    End If
End Function


Public Function AfxRepositionWindow(lpLayout As AFX_SIZEPARENTPARAMS, _
                                                hWnd As Long, lpRect As RECT)
    Dim hWndParent As Long
    hWndParent = GetParent(hWnd)

    'if (lpLayout != NULL && lpLayout->hDWP == NULL)
    '    return;

    ' first check if the new rectangle is the same as the current
    Dim rectOld As RECT
    Call GetWindowRect(hWnd, rectOld)

    Dim ptTopLeft As POINTAPI, ptBottomRight As POINTAPI
    ptTopLeft = GetTopLeft(rectOld)
    ptBottomRight = GetBottomRight(rectOld)
    Call ScreenToClient(hWndParent, ptTopLeft)
    Call ScreenToClient(hWndParent, ptBottomRight)
    rectOld.Top = ptTopLeft.y
    rectOld.Left = ptTopLeft.x
    rectOld.Bottom = ptBottomRight.y
    rectOld.Right = ptBottomRight.x

    If (EqualRect(rectOld, lpRect)) Then
        Exit Function               ' nothing to do
    End If

    ' try to use DeferWindowPos for speed, otherwise use SetWindowPos
    If (lpLayout.hDWP <> 0) Then
        lpLayout.hDWP = DeferWindowPos(lpLayout.hDWP, hWnd, 0, _
                lpRect.Left, lpRect.Top, lpRect.Right - lpRect.Left, _
                lpRect.Bottom - lpRect.Top, SWP_NOACTIVATE Or SWP_NOZORDER)
    Else
        Call SetWindowPos(hWnd, 0, lpRect.Left, lpRect.Top, _
                lpRect.Right - lpRect.Left, lpRect.Bottom - lpRect.Top, _
                SWP_NOACTIVATE Or SWP_NOZORDER)
    End If
End Function


Public Function PtInRectNEW(lpRect As RECT, pt As POINTAPI) As Boolean
    If ((pt.x >= lpRect.Left And pt.x <= lpRect.Right) And (pt.y >= lpRect.Top And pt.y <= lpRect.Bottom)) Then
        PtInRectNEW = True
    Else
        PtInRectNEW = False
    End If
End Function


Public Function AfxGetDlgCtrlID(hWnd As Long) As Long
    AfxGetDlgCtrlID = GetDlgCtrlID(hWnd)
End Function


Public Function GetParentFrame(hWnd As Long) As CFrame
    Set GetParentFrame = WindowExistsInList(GetParentFrameMFC(hWnd))
End Function


Public Function CRECTtoSize(r As RECT) As Size
    Dim ret As Size
    ret.cx = (r.Right - r.Left)
    ret.cy = (r.Bottom - r.Top)
    CRECTtoSize = ret
End Function


Public Function PtSizeToRECT(point As POINTAPI, sizea As Size) As RECT
    Dim ret As RECT

    ret.Left = point.x
    ret.Right = ret.Left + sizea.cx

    ret.Top = point.y
    ret.Bottom = ret.Top + sizea.cy

    PtSizeToRECT = ret
End Function


Public Function GetHiWord(ByVal x As Long) As Integer
    Dim a As AllLong
    a.a = x
    Dim hl As HiLoWord

    LSet hl = a
    GetHiWord = hl.h
End Function

Public Function GetLoWord(ByVal x As Long) As Integer
    Dim a As AllLong
    a.a = x
    Dim hl As HiLoWord

    LSet hl = a
    GetLoWord = hl.l
End Function


Public Function AfxIsDescendant(ByVal hWndParent As Long, ByVal hWndChild As Long) As Boolean
    ' helper for detecting whether child descendent of parent
    ' (works with owned popups as well)

    'ASSERT(::IsWindow(hWndParent));
    'ASSERT(::IsWindow(hWndChild));

    Do
        If (hWndParent = hWndChild) Then
            AfxIsDescendant = True
            Exit Function
        End If

        hWndChild = AfxGetParentOwner(hWndChild)
    Loop While (hWndChild <> 0)

    AfxIsDescendant = False
End Function


Public Function AfxGetParentOwner(hWnd As Long)
    ' check for permanent-owned window first
    Dim pWnd As Long
    pWnd = hWnd

    If (pWnd <> 0) Then
        AfxGetParentOwner = GetOwner(pWnd)
        Exit Function
    End If

    ' otherwise, return parent in the Windows sense
    Dim tmp As Long
    If (GetWindowLong(hWnd, GWL_STYLE) And WS_CHILD) <> 0 Then
        tmp = GetParent(hWnd)
    Else
        tmp = GetWindow(hWnd, GW_OWNER)
    End If
    AfxGetParentOwner = tmp
End Function


Public Function GetOwner(pWnd As Long) As Long
'    If (m_hWndOwner <> 0) Then
'        GetOwner = m_hWndOwner
'    Else
'        GetOwner = GetParent(m_hWnd)
'    End If
    GetOwner = GetParent(pWnd)
End Function


Public Function LShiftWord(w As Integer, c As Integer) As Integer
    LShiftWord = w * (2 ^ c)
End Function

Public Function RShiftWord(w As Integer, c As Integer) As Integer
    RShiftWord = w \ (2 ^ c)
End Function



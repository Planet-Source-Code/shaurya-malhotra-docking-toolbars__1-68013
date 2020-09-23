VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMiniFrameWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CMiniFrameWnd.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Converted from MFC's CMiniFrameWnd class and adapted to
' Visual Basic
' ----------------------------------------------------------------- '

Option Explicit

Implements CFrame
Private Frame As New CFrame

Private m_bSysTracking As Boolean
Private m_bInSys As Boolean
Private m_bActive As Boolean
Private m_strCaption As String


Private Function CFrame_CreateEx(dwExStyle As Long, lpszClassName As String, lpszWindowName As String, dwStyle As Long, x As Long, y As Long, nWidth As Long, nHeight As Long, hWndParent As Long, nIDorHMenu As Long, lpParam As Long, Optional obj As Object = Nothing) As Boolean
    CFrame_CreateEx = Frame.CreateEx(dwExStyle, lpszClassName, lpszWindowName, dwStyle, x, y, nWidth, nHeight, hWndParent, nIDorHMenu, lpParam, obj)
End Function

Private Function CFrame_EnableDocking(dwDockStyle As Long) As Variant
'
End Function

Private Property Let CFrame_hWnd(RHS As Long)
'
End Property

Private Property Get CFrame_hWnd() As Long
    CFrame_hWnd = Frame.hWnd
End Property

Private Property Let CFrame_hWndOwner(ByVal RHS As Long)
'
End Property

Private Property Get CFrame_hWndOwner() As Long
    CFrame_hWndOwner = Frame.hWndOwner
End Property

Private Function CFrame_IsFrameWnd() As Boolean
    CFrame_IsFrameWnd = Frame.IsFrameWnd
End Function

Private Sub Class_Initialize()
    m_bActive = False
    Call Initialize
End Sub

Private Function Initialize()
    'to be implemented
End Function

Private Property Get m_bInRecalcLayout() As Boolean
    m_bInRecalcLayout = Frame.m_bInRecalcLayout
End Property

Private Property Let m_bInRecalcLayout(RHS As Boolean)
    Frame.m_bInRecalcLayout = RHS
End Property

Public Property Get bInRecalcLayout() As Boolean
    bInRecalcLayout = m_bInRecalcLayout
End Property

Public Property Let bInRecalcLayout(RHS As Boolean)
    m_bInRecalcLayout = RHS
End Property

Public Property Get bSysTracking() As Boolean
    bSysTracking = m_bSysTracking
End Property

Public Property Let bSysTracking(RHS As Boolean)
    m_bSysTracking = RHS
End Property

Public Property Get bInSys() As Boolean
    bInSys = m_bInSys
End Property

Public Property Let bInSys(RHS As Boolean)
    m_bInSys = RHS
End Property

Public Property Get bActive() As Boolean
    bActive = m_bActive
End Property

Public Property Let bActive(RHS As Boolean)
    m_bActive = RHS
End Property

Public Property Get strCaption() As String
    strCaption = m_strCaption
End Property

Public Property Let strCaption(RHS As String)
    m_strCaption = RHS
End Property

Public Property Get hWnd() As Long
    hWnd = Frame.m_hWnd
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = Frame.hWndOwner
End Property


Friend Function ActivateTopParent()
    Call Frame.ActivateTopParent
End Function

Friend Function RecalcLayout(Optional bNotify As Boolean = True, Optional obj As CMiniDockFrameWnd = Nothing)
    Call Frame.RecalcLayout(bNotify, obj)
End Function

Friend Function DelayRecalcLayout(Optional bNotify As Boolean = True)
    Call Frame.DelayRecalcLayout(bNotify)
End Function

Friend Function DestroyWindow()
    Call Frame.DestroyWindow
End Function


Friend Function CreateEx(dwExStyle As Long, lpClassName As String, _
                    lpszWindowName As String, dwStyle As Long, recta As RECT, _
                    pParentWnd As Long, Optional nID As Long = 0) As Boolean

    m_strCaption = lpszWindowName
    If lpClassName = "" Then lpClassName = afxMiniFrameWndClass

    CreateEx = Frame.CreateEx(dwExStyle, lpClassName, _
                    lpszWindowName, dwStyle, recta.Left, recta.Top, recta.Right - recta.Left, _
                    recta.Bottom - recta.Top, pParentWnd, nID, 0, this)
End Function


Friend Function PreCreateWindow(cs As CREATESTRUCT) As Boolean
    If (afxData.bSmCaption) Then
        ' WS_4THICKFRAME and MFS_THICKFRAME imply WS_THICKFRAME
        If (cs.style And (MFS_4THICKFRAME Or MFS_THICKFRAME)) <> 0 Then
            cs.style = cs.style Or WS_THICKFRAME
        End If

        ' WS_CAPTION implies WS_EX_TOOLWINDOW
        If (cs.style And WS_CAPTION) <> 0 Then
            cs.dwExStyle = cs.dwExStyle Or WS_EX_TOOLWINDOW
        End If
    End If

    'VERIFY(CFrameWnd::PreCreateWindow(cs));
    cs.dwExStyle = cs.dwExStyle And (Not (WS_EX_CLIENTEDGE))

    PreCreateWindow = True
End Function


Friend Function CalcWindowRect(lpClientRect As RECT, Optional nAdjustType As Long = adjustBorder)
    If (afxData.bSmCaption) Then
        Call Frame.CalcWindowRect(lpClientRect, nAdjustType)
        Exit Function
    End If

    Dim dwStyle As Long
    dwStyle = GetStyle(m_hWnd)
    If (dwStyle And (MFS_4THICKFRAME Or WS_THICKFRAME Or MFS_THICKFRAME)) Then
        Call InflateRect(lpClientRect, _
            GetSystemMetrics(SM_CXFRAME), GetSystemMetrics(SM_CYFRAME))
    Else
        Call InflateRect(lpClientRect, _
            GetSystemMetrics(SM_CXBORDER), GetSystemMetrics(SM_CYBORDER))
    End If

End Function


Private Property Get m_hWnd() As Long
    m_hWnd = Frame.m_hWnd
End Property

Private Property Get this() As CMiniFrameWnd
    Set this = Me
End Property


Friend Function OnFloatStatus(wParam As Long, lParam As Long) As Long
    ' these asserts make sure no conflicting actions are requested
    'ASSERT(!((wParam & FS_SHOW) && (wParam & FS_HIDE)));
    'ASSERT(!((wParam & FS_ENABLE) && (wParam & FS_DISABLE)));
    'ASSERT(!((wParam & FS_ACTIVATE) && (wParam & FS_DEACTIVATE)));

    ' FS_SYNCACTIVE is used to detect MFS_SYNCACTIVE windows
    Dim lResult As Long
    lResult = 0

    If (((GetStyle(m_hWnd) And MFS_SYNCACTIVE) <> 0) And ((wParam And FS_SYNCACTIVE) <> 0)) Then
        lResult = 1
    End If

    If (wParam And (FS_SHOW Or FS_HIDE)) <> 0 Then
        Dim tmp As Long
        If ((wParam And FS_SHOW) <> 0) Then
            tmp = SWP_SHOWWINDOW
        Else
            tmp = SWP_HIDEWINDOW
        End If
        Call SetWindowPos(m_hWnd, 0, 0, 0, 0, 0, _
                            tmp Or SWP_NOZORDER Or _
                            SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
    End If

    If (wParam And (FS_ENABLE Or FS_DISABLE)) <> 0 Then
        Call EnableWindow(m_hWnd, CBool((wParam And FS_ENABLE) <> 0))
    End If

    If (((wParam And (FS_ACTIVATE Or FS_DEACTIVATE)) <> 0) And _
                (GetStyle(m_hWnd) And MFS_SYNCACTIVE) <> 0) Then
        Call ModifyStyle(m_hWnd, MFS_SYNCACTIVE, 0)
        Call SendMessage(m_hWnd, WM_NCACTIVATE, CBool((wParam And FS_ACTIVATE) <> 0), 0)
        Call ModifyStyle(m_hWnd, 0, MFS_SYNCACTIVE)
    End If

    OnFloatStatus = lResult
End Function


Friend Function OnActivate(nState As Long, pWndOther As Long, bMinimized As Boolean)
    OnActivate = Frame.OnActivate(nState, pWndOther, bMinimized)
End Function

Friend Function OnNcActivate(bActive As Boolean) As Boolean
    If ((GetStyle(m_hWnd) And MFS_SYNCACTIVE) = 0) Then
        If (afxData.bSmCaption) Then
            OnNcActivate = DefWindowProc(m_hWnd, WM_NCACTIVATE, bActive, 0)
            Exit Function
        End If

        If (m_bActive <> bActive) Then
            m_bActive = bActive
            Call SendMessage(m_hWnd, WM_NCPAINT, 0, 0)
        End If
    ElseIf (m_nFlags And WF_KEEPMINIACTIVE) <> 0 Then
        OnNcActivate = False
        Exit Function
    End If

    OnNcActivate = True
End Function


Friend Function OnNcCreate(lpcs As CREATESTRUCT) As Boolean
    If (Not Frame.OnNcCreate(lpcs)) Then
        Exit Function
    End If

    If (GetStyle(m_hWnd) And MFS_SYNCACTIVE) <> 0 Then
        ' syncronize activation state with top level parent
        Dim pParentWnd As Long
        pParentWnd = GetTopLevelParent()
        'ASSERT(pParentWnd != NULL);
        Dim pActiveWnd As Long
        pActiveWnd = GetForegroundWindow()

        Dim bActive As Boolean
        If (pParentWnd = pActiveWnd) Then
            bActive = True
        Else
            If (GetLastActivePopup(pParentWnd) = pActiveWnd) Then
                If (SendMessage(pActiveWnd, WM_FLOATSTATUS, FS_SYNCACTIVE, 0&) <> 0) Then
                    bActive = True
                Else
                    bActive = False
                End If
            Else
                bActive = False
            End If
        End If

        ' the WM_FLOATSTATUS does the actual work
        Call SendMessage(m_hWnd, WM_FLOATSTATUS, IIf(bActive, FS_ACTIVATE, FS_DEACTIVATE), 0&)
    End If

    OnNcCreate = True
End Function


Friend Function GetTopLevelParent() As Long
    GetTopLevelParent = Frame.GetTopLevelParent()
End Function


Private Property Get m_nFlags() As Long
    m_nFlags = Frame.m_nFlags
End Property

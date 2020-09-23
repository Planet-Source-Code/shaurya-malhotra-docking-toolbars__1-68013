Attribute VB_Name = "Subclass"
' ----------------------------------------------------------------- '
' Filename: Subclass.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Subclassing functions
' ----------------------------------------------------------------- '

Option Explicit


Public Function TB_SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                            ByVal wParam As Long, lParam As Long) As Long
    On Error Resume Next

    Dim Toolbar As CToolbar
    Dim ptrObject As Long

    ptrObject = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory Toolbar, ptrObject, 4

    TB_SubWndProc = Toolbar.WindowPROC(hWnd, Msg, _
    wParam, lParam)
    
    CopyMemory Toolbar, 0&, 4
    Set Toolbar = Nothing
End Function


Public Function DB_SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                            ByVal wParam As Long, lParam As Long) As Long
    On Error Resume Next

    Dim Dockbar As CDockBar
    Dim ptrObject As Long

    ptrObject = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory Dockbar, ptrObject, 4

    DB_SubWndProc = Dockbar.WindowPROC(hWnd, Msg, _
    wParam, lParam)

    CopyMemory Dockbar, 0&, 4
    Set Dockbar = Nothing
End Function


Public Function MDF_SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                            ByVal wParam As Long, lParam As Long) As Long
    On Error Resume Next

    Dim miniDockFrame As CMiniDockFrameWnd
    Dim ptrObject As Long

    ptrObject = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory miniDockFrame, ptrObject, 4

    MDF_SubWndProc = miniDockFrame.WindowPROC(hWnd, Msg, _
    wParam, lParam)

    CopyMemory miniDockFrame, 0&, 4
    Set miniDockFrame = Nothing
End Function


Public Function FM_SubWndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                            ByVal wParam As Long, lParam As Long) As Long

    On Error Resume Next

    Dim Frame As CFrame
    Dim ptrObject As Long

    ptrObject = GetWindowLong(hWnd, GWL_USERDATA)
    CopyMemory Frame, ptrObject, 4

    FM_SubWndProc = Frame.WindowPROC(hWnd, Msg, _
    wParam, lParam)

    CopyMemory Frame, 0&, 4
    Set Frame = Nothing
End Function


Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, _
                        ByVal wParam As Long, lParam As Long) As Long
    WndProc = DefWindowProc(hWnd, Msg, wParam, lParam)
End Function


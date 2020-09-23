Attribute VB_Name = "ClassInits"
' ----------------------------------------------------------------- '
' Filename: ClassInits.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Code for custom window classes
' ----------------------------------------------------------------- '

Option Explicit

Public Const afxWndControlBar As String = "AfxControlBar42d"    ' AfxDockBar42d == AfxControlBar42d
Public Const afxMiniFrameWndClass As String = "AfxMiniFrameWnd"

Public Function Init_DockBarClass()
    Dim DockClass As WNDCLASS

    DockClass.style = CS_DBLCLKS
    DockClass.lpfnWndProc = GetProc(AddressOf WndProc)
    DockClass.cbClsExtra = 0
    DockClass.cbWndExtra2 = 0
    DockClass.hInstance = App.hInstance
    DockClass.hIcon = 0
    DockClass.hCursor = LoadCursor(0, IDC_ARROW)
    DockClass.hbrBackground = COLOR_BTNFACE + 1
    DockClass.lpszMenuName = ""
    DockClass.lpszClassName = afxWndControlBar

    Init_DockBarClass = RegisterClass(DockClass)
End Function

Public Function UnInit_DockBarClass()
    UnInit_DockBarClass = UnregisterClass(afxWndControlBar, App.hInstance)
End Function

Public Function Initialize()
    Call LoadLibrary("comctl32.dll")
    Call InitSysColors
    Call Init_DockBarClass
    Call Init_MiniFrameWndClass
End Function

Public Function UnInitialize()
    Call UnInit_DockBarClass
    Call UnInit_MiniFrameWndClass
End Function

Function GetProc(proc As Long) As Long
    GetProc = proc
End Function


Public Function Init_MiniFrameWndClass()
    Dim MiniFrameWndClass As WNDCLASS

    MiniFrameWndClass.style = CS_DBLCLKS
    MiniFrameWndClass.lpfnWndProc = GetProc(AddressOf WndProc) 'GetProc(AddressOf DockBarWinPROC)
    MiniFrameWndClass.cbClsExtra = 0
    MiniFrameWndClass.cbWndExtra2 = 0
    MiniFrameWndClass.hInstance = App.hInstance
    MiniFrameWndClass.hIcon = 0
    MiniFrameWndClass.hCursor = LoadCursor(0, IDC_ARROW)
    MiniFrameWndClass.hbrBackground = COLOR_BACKGROUND
    MiniFrameWndClass.lpszMenuName = ""
    MiniFrameWndClass.lpszClassName = afxMiniFrameWndClass

    Init_MiniFrameWndClass = RegisterClass(MiniFrameWndClass)
End Function

Public Function UnInit_MiniFrameWndClass()
    UnInit_MiniFrameWndClass = UnregisterClass(afxMiniFrameWndClass, App.hInstance)
End Function


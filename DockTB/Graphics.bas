Attribute VB_Name = "Graphics"
' ----------------------------------------------------------------- '
' Filename: Graphics.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Other graphics functions
' ----------------------------------------------------------------- '

Option Explicit

Dim m_dwStyle As Long
Global clrBtnFace As Long, clrBtnShadow As Long, clrBtnHilite As Long
Global clrBtnText As Long, clrWindowFrame As Long

Public Const ETO_OPAQUE = 2
Public Const ETO_GRAYED = 1
Public Const ETO_CLIPPED = 4


Public Function MyScreenToClient(hWnd As Long, ByRef r As RECT)
    Dim pt1 As POINTAPI, pt2 As POINTAPI
    pt1.x = r.Left
    pt1.y = r.Top

    pt2.x = r.Right
    pt2.y = r.Bottom

    Call ScreenToClient(hWnd, pt1)
    Call ScreenToClient(hWnd, pt2)

    r.Left = pt1.x
    r.Top = pt1.y
    r.Right = pt2.x
    r.Bottom = pt2.y
End Function


Public Function FillSolidRect(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal clr As Long)
    Call SetBkColor(hdc, clr)

    Dim recta As RECT

    recta.Left = x
    recta.Top = y
    recta.Right = x + cx
    recta.Bottom = y + cy

    Call ExtTextOut(hdc, 0, 0, ETO_OPAQUE, recta, 0, 0, 0)
End Function


Public Function Draw3dRect(hdc As Long, x As Long, y As Long, cx As Long, _
    cy As Long, clrTopLeft As Long, clrBottomRight As Long)

    Call FillSolidRect(hdc, x, y, cx - 1, 1, clrTopLeft)
    Call FillSolidRect(hdc, x, y, 1, cy - 1, clrTopLeft)
    Call FillSolidRect(hdc, x + cx, y, -1, cy, clrBottomRight)
    Call FillSolidRect(hdc, x, y + cy, cx, -1, clrBottomRight)
End Function


Public Function InitSysColors()
    clrBtnFace = GetSysColor(COLOR_BTNFACE)
    clrBtnShadow = GetSysColor(COLOR_BTNSHADOW)
    clrBtnHilite = GetSysColor(COLOR_BTNHIGHLIGHT)
    clrBtnText = GetSysColor(COLOR_BTNTEXT)
    clrWindowFrame = GetSysColor(COLOR_WINDOWFRAME)

End Function

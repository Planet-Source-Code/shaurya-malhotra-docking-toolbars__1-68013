Attribute VB_Name = "Tracking"
' ----------------------------------------------------------------- '
' Filename: Tracking.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Tracking (rubber-banding) functions
' ----------------------------------------------------------------- '

Option Explicit

Private afxHalftoneBrush  As Long

Public Function GetHalftoneBrush() As Long
    If (afxHalftoneBrush = 0) Then
        Dim grayPattern(7) As Integer

        Dim i As Integer
        For i = 0 To 7
            grayPattern(i) = CLng(21845 / (2 ^ (i And 1)))
        Next i

        Dim grayBitmap As Long
        grayBitmap = CreateBitmap(8, 8, 1, 1, grayPattern(0))

        If grayBitmap <> 0 Then
            afxHalftoneBrush = CreatePatternBrush(grayBitmap)
            DeleteObject (grayBitmap)
        End If
    End If
    GetHalftoneBrush = afxHalftoneBrush
End Function


Public Function DrawDragRect(lpRect As RECT, Size As Size, _
                        lpRectLast As RECT, sizeLast As Size, _
                        pBrush As Long, pBrushLast As Long, h_DC As Long)

    ' first, determine the update region and select it
    Dim rgnNew As Long
    Dim rgnOutside As Long, rgnInside As Long

    rgnOutside = CreateRectRgnIndirect(lpRect)

    Dim tmp_rect As RECT
    tmp_rect = lpRect

    Call InflateRect(tmp_rect, -Size.cx, -Size.cy)
    Call IntersectRect(tmp_rect, tmp_rect, lpRect)

    rgnInside = CreateRectRgnIndirect(tmp_rect)
    rgnNew = CreateRectRgn(0, 0, 0, 0)
    Call CombineRgn(rgnNew, rgnOutside, rgnInside, RGN_XOR)

    Dim pBrushOld As Long
    pBrushOld = 0

    If pBrush = 0 Then pBrush = GetHalftoneBrush()
    If pBrushLast = 0 Then pBrushLast = pBrush

    Dim rgnLast As Long, rgnUpdate As Long

    If ((lpRectLast.Top < lpRectLast.Bottom) And (lpRectLast.Left < lpRectLast.Right)) Then
        ' find difference between new region and old region
        rgnLast = CreateRectRgn(0, 0, 0, 0)
        Call SetRectRgn(rgnOutside, lpRectLast.Left, lpRectLast.Top, lpRectLast.Right, lpRectLast.Bottom) '*****'
        tmp_rect = lpRectLast

        Call InflateRect(tmp_rect, -sizeLast.cx, -sizeLast.cy)
        Call IntersectRect(tmp_rect, tmp_rect, lpRectLast)
        Call SetRectRgn(rgnInside, tmp_rect.Left, tmp_rect.Top, tmp_rect.Right, tmp_rect.Bottom)
        Call CombineRgn(rgnLast, rgnOutside, rgnInside, RGN_XOR)

        ' only diff them if brushes are the same
        If pBrush = pBrushLast Then
            rgnUpdate = CreateRectRgn(0, 0, 0, 0)
            Call CombineRgn(rgnUpdate, rgnLast, rgnNew, RGN_XOR)
        End If
    End If

    If ((pBrush <> pBrushLast) And ((lpRectLast.Top < lpRectLast.Bottom) And (lpRectLast.Left < lpRectLast.Right))) Then
        ' brushes are different -- erase old region first
        Call SelectClipRgn(h_DC, rgnLast)
        Call GetClipBox(h_DC, tmp_rect)
        pBrushOld = SelectObject(h_DC, pBrushLast)

        Call PatBlt(h_DC, tmp_rect.Left, tmp_rect.Top, (tmp_rect.Right - tmp_rect.Left), (tmp_rect.Bottom - tmp_rect.Top), PATINVERT)

        Call SelectObject(h_DC, pBrushOld)
        pBrushOld = 0
    End If

    ' draw into the update/new region
    If rgnUpdate <> 0 Then
        Call SelectClipRgn(h_DC, rgnUpdate)
    Else
        Call SelectClipRgn(h_DC, rgnNew)
    End If

    Call GetClipBox(h_DC, tmp_rect)
    pBrushOld = SelectObject(h_DC, pBrush)
    Call PatBlt(h_DC, tmp_rect.Left, tmp_rect.Top, (tmp_rect.Right - tmp_rect.Left), (tmp_rect.Bottom - tmp_rect.Top), PATINVERT)

    ' cleanup DC
    If (pBrushOld <> 0) Then Call SelectObject(h_DC, pBrushOld)

    Call SelectClipRgn(h_DC, 0)


    '-------- Free Resources of DrawDragRect --------
    If rgnNew <> 0 Then DeleteObject (rgnNew)
    If rgnOutside <> 0 Then DeleteObject (rgnOutside)
    If rgnInside <> 0 Then DeleteObject (rgnInside)
    If pBrushOld <> 0 Then DeleteObject (pBrushOld)
    If rgnLast <> 0 Then DeleteObject (rgnLast)
    If rgnUpdate <> 0 Then DeleteObject (rgnUpdate)
    '------------------------------------------------
End Function



Attribute VB_Name = "FloatingFramesList"
' ----------------------------------------------------------------- '
' Filename: FloatingFramesList.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Keeps track of floating frames
' ----------------------------------------------------------------- '

Option Explicit

Private floatBars As New CCollection


Public Function AddBar(pBar As CMiniDockFrameWnd)
    floatBars.Insert pBar
End Function


Public Function GetParentBar(hWnd As Long) As CMiniDockFrameWnd
    Dim p_hWnd As Long  'parent's hWnd
    p_hWnd = GetParent(hWnd)
    
    Dim i As Long
    i = 1
    Dim o As Object
    
    Do While (i <= floatBars.Size)
    Set o = floatBars.item(i)
        If Not (floatBars.item(i) Is Nothing) Then
            If floatBars.item(i).m_hWnd = p_hWnd Then
                Set GetParentBar = floatBars.item(i)
                Exit Function
            End If
        End If
    i = i + 1
    Loop
    Set GetParentBar = Nothing  'not found
End Function




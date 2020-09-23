Attribute VB_Name = "AllWindowsList"
' ----------------------------------------------------------------- '
' Filename: AllWindowsList.bas
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Keeps track of all the windows in the application
' ----------------------------------------------------------------- '

Option Explicit

Private cAllWindows As New CCollection


Public Function WindowExistsInList(hWnd As Long) As Object
    Dim i As Long
    i = cAllWindows.Size

    Do While (i > 0)
        If cAllWindows.item(i).hWnd = hWnd Then
            Set WindowExistsInList = cAllWindows.item(i)
            Exit Function
        End If
        i = i - 1
    Loop
Set WindowExistsInList = Nothing
End Function


Public Function AddWindowInList(obj As Object)
    Call cAllWindows.Insert(obj)
End Function


Public Function RemoveWindowInList(obj As Object)
    Dim i As Long
    i = cAllWindows.Size
    
    Do While (i > 0)
        If (obj Is cAllWindows.item(i)) Then
            cAllWindows.Remove (i)
            Exit Function
        End If
        i = i - 1
    Loop
End Function

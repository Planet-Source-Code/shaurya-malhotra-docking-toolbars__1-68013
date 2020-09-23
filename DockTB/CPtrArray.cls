VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPtrArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CPtrArray.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' A class built on the Visual Basic's Collection class, which
' stores/retrieves variant types at specified index locations
' (used here as an array of pointers)
' ----------------------------------------------------------------- '

Option Explicit

Private items As New Collection


Public Sub Add(item As Variant)
    items.Add item
End Sub

Public Sub InsertAt(ByVal pos As Long, item As Variant)
    pos = pos + 1

    If pos <= items.Count Then
        items.Add item, , pos
    Else
        Dim i As Integer
        For i = 1 To (pos - items.Count) - 1
            items.Add Nothing
        Next i
        items.Add item
    End If
End Sub

Public Sub RemoveAt(ByVal pos As Long)
    pos = pos + 1
    items.Remove (pos)
End Sub


Public Property Get getCollection() As Collection
    Set getCollection = items
End Property

Public Function GetItem(ByVal pos As Integer) As Variant
On Error GoTo ErrHandler
    pos = pos + 1
    Set GetItem = IIf(IsObject(items(pos)), items(pos), Nothing)
    Exit Function
ErrHandler:
    Debug.Print "pos = " & pos & vbTab & "count = " & items.Count
    Stop
End Function


Public Property Get GetSize() As Long
    GetSize = items.Count
End Property

Public Sub SetAtPos(ByVal pos As Integer, ob As Variant)
    pos = pos + 1
    Call items.Add(ob, , pos)
    Call items.Remove(pos + 1)
End Sub


Public Function GetItemID(ByVal pos As Integer) As Long
    pos = pos + 1
    If IsObject(items(pos)) Then
        If Not (items(pos) Is Nothing) Then
            GetItemID = items(pos).m_hWnd
        Else
            GetItemID = -100
        End If
    Else
        GetItemID = items(pos)
    End If
End Function


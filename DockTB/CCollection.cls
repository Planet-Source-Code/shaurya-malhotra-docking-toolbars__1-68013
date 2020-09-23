VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ----------------------------------------------------------------- '
' Filename: CCollection.cls
' Author:   Shaurya Malhotra (shauryamal@gmail.com)
' Date:     24 February 2007
'
' Wrapper for the Visual Basic's Collection class
' ----------------------------------------------------------------- '

Option Explicit

Private items As New Collection

Public Sub Insert(item As Object)
    items.Add item
End Sub

Public Sub Remove(pos As Long)
    items.Remove (pos)
End Sub

Public Property Get item(pos As Long) As Object
    Set item = items(pos)
End Property

Public Property Get Size() As Long
    Size = items.Count
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

Implements IList

Private Type Data
  Items As Collection
End Type

Private This As Data

Private Sub Class_Initialize()
  Set This.Items = New Collection
End Sub

Public Sub Add(ByRef Item As Variant)
  This.Items.Add Item
End Sub

Public Function Build() As IList
  Set Build = Me
End Function

Private Property Get IList_Count() As Long
  IList_Count = This.Items.Count
End Property

Private Property Get IList_Item(ByVal Index As Long) As Variant
  LetSet(IList_Item) = This.Items(Index)
End Property

Private Property Get IList_NewEnum() As stdole.IUnknown
  Set IList_NewEnum = This.Items.[_NewEnum]
End Property

Private Sub Class_Terminate()
  Set This.Items = Nothing
End Sub

Private Property Get IList_ToString() As String
  Dim Builder As String
  With StrJoin(", ")
    Dim I As Long: For I = 1 To This.Items.Count
      Builder = .Apply(Builder)(This.Items(I)).Run()
    Next I
  End With
  With StrJoin(" ")
    IList_ToString = .Apply(.Apply("[")(Builder).Run())("]").Run
  End With
End Property

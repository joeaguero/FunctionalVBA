Attribute VB_Name = "Data"
'@Folder "Library"
Option Explicit

Public Function Pair(ByRef Item1 As Variant, ByRef Item2 As Variant) As IPair
  With New Pair
    Set Pair = .Init(Item1, Item2)
  End With
End Function

Public Function List(ParamArray Items() As Variant) As IList
  Dim I As Long
  
  With New List
    For I = LBound(Items) To UBound(Items)
      .Add Items(I)
    Next I
    Set List = .Build
  End With
End Function


Attribute VB_Name = "Data"
'@Folder "Library"
Option Explicit

Public Function Pair(ByRef Item1 As Variant, ByRef Item2 As Variant) As IPair
  With New Pair
    Set Pair = .Init(Item1, Item2)
  End With
End Function

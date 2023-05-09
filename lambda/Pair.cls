VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Concrete"
Option Explicit
Implements IPair

Private Type Data
  Item1 As Variant
  Item2 As Variant
End Type

Private This As Data

Public Function Init(ByRef Item1 As Variant, ByRef Item2 As Variant) As IPair
  LetSet(This.Item1) = Item1
  LetSet(This.Item2) = Item2
  Set Init = Me
End Function

Private Property Get IPair_Item1() As Variant
  LetSet(IPair_Item1) = This.Item1
End Property

Private Property Get IPair_Item2() As Variant
  LetSet(IPair_Item2) = This.Item2
End Property

Private Function IPair_Map(ByRef Func As Variant) As IPair
  Set IPair_Map = Pair.Of(Func.Of(This.Item1))(Func.Of(This.Item2))
End Function

Private Property Get IPair_Show() As String
  IPair_Show = "( " & Show(This.Item1) & ", " & Show(This.Item2) & " )"
End Property

Private Sub Class_Terminate()
  Set This.Item1 = Nothing
  Set This.Item2 = Nothing
End Sub

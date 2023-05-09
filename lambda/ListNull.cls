VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Concrete"
Option Explicit

Implements IList

Public Function Init() As IList
  Set Init = Me
End Function

Private Property Get IList_Head() As Variant
  Err.Raise 6969, "ListNull", "Empty List"
End Property

Private Property Get IList_Tail() As IList
  Set IList_Tail = Me
End Property

Private Property Get IList_IsEmpty() As Boolean
  IList_IsEmpty = True
End Property

Private Property Get IList_Show() As String
  IList_Show = "[]"
End Property

Private Function IList_Map(ByRef Func As Variant) As IList
  Set IList_Map = ListNull()
End Function

Private Function IList_Bind(ByRef Func As Variant) As IList
  Set IList_Bind = ListNull()
End Function

Private Function IList_Reduce(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IList_Reduce) = Default
End Function

Private Function IList_ReduceL(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IList_ReduceL) = Default
End Function

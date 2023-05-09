VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Just"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Concrete"
Option Explicit

Implements IMaybe

Private Type Data
  Item As Variant
End Type

Private This As Data

Public Function Init(ByRef Item As Variant) As IMaybe
  LetSet(This.Item) = Item
  Set Init = Me
End Function

Private Function IMaybe_Bind(ByRef Func As Variant) As IMaybe
  Set IMaybe_Bind = Func.Of(This.Item)
End Function

Private Function IMaybe_Map(ByRef Func As Variant) As IMaybe
  Set IMaybe_Map = Just.Of(Func.Of(This.Item))
End Function

Private Function IMaybe_Reduce(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IMaybe_Reduce) = Func.Of(This.Item).Of(Default)
End Function

Private Function IMaybe_ReduceL(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IMaybe_ReduceL) = Func.Of(This.Item).Of(Default)
End Function

Private Property Get IMaybe_Show() As String
  IMaybe_Show = "Just " & Show(This.Item)
End Property

Private Property Get IMaybe_IsEmpty() As Boolean
  IMaybe_IsEmpty = False
End Property

Private Sub Class_Terminate()
  Set This.Item = Nothing
End Sub

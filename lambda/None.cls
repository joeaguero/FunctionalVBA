VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "None"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Concrete"
Option Explicit

Implements IMaybe

Public Function Init() As IMaybe
  Set Init = Me
End Function

Private Function IMaybe_Bind(ByRef Func As Variant) As IMaybe
  Set IMaybe_Bind = Me
End Function

Private Function IMaybe_Map(ByRef Func As Variant) As IMaybe
  Set IMaybe_Map = Me
End Function

Private Function IMaybe_Reduce(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IMaybe_Reduce) = Default
End Function

Private Function IMaybe_ReduceL(ByRef Func As ILambda2, ByRef Default As Variant) As Variant
  LetSet(IMaybe_ReduceL) = Default
End Function

Private Property Get IMaybe_IsEmpty() As Boolean
  IMaybe_IsEmpty = True
End Property

Private Property Get IMaybe_Show() As String
  IMaybe_Show = "Null"
End Property


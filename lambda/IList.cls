VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Immutable Collection Wrapper"
'@Folder "Library.Data.Abstraction"
Option Explicit

Public Property Get Head() As Variant: End Property
Public Property Get Tail() As IList: End Property
Public Property Get IsEmpty() As Boolean: End Property
Public Property Get Show() As String
Attribute Show.VB_UserMemId = 0
End Property
Public Function Map(ByRef Func As Variant) As IList: End Function
Public Function Bind(ByRef Func As Variant) As IList: End Function
Public Function Reduce(ByRef Func As ILambda2, ByRef Default As Variant) As Variant: End Function
Public Function ReduceL(ByRef Func As ILambda2, ByRef Default As Variant) As Variant: End Function

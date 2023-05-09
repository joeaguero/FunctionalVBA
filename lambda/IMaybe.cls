VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMaybe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Abstraction"
Option Explicit

Public Function Bind(ByRef Func As Variant) As IMaybe: End Function
Public Function Map(ByRef Func As Variant) As IMaybe: End Function
Public Function Reduce(ByRef Func As ILambda2, ByRef Default As Variant) As Variant: End Function
Public Function ReduceL(ByRef Func As ILambda2, ByRef Default As Variant) As Variant: End Function
Public Property Get Show() As String
Attribute Show.VB_UserMemId = 0
End Property
Public Property Get IsEmpty() As Boolean: End Property

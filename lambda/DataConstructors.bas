Attribute VB_Name = "DataConstructors"
'@Folder "Library.Data"
Option Explicit
Option Private Module

Public Function List(ParamArray Items() As Variant) As IList
  Set List = ToList.Of(CVar(Items))
End Function

Public Function ListNull() As IList
  With New ListNull
    Set ListNull = .Init()
  End With
End Function

Public Function None() As IMaybe
  With New None
    Set None = .Init()
  End With
End Function

Public Function EmptyArray() As Variant
  Dim Arr() As Byte
  ReDim Arr(0)
  EmptyArray = Arr
End Function

' VBA Utilties
Public Function Str(ByVal Text As String) As IList
  Dim Bytes() As Byte
  Bytes = StrConv(Text, vbFromUnicode)
  Set Str = ToList.Of(Bytes)
End Function

Public Function Decode(ByRef Text As IList) As String
  Decode = StrConv(ToArray.Of(Text), vbUnicode)
End Function

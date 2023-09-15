Attribute VB_Name = "Misc"
'@Folder "Lambda.Primitives"
Option Explicit
Option Private Module

Public Function PadLeft(Char As String, Size As Long, Text As Variant)
  Dim OutText As String: OutText = CStr(Text)
  Dim PadLength As Long: PadLength = Size - Len(OutText)
  If PadLength < 0 Then
    PadLeft = OutText
  Else
    PadLeft = String(PadLength, Char) & OutText
  End If
End Function

Public Function PadRight(Char As String, Size As Long, Text As Variant)
  Dim OutText As String: OutText = CStr(Text)
  Dim PadLength As Long: PadLength = Size - Len(OutText)
  If PadLength < 0 Then
    PadRight = OutText
  Else
    PadRight = OutText & String(PadLength, Char)
  End If
End Function

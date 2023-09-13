Attribute VB_Name = "FnUtils"
'@Folder "Lambda.Primitives"
Option Explicit

Public Property Let LetSet(ByRef LHS As Variant, RHS As Variant)
    If IsObject(RHS) Then
        Set LHS = RHS
    Else
        LHS = RHS
    End If
End Property

Public Function Fn(Target As String) As Lambda
  With New CLambda
    Set Fn = .Init(Target)
  End With
End Function

' Misc Utils
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

Attribute VB_Name = "LibraryOperators"
'@Folder("Library.Definitions")
Option Explicit
Option Private Module

Private Function Addition(ByVal LHS As Double, ByVal RHS As Double) As Double
  Addition = LHS + RHS
End Function

Private Function Constant(ByRef Source As Variant, ByRef Ignored As Variant) As Variant
  LetSet(Constant) = Source
End Function

Private Function JoinWith(ByRef Sep As IList, ByRef LHS As IList, ByRef RHS As IList) As IList
  If IsEmpty(LHS) Then
    Set JoinWith = RHS
  ElseIf IsEmpty(RHS) Then
    Set JoinWith = LHS
  Else
    Set JoinWith = Append(Append(LHS)(Sep))(RHS)
  End If
End Function

' TODO: Implement Reflection through objects
Private Function EQ(ByRef LHS As Variant, ByRef RHS As Variant) As Boolean
  EQ = RHS = LHS
End Function

Private Function GT(ByRef LHS As Variant, ByRef RHS As Variant) As Boolean
  GT = RHS > LHS
End Function

Private Function LT(ByRef LHS As Variant, ByRef RHS As Variant) As Boolean
  LT = RHS < LHS
End Function

Private Function NotBool(ByVal Bool As Boolean) As Boolean
  NotBool = Not Bool
End Function


' VBA Utilities
Private Function TextJoin(ByVal Sep As String, ByRef LHS As Variant, ByRef RHS As Variant) As String
  LHS = Show(LHS)
  RHS = Show(RHS)
  
  If RHS = "" Then
    TextJoin = LHS
  ElseIf LHS = "" Then
    TextJoin = RHS
  Else
    TextJoin = LHS & Sep & RHS
  End If
End Function

Private Function AppendArray(ByRef Value As Byte, ByRef Arr() As Byte) As Variant
  If Arr(LBound(Arr)) <> Empty Then ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
  Arr(UBound(Arr)) = Value
  AppendArray = Arr
End Function

Private Function MaxByte(ByVal LHS As Byte, ByVal RHS As Byte) As Byte
  If LHS < RHS Then MaxByte = RHS Else MaxByte = LHS
End Function

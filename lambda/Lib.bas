Attribute VB_Name = "Lib"
'@Folder "Library"
Option Explicit

Public Property Get Zip() As ILambda2
  Set Zip = ZipWith(EnPair)
End Property

Public Property Get Addition() As ILambda2
  Set Addition = Lambda2("AdditionFunc")
End Property

Private Function AdditionFunc(ByVal LHS As Double, ByVal RHS As Double) As Double
  AdditionFunc = LHS + RHS
End Function

Public Property Get EnPair() As ILambda2
  Set EnPair = Lambda2("EnPairFunc")
End Property

Private Function EnPairFunc(ByRef LHS As Variant, ByRef RHS As Variant) As IPair
  Set EnPairFunc = Pair(LHS, RHS)
End Function

Public Property Get ZipWith() As ILambda3
  Set ZipWith = Lambda3("ZipWithFunc")
End Property

Private Function ZipWithFunc(ByRef BinaryOperator As ILambda2, ByRef LHS As IList, ByRef RHS As IList) As IList
  Dim I As Long, Length As Long
  
  If LHS.Count < RHS.Count Then Length = LHS.Count Else Length = RHS.Count
  
  With New List
    For I = 1 To IIf(LHS.Count < RHS.Count, LHS.Count, RHS.Count)
      .Add BinaryOperator(LHS.Item(I))(RHS.Item(I)).Run()
    Next I
    Set ZipWithFunc = .Build
  End With
End Function

Public Property Get StrJoin() As ILambda3
  Set StrJoin = Lambda3("StrJoinFunc")
End Property

Private Function StrJoinFunc(ByVal Sep As String, ByVal LHS As String, ByVal RHS As String) As String
  If LHS = "" Then StrJoinFunc = RHS: Exit Function
  If RHS = "" Then StrJoinFunc = LHS: Exit Function
  StrJoinFunc = LHS & Sep & RHS
End Function

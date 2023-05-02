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

Private Function ZipWithFunc(ByRef BinaryOperator As ILambda2, ByRef LHS As Collection, ByRef RHS As Collection) As Collection
  Dim I As Long, Length As Long
  
  If LHS.Count < RHS.Count Then Length = LHS.Count Else Length = RHS.Count
  
  Set ZipWithFunc = New Collection
  
  For I = 1 To IIf(LHS.Count < RHS.Count, LHS.Count, RHS.Count)
    ZipWithFunc.Add BinaryOperator(LHS(I))(RHS(I)).Run()
  Next I
End Function

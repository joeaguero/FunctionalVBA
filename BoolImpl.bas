Attribute VB_Name = "BoolImpl"
'@Folder "Lambda.Primitives"
Option Explicit
Option Private Module

Private Function Yes(YesVal As Variant, NoVal As Variant) As Variant
  LetSet(Yes) = YesVal
End Function

Private Function No(YesVal As Variant, NoVal As Variant) As Variant
  LetSet(No) = NoVal
End Function

Private Function Both(LHS As Lambda02, RHS As Lambda02) As Lambda02
  Set Both = LHS(RHS)(Bool.No).Run()
End Function

Private Function Either(LHS As Lambda02, RHS As Lambda02) As Lambda02
  Set Either = LHS(Bool.Yes)(RHS).Run()
End Function

Private Function Invert(LHS As Lambda02) As Lambda02
  Set Invert = LHS(Bool.No)(Bool.Yes).Run()
End Function

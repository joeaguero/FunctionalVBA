Attribute VB_Name = "Bool"
'@Folder "Lambda.Primitives"
Option Explicit

Public Function Show(Bool As Lambda02) As Boolean
  Show = Bool(True)(False).Run()
End Function

Public Function Read(Val As Boolean) As Lambda02
  If Val Then
    Set Read = Bool.Yes
  Else
    Set Read = Bool.No
  End If
End Function

Public Function Yes() As Lambda02: Set Yes = Fn02("BoolImpl.Yes"): End Function
Public Function No() As Lambda02: Set No = Fn02("BoolImpl.No"): End Function
Public Function Both() As Lambda02: Set Both = Fn02("BoolImpl.Both"): End Function
Public Function Either() As Lambda02: Set Either = Fn02("BoolImpl.Either"): End Function
Public Function Invert() As Lambda01: Set Invert = Fn01("BoolImpl.Invert"): End Function



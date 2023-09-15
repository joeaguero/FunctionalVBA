Attribute VB_Name = "Lambda"
'@Folder "Lambda.Primitives"
Option Explicit
Option Private Module

Public Property Let LetSet(ByRef LHS As Variant, RHS As Variant)
    If IsObject(RHS) Then
        Set LHS = RHS
    Else
        LHS = RHS
    End If
End Property

Public Function Fn10(Target As String) As Lambda10
  With New Lambda10Impl
    Set Fn10 = .Init(Target)
  End With
End Function

Public Function Fn09(Target As String) As Lambda09
  With New Lambda09Impl
    Set Fn09 = .Init(Target)
  End With
End Function

Public Function Fn08(Target As String) As Lambda08
  With New Lambda08Impl
    Set Fn08 = .Init(Target)
  End With
End Function

Public Function Fn07(Target As String) As Lambda07
  With New Lambda07Impl
    Set Fn07 = .Init(Target)
  End With
End Function

Public Function Fn06(Target As String) As Lambda06
  With New Lambda06Impl
    Set Fn06 = .Init(Target)
  End With
End Function

Public Function Fn05(Target As String) As Lambda05
  With New Lambda05Impl
    Set Fn05 = .Init(Target)
  End With
End Function

Public Function Fn04(Target As String) As Lambda04
  With New Lambda04Impl
    Set Fn04 = .Init(Target)
  End With
End Function

Public Function Fn03(Target As String) As Lambda03
  With New Lambda03Impl
    Set Fn03 = .Init(Target)
  End With
End Function

Public Function Fn02(Target As String) As Lambda02
  With New Lambda02Impl
    Set Fn02 = .Init(Target)
  End With
End Function

Public Function Fn01(Target As String) As Lambda01
  With New Lambda01Impl
    Set Fn01 = .Init(Target)
  End With
End Function

Public Function Fn00(Target As String) As Closure
  With New ClosureImpl
    Set Fn00 = .Init(Target)
  End With
End Function
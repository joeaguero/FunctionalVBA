VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Functional.Concrete")
Option Explicit

Implements ILambda4
Implements ILambda3
Implements ILambda2
Implements ILambda1

Private Type Data
  Target As String
  Arg1 As Variant
  Arg2 As Variant
  Arg3 As Variant
End Type

Private This As Data

Public Function Init4(ByVal Target As String) As ILambda4
  This.Target = Target
  Set Init4 = Me
End Function

Public Function Init3(ByVal Target As String, ByRef Arg1 As Variant) As ILambda3
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  Set Init3 = Me
End Function

Public Function Init2(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant) As ILambda2
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  Set Init2 = Me
End Function

Public Function Init1(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant, ByRef Arg3 As Variant) As ILambda1
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  LetSet(This.Arg3) = Arg3
  Set Init1 = Me
End Function

Private Function ILambda4_Of(ByRef Arg As Variant) As ILambda3
  With New Lambda4
    Set ILambda4_Of = .Init3(This.Target, Arg)
  End With
End Function

Private Function ILambda3_Of(ByRef Arg As Variant) As ILambda2
  With New Lambda4
    Set ILambda3_Of = .Init2(This.Target, This.Arg1, Arg)
  End With
End Function

Private Function ILambda2_Of(ByRef Arg As Variant) As ILambda1
  With New Lambda4
    Set ILambda2_Of = .Init1(This.Target, This.Arg1, This.Arg2, Arg)
  End With
End Function

Private Function ILambda1_Of(ByRef Arg As Variant) As Variant
  LetSet(ILambda1_Of) = Application.Run(This.Target, This.Arg1, This.Arg2, This.Arg3, Arg)
End Function

Private Sub Class_Terminate()
  Set This.Arg1 = Nothing
  Set This.Arg2 = Nothing
  Set This.Arg3 = Nothing
End Sub

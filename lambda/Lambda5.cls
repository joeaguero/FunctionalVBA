VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Functional.Concrete")
Option Explicit

Implements ILambda5
Implements ILambda4
Implements ILambda3
Implements ILambda2
Implements ILambda1
Implements INullary

Private Type Data
  Target As String
  Arg1 As Variant
  Arg2 As Variant
  Arg3 As Variant
  Arg4 As Variant
  Arg5 As Variant
End Type

Private This As Data

Public Function Init5(ByVal Target As String) As ILambda5
  This.Target = Target
  Set Init5 = Me
End Function

Public Function Init4(ByVal Target As String, ByRef Arg1 As Variant) As ILambda4
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  Set Init4 = Me
End Function

Public Function Init3(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant) As ILambda3
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  Set Init3 = Me
End Function

Public Function Init2(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant, ByRef Arg3 As Variant) As ILambda2
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  LetSet(This.Arg3) = Arg3
  Set Init2 = Me
End Function

Public Function Init1(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant, ByRef Arg3 As Variant, ByRef Arg4 As Variant) As ILambda1
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  LetSet(This.Arg3) = Arg3
  LetSet(This.Arg4) = Arg4
  Set Init1 = Me
End Function

Public Function Init0(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant, ByRef Arg3 As Variant, ByRef Arg4 As Variant, ByRef Arg5 As Variant) As INullary
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  LetSet(This.Arg3) = Arg3
  LetSet(This.Arg4) = Arg4
  LetSet(This.Arg5) = Arg5
  Set Init0 = Me
End Function

Public Function ILambda5_Apply(ByRef Arg As Variant) As ILambda4
  With New Lambda5
    Set ILambda5_Apply = .Init4(This.Target, Arg)
  End With
End Function

Public Function ILambda4_Apply(ByRef Arg As Variant) As ILambda3
  With New Lambda5
    Set ILambda4_Apply = .Init3(This.Target, This.Arg1, Arg)
  End With
End Function

Public Function ILambda3_Apply(ByRef Arg As Variant) As ILambda2
  With New Lambda5
    Set ILambda3_Apply = .Init2(This.Target, This.Arg1, This.Arg2, Arg)
  End With
End Function

Public Function ILambda2_Apply(ByRef Arg As Variant) As ILambda1
  With New Lambda5
    Set ILambda2_Apply = .Init1(This.Target, This.Arg1, This.Arg2, This.Arg3, Arg)
  End With
End Function

Public Function ILambda1_Apply(ByRef Arg As Variant) As INullary
  With New Lambda5
    Set ILambda1_Apply = .Init0(This.Target, This.Arg1, This.Arg2, This.Arg3, This.Arg4, Arg)
  End With
End Function

Private Function INullary_Run() As Variant
  LetSet(INullary_Run) = Application.Run(This.Target, This.Arg1, This.Arg2, This.Arg3, This.Arg4, This.Arg5)
End Function

Private Sub Class_Terminate()
  Set This.Arg1 = Nothing
  Set This.Arg2 = Nothing
  Set This.Arg3 = Nothing
  Set This.Arg4 = Nothing
  Set This.Arg5 = Nothing
End Sub

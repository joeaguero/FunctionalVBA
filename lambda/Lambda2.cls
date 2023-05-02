VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Functional.Concrete")
Option Explicit

Implements ILambda2
Implements ILambda1
Implements INullary

Private Type Data
  Target As String
  Arg1 As Variant
  Arg2 As Variant
End Type

Private This As Data

Public Function Init2(ByVal Target As String) As ILambda2
  This.Target = Target
  Set Init2 = Me
End Function

Public Function Init1(ByVal Target As String, ByRef Arg1 As Variant) As ILambda1
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  Set Init1 = Me
End Function

Public Function Init0(ByVal Target As String, ByRef Arg1 As Variant, ByRef Arg2 As Variant) As INullary
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  LetSet(This.Arg2) = Arg2
  Set Init0 = Me
End Function

Public Function ILambda2_Apply(ByRef Arg As Variant) As ILambda1
  With New Lambda2
    Set ILambda2_Apply = .Init1(This.Target, Arg)
  End With
End Function

Public Function ILambda1_Apply(ByRef Arg As Variant) As INullary
  With New Lambda2
    Set ILambda1_Apply = .Init0(This.Target, This.Arg1, Arg)
  End With
End Function

Private Function INullary_Run() As Variant
  LetSet(INullary_Run) = Application.Run(This.Target, This.Arg1, This.Arg2)
End Function

Private Sub Class_Terminate()
  Set This.Arg1 = Nothing
  Set This.Arg2 = Nothing
End Sub

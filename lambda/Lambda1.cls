VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Functional.Concrete")
Option Explicit

Implements ILambda1
Implements INullary

Private Type Data
  Target As String
  Arg1 As Variant
End Type

Private This As Data

Public Function Init1(ByVal Target As String) As ILambda1
  This.Target = Target
  Set Init1 = Me
End Function

Public Function Init0(ByVal Target As String, ByRef Arg1 As Variant) As INullary
  This.Target = Target
  LetSet(This.Arg1) = Arg1
  Set Init0 = Me
End Function

Public Function ILambda1_Apply(ByRef Arg As Variant) As INullary
  With New Lambda1
    Set ILambda1_Apply = .Init0(This.Target, Arg)
  End With
End Function

Private Function INullary_Run() As Variant
  LetSet(INullary_Run) = Application.Run(This.Target, This.Arg1)
End Function


Private Sub Class_Terminate()
  This.Arg1 = Empty
End Sub

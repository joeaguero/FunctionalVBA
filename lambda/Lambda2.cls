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
  Target As String
End Type

Private This As Data
  This.Target = Target
  Set Init2 = Me
End Function
  This.Target = Target
  Set Init1 = Me
End Function
  This.Target = Target
  Set Init0 = Me
End Function
  With New Lambda2
    Set ILambda2_Apply = .Init1(This.Target, Arg)
  End With
End Function

  With New Lambda2
    Set ILambda1_Apply = .Init0(This.Target, This.Arg1, Arg)
  End With
End Function

  LetSet(INullary_Run) = Application.Run(This.Target, This.Arg1, This.Arg2)
End Function

End Sub
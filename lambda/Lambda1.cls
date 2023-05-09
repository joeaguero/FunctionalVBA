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

Private Type Data
  Target As String
End Type

Private This As Data

Public Function Init1(ByVal Target As String) As ILambda1
  This.Target = Target
  Set Init1 = Me
End Function

Private Function ILambda1_Of(ByRef Arg As Variant) As Variant
  LetSet(ILambda1_Of) = Application.Run(This.Target, Arg)
End Function

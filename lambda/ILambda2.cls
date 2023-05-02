VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILambda2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ILambda1 Factory"
'@Folder("Functional.Abstractions")
Option Explicit

Public Function Apply(ByRef Arg As Variant) As ILambda1
Attribute Apply.VB_Description = "ILambda1 Factory Method"
Attribute Apply.VB_UserMemId = 0
End Function

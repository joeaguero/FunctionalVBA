VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILambda4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "ILambda3 Factory"
'@Folder("Functional.Abstractions")
Option Explicit

Public Function Apply(ByRef Arg As Variant) As ILambda3
  Attribute Apply.VB_Description = "ILambda3 Factory Method"
  Attribute Apply.VB_UserMemID = 0
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "{current_name}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "{next_name} Factory"
'@Folder("Functional.Abstractions")
Option Explicit

Public Function Apply(ByRef Arg As Variant) As {next_name}
  Attribute Apply.VB_Description = "{next_name} Factory Method"
  Attribute Apply.VB_UserMemID = 0
End Function

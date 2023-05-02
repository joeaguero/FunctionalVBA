VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Nullary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Functional.Concrete")
Option ExplicitImplements INullaryPrivate Type Data
  Target As String
End Type

Private This As DataPublic Function Init0(ByVal Target As String) As INullary
  This.Target = Target
  Set Init0 = Me
End FunctionPrivate Function INullary_Run() As Variant
  LetSet(INullary_Run) = Application.Run(This.Target)
End Function

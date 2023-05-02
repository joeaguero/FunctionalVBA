Attribute VB_Name = "FunctionalConstructors"
'@Folder("Functional")
Option Private Module
Option Explicit

Public Property Let LetSet(ByRef Target As Variant, ByRef Source As Variant)
  If IsObject(Source) Then
    Set Target = Source
  Else
    Target = Source
  End If
End Property

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

Public Property Get Nullary(ByVal Target As String) As INullary
Attribute Nullary.VB_Description = "Builds a lambda function out of the target reference"
  With New Nullary
    Set Nullary = .Init0(Target)
  End With
End Property

Public Property Get Lambda1(ByVal Target As String) As ILambda1
Attribute Lambda1.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda1
    Set Lambda1 = .Init1(Target)
  End With
End Property

Public Property Get Lambda2(ByVal Target As String) As ILambda2
Attribute Lambda2.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda2
    Set Lambda2 = .Init2(Target)
  End With
End Property

Public Property Get Lambda3(ByVal Target As String) As ILambda3
Attribute Lambda3.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda3
    Set Lambda3 = .Init3(Target)
  End With
End Property

Public Property Get Lambda4(ByVal Target As String) As ILambda4
Attribute Lambda4.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda4
    Set Lambda4 = .Init4(Target)
  End With
End Property

Public Property Get Lambda5(ByVal Target As String) As ILambda5
Attribute Lambda5.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda5
    Set Lambda5 = .Init5(Target)
  End With
End Property

Public Property Get Lambda6(ByVal Target As String) As ILambda6
Attribute Lambda6.VB_Description = "Builds a lambda function out of the target reference"
  With New Lambda6
    Set Lambda6 = .Init6(Target)
  End With
End Property

Public Property Get {name}(ByVal Target As String) As {abstract_name}
  Attribute {name}.VB_Description = "Builds a lambda function out of the target reference"
  With New {name}
    Set {name} = .Init{level}(Target)
  End With
End Property

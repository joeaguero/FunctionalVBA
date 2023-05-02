Public Function {abstract_name}_Apply(ByRef Arg As Variant) As {next_name}
  With New {parent_name}
    Set {abstract_name}_Apply = .Init{next_level}(This.Target{members}, Arg)
  End With
End Function

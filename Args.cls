VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Args"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Primitives"
Option Explicit
Implements IArgs

Private Type Data
  Arg As Variant
  Prev As IArgs
  Size As Long
End Type

Private This As Data

Public Function Init(Arg As Variant, Optional Prev As IArgs = Nothing) As IArgs
  LetSet(This.Arg) = Arg
  Set This.Prev = Prev
  If Prev Is Nothing Then This.Size = 1 Else This.Size = Prev.Size + 1
  Set Init = Me
End Function

Private Property Get IArgs_Arg() As Variant
  LetSet(IArgs_Arg) = This.Arg
End Property

Private Property Get IArgs_Prev() As IArgs
  Set IArgs_Prev = This.Prev
End Property

Private Property Get IArgs_Size() As Long
  IArgs_Size = This.Size
End Property

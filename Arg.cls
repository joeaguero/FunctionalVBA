VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Primitives"
Option Explicit
Implements Args

Private Type Data
  Arg As Variant
  Prev As Args
  Size As Long
End Type

Private This As Data

Public Function Init(Arg As Variant, Optional Prev As Args = Nothing) As Args
  LetSet(This.Arg) = Arg
  Set This.Prev = Prev
  If Prev Is Nothing Then This.Size = 1 Else This.Size = Prev.Size + 1
  Set Init = Me
End Function

Private Property Get Args_Arg() As Variant
  LetSet(Args_Arg) = This.Arg
End Property

Private Property Get Args_Prev() As Args
  Set Args_Prev = This.Prev
End Property

Private Property Get Args_Size() As Long
  Args_Size = This.Size
End Property

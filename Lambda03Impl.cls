VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda03Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Primitives.Concrete"
Option Explicit
Implements Lambda03

Private Type Data
  Fn As String
  Args As Args
End Type

Private This As Data

Public Function Init(Fn As String, Optional Args As Args = Nothing) As Lambda03
  This.Fn = Fn
  Set This.Args = Args
  Set Init = Me
End Function

Private Function Lambda03_Add(Item As Variant) As Lambda02
	Dim Temp As Args

  With New Arg
    Set Temp = .Init(Item, This.Args)
  End With
  
  With New Lambda02Impl
    Set Lambda03_Add = .Init(This.Fn, Temp)
  End With

  Set Temp = Nothing
End Function
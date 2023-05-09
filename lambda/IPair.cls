VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Abstraction"
Option Explicit

Public Property Get Item1() As Variant: End Property
Public Property Get Item2() As Variant: End Property
Public Function Map(ByRef Func As Variant) As IPair: End Function
Public Property Get Show() As String
Attribute Show.VB_UserMemId = 0
End Property

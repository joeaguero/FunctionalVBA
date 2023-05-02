VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Immutable Collection Wrapper"
Option Explicit

Public Property Get Count() As Long
Attribute Count.VB_Description = "Count of Items in Collection"
End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_Description = "Retrieve the item based on the index"
End Property

Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for For Each"
Attribute NewEnum.VB_UserMemId = -4
End Property

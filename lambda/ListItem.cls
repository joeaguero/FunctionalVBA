VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Library.Data.Concrete"
Option Explicit

Implements IList

Private Type Data
  Head As Variant
  Tail As IList
End Type

Private This As Data

Public Function Init(ByRef Head As Variant, ByRef Tail As IList) As IList
  LetSet(This.Head) = Head
  Set This.Tail = Tail
  Set Init = Me
End Function

Private Function IList_Bind(ByRef Func As Variant) As IList
  Set IList_Bind = Append.Of(Func.Of(This.Head)).Of(This.Tail.Bind(Func))
End Function

Private Property Get IList_Head() As Variant
  LetSet(IList_Head) = This.Head
End Property

Private Function IList_Reduce(Func As ILambda2, Default As Variant) As Variant
  LetSet(IList_Reduce) = Func.Of(This.Head).Of(This.Tail.Reduce(Func, Default))
End Function

Private Function IList_ReduceL(Func As ILambda2, Default As Variant) As Variant
  LetSet(IList_ReduceL) = This.Tail.ReduceL(Func, Func.Of(This.Head).Of(Default))
End Function

Private Property Get IList_Tail() As IList
  Set IList_Tail = This.Tail
End Property

Private Property Get IList_IsEmpty() As Boolean
  IList_IsEmpty = False
End Property

Private Property Get IList_Show() As String
  Dim Sep As String
  Select Case TypeName(This.Head)
    Case "ListItem"
      Sep = " "
    Case "ListNull"
      Sep = " "
    Case "Pair"
      Sep = " "
    Case "Just"
      Sep = " "
    Case "None"
      Sep = " "
    Case Else
      ' In case of List[Str] print as a string
      If VarType(This.Head) = vbByte Then IList_Show = """" & Decode(Me) & """": Exit Property
      Sep = " "
  End Select
  
  ' Default to printing as a list
  IList_Show = "[ " & IList_Reduce(TextJoin.Of("," & Sep), "") & " ]"
End Property

Private Function IList_Map(ByRef Func As Variant) As IList
  Set IList_Map = Cons.Of(Func.Of(This.Head)).Of(This.Tail.Map(Func))
End Function

Private Sub Class_Terminate()
  Set This.Head = Nothing
  Set This.Tail = Nothing
End Sub

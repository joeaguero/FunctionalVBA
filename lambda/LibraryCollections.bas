Attribute VB_Name = "LibraryCollections"
'@Folder("Library.Definitions")
Option Explicit
Option Private Module

Public Function Show(ByRef Item As Variant) As String
  Select Case TypeName(Item)
    Case "ListItem"
      Show = Item.Show: Exit Function
    Case "ListNull"
      Show = Item.Show: Exit Function
    Case "Pair"
      Show = Item.Show: Exit Function
    Case "None"
      Show = Item.Show: Exit Function
    Case "Just"
      Show = Item.Show: Exit Function
  End Select
  
  If VarType(Item) = vbByte Then Show = "'" & Chr(Item) & "'" Else Show = CStr(Item)
End Function

' Combiners
Private Function ZipWith(ByRef Operator As ILambda2, ByRef LHS As IList, ByRef RHS As IList) As IList
  If LHS.IsEmpty Or RHS.IsEmpty Then
    Set ZipWith = ListNull
  Else
    Set ZipWith = Cons(Operator.Of(LHS.Head).Of(RHS.Head), ZipWith(Operator, LHS.Tail, RHS.Tail))
  End If
End Function

' Generators
Private Function Range(ByVal Start As Long, ByVal Count As Long, ByVal Step As Long) As IList
  If Count < 0 Then Step = Step * -1: Count = Abs(Count)
  If Count = 0 Then
    Set Range = ListNull
  Else
    Set Range = Cons(Start, Range(Start + Step, Count - 1, Step))
  End If
End Function

Private Function RangeDouble(ByVal Start As Double, ByVal Count As Long, ByVal Step As Double) As IList
  If Count < 0 Then Step = Step * -1: Count = Abs(Count)
  If Count = 0 Then
    Set RangeDouble = ListNull
  Else
    Set RangeDouble = Cons(Start, RangeDouble(Start + Step, Count - 1, Step))
  End If
End Function

' Functor
Private Function Map(ByRef Func As Variant, ByRef Functor As Variant) As Variant
  Set Map = Functor.Map(Func)
End Function

' Monad
Private Function Bind(ByRef Func As Variant, ByRef Monad As Variant) As Variant ' Monad a
  Set Bind = Monad.Bind(Func)
End Function

' Foldable
Private Function Fold(ByRef Func As ILambda2, ByRef Default As Variant, ByRef Foldable As Variant) As Variant
  LetSet(Fold) = Foldable.Reduce(Func, Default)
End Function

Private Function Foldl(ByRef Func As ILambda2, ByRef Default As Variant, ByRef Foldable As Variant) As Variant
  LetSet(Foldl) = Foldable.ReduceL(Func, Default)
End Function

Private Function IsEmpty(ByRef Nullable As Variant) As Boolean
  IsEmpty = Nullable.IsEmpty
End Function

' Pair
Private Function Pair(ByRef Item1 As Variant, ByRef Item2 As Variant) As IPair
  With New Pair
    Set Pair = .Init(Item1, Item2)
  End With
End Function

Private Function Fst(ByRef Source As IPair) As Variant
  LetSet(Fst) = Source.Item1
End Function

Private Function Snd(ByRef Source As IPair) As Variant
  LetSet(Snd) = Source.Item2
End Function


' Maybe
Private Function Just(ByRef Value As Variant) As IMaybe
  With New Just
    Set Just = .Init(Value)
  End With
End Function


' List
Private Function Cons(ByRef Value As Variant, ByRef Source As IList) As IList
  With New ListItem
    Set Cons = .Init(Value, Source)
  End With
End Function

Private Function Head(ByRef Source As IList) As IMaybe
  If IsEmpty(Source) Then Set Head = None Else Set Head = Just(Source.Head)
End Function

Private Function Tail(ByRef Source As IList) As IList
  Set Tail = Source.Tail
End Function

Private Function Append(ByRef LHS As IList, ByRef RHS As IList) As IList
  Set Append = Fold(Lib.Cons, RHS, LHS)
End Function

Private Function ToList(ByRef Items As Variant) As IList
  Set ToList = ListNull
  Dim I As Long: For I = UBound(Items) To LBound(Items) Step -1
    Set ToList = Cons(Items(I), ToList)
  Next I
End Function

Private Function Filter(ByRef Predicate As ILambda1, ByRef Source As IList) As IList
  If IsEmpty(Source) Then
    Set Filter = ListNull
  ElseIf Predicate.Of(Source.Head) Then
    Set Filter = Cons(Source.Head, Filter(Predicate, Source.Tail))
  Else
    Set Filter = Filter(Predicate, Source.Tail)
  End If
End Function

Private Function SplitWith(ByRef Predicate As ILambda1, ByRef Source As IList) As IList
  Set SplitWith = Fold(SplitWithHelper(Predicate), ListNull, Source)
End Function

Private Function SplitWithHelper() As ILambda3: Set SplitWithHelper = Lambda3("SplitWithHelperFunc"): End Function

Private Function SplitWithHelperFunc(ByRef Predicate As ILambda1, ByRef LHS As Variant, ByRef RHS As IList) As IList
  If Predicate(LHS) Then
    Set SplitWithHelperFunc = Cons(List(), RHS)
  ElseIf IsEmpty(RHS) Then
    Set SplitWithHelperFunc = Cons(List(LHS), RHS)
  Else
    Set SplitWithHelperFunc = Cons(Cons(LHS, RHS.Head), RHS.Tail)
  End If
End Function

Private Function Lines(ByRef RHS As IList) As IList
  Set Lines = Filter(Lib.Compose(Length)(GT(0)), SplitWith(EQ(CByte(13)), Filter(Lib.Compose(EQ(CByte(10)))(NotBool), RHS)))
End Function

' Purely Functional
Private Function Compose(ByRef Func1 As Variant, ByRef Func2 As Variant, ByRef Arg As Variant) As Variant
  LetSet(Compose) = Func2.Of(Func1.Of(Arg))
End Function

Private Function Apply(ByRef Arg As Variant, ByRef Func As Variant) As Variant
  LetSet(Apply) = Func.Of(Arg)
End Function



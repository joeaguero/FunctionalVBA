Attribute VB_Name = "Lib"
'@Folder "Library"
Option Explicit
Option Private Module

' Operators
Function Addition() As ILambda2: Set Addition = Lambda2("LibraryOperators.Addition"): End Function
Function Constant() As ILambda2: Set Constant = Lambda2("LibraryOperators.Constant"): End Function
Function TextJoin() As ILambda3: Set TextJoin = Lambda3("LibraryOperators.TextJoin"): End Function
Function AppendArray() As ILambda2: Set AppendArray = Lambda2("LibraryOperators.AppendArray"): End Function
Function EQ() As ILambda2: Set EQ = Lambda2("LibraryOperators.EQ"): End Function
Function GT() As ILambda2: Set GT = Lambda2("LibraryOperators.GT"): End Function
Function LT() As ILambda2: Set LT = Lambda2("LibraryOperators.LT"): End Function
Function NotBool() As ILambda1: Set NotBool = Lambda1("LibraryOperators.NotBool"): End Function

' List of Pairs
Function Zip() As ILambda2: Set Zip = ZipWith.Of(Pair): End Function


' Arrays
  ' Constructors
  Function ToArray() As ILambda1: Set ToArray = Foldl.Of(AppendArray).Of(EmptyArray): End Function

' Lists
  ' Constructors
  Function Cons() As ILambda2: Set Cons = Lambda2("LibraryCollections.Cons"): End Function
  Function Append() As ILambda2: Set Append = Lambda2("LibraryCollections.Append"): End Function
  Function ToList() As ILambda1: Set ToList = Lambda1("LibraryCollections.ToList"): End Function
  
  ' Generators
  Function Range() As ILambda3: Set Range = Lambda3("LibraryCollections.Range"): End Function
  Function RangeDouble() As ILambda3: Set RangeDouble = Lambda3("LibraryCollections.RangeDouble"): End Function
  
  'Properties
  Function Length() As ILambda1: Set Length = Compose.Of(Map.Of(Constant.Of(1))).Of(Sum): End Function
  
  ' Operations
  Function ZipWith() As ILambda3: Set ZipWith = Lambda3("LibraryCollections.ZipWith"): End Function
  Function Filter() As ILambda2: Set Filter = Lambda2("LibraryCollections.Filter"): End Function
  Function SplitWith() As ILambda2: Set SplitWith = Lambda2("LibraryCollections.SplitWith"): End Function
  Function Words() As ILambda1: Set Words = Compose(SplitWith(EQ(CByte(32))))(Filter(Compose(Length)(GT(0)))): End Function
  Function UnWords() As ILambda1: Set UnWords = Fold(JoinWith(Str(" ")))(Str("")): End Function
  Function Lines() As ILambda1: Set Lines = Lambda1("LibraryCollections.Lines"): End Function
  Function UnLines() As ILambda1: Set UnLines = Fold(JoinWith(Str(vbNewLine)))(Str("")): End Function
  Function JoinWith() As ILambda3: Set JoinWith = Lambda3("LibraryOperators.JoinWith"): End Function
  
  ' Data Access
  Function Head() As ILambda1: Set Head = Lambda1("LibraryCollections.Head"): End Function
  Function Tail() As ILambda1: Set Tail = Lambda1("LibraryCollections.Tail"): End Function


' Pairs
  ' Constructors
  Function Pair() As ILambda2: Set Pair = Lambda2("LibraryCollections.Pair"): End Function
  
  ' Data Access
  Function Fst() As ILambda1: Set Fst = Lambda1("LibraryCollections.Fst"): End Function
  Function Snd() As ILambda1: Set Snd = Lambda1("LibraryCollections.Snd"): End Function


' Maybes
  ' Constructors
  Function Just() As ILambda1: Set Just = Lambda1("LibraryCollections.Just"): End Function


' Foldable
Function Fold() As ILambda3: Set Fold = Lambda3("LibraryCollections.Fold"): End Function
Function Foldl() As ILambda3: Set Foldl = Lambda3("LibraryCollections.Foldl"): End Function
Function IsEmpty() As ILambda1: Set IsEmpty = Lambda1("LibraryCollections.IsEmpty"): End Function
Function Sum() As ILambda1: Set Sum = Fold.Of(Addition).Of(0): End Function


' Functors
Function Map() As ILambda2: Set Map = Lambda2("LibraryCollections.Map"): End Function


' Applicative
Function Apply() As ILambda2: Set Apply = Lambda2("LibraryCollections.Apply"): End Function


' Monads
Function Bind() As ILambda2: Set Bind = Lambda2("LibraryCollections.Bind"): End Function


' Composition
Function Compose() As ILambda3: Set Compose = Lambda3("LibraryCollections.Compose"): End Function
Function Sequence() As ILambda1: Set Sequence = Compose.Of(Apply).Of(Map): End Function




VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Primitives"
Option Explicit
Implements ILambda

Private Type Data
  Fn As String
  Args As IArgs
End Type

Private This As Data

Public Function Init(Fn As String, Optional Args As IArgs = Nothing) As ILambda
  This.Fn = Fn
  Set This.Args = Args
  Set Init = Me
End Function

Private Function ILambda_Add(ParamArray Items() As Variant) As ILambda
  Dim Temp As IArgs: Set Temp = This.Args
  Dim I As Long: For I = LBound(Items) To UBound(Items)
    With New Args
      Set Temp = .Init(Items(I), Temp)
    End With
  Next I
  
  With New Lambda
    Set ILambda_Add = .Init(This.Fn, Temp)
  End With
  Set Temp = Nothing
End Function

Private Function ILambda_Run() As Variant
  If This.Args Is Nothing Then LetSet(ILambda_Run) = Application.Run(This.Fn): Exit Function
  Dim Arg0 As IArgs, Arg1 As IArgs, Arg2 As IArgs, Arg3 As IArgs, Arg4 As IArgs
  Dim Arg5 As IArgs, Arg6 As IArgs, Arg7 As IArgs, Arg8 As IArgs, Arg9 As IArgs
  
  Select Case This.Args.Size
    Case 1
      Set Arg0 = This.Args
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg0.Arg)
      
    Case 2
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg1.Arg, Arg0.Arg)
      
    Case 3
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 4
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 5
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 6
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 7
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 8
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      Set Arg7 = Arg6.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 9
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      Set Arg7 = Arg6.Prev
      Set Arg8 = Arg7.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg8.Arg, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 10
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      Set Arg7 = Arg6.Prev
      Set Arg8 = Arg7.Prev
      Set Arg9 = Arg8.Prev
      LetSet(ILambda_Run) = Application.Run(This.Fn, Arg9.Arg, Arg8.Arg, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case Else
      Err.Raise 10000 + This.Args.Size, "Curry.Run()", "Not Implemented: Too Many Arguments (" & This.Args.Size & ")"
  End Select
End Function

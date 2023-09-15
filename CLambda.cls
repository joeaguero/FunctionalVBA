VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLambda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Primitives"
Option Explicit
Implements Lambda

Private Type Data
  Fn As String
  Args As Args
End Type

Private This As Data

Public Function Init(Fn As String, Optional Args As Args = Nothing) As Lambda
  This.Fn = Fn
  Set This.Args = Args
  Set Init = Me
End Function

Private Function Lambda_Add(ParamArray Items() As Variant) As Lambda
  Dim Temp As Args: Set Temp = This.Args
  Dim I As Long: For I = LBound(Items) To UBound(Items)
    With New CArgs
      Set Temp = .Init(Items(I), Temp)
    End With
  Next I
  
  With New CLambda
    Set Lambda_Add = .Init(This.Fn, Temp)
  End With
  Set Temp = Nothing
End Function

Private Function Lambda_Run() As Variant
  If This.Args Is Nothing Then LetSet(Lambda_Run) = Application.Run(This.Fn): Exit Function

  Dim Arg0 As Args, Arg1 As Args, Arg2 As Args, Arg3 As Args, Arg4 As Args
  Dim Arg5 As Args, Arg6 As Args, Arg7 As Args, Arg8 As Args, Arg9 As Args
  
  Select Case This.Args.Size
    Case 1
      Set Arg0 = This.Args
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg0.Arg)
      
    Case 2
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg1.Arg, Arg0.Arg)
      
    Case 3
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 4
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 5
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 6
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 7
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case 8
      Set Arg0 = This.Args
      Set Arg1 = Arg0.Prev
      Set Arg2 = Arg1.Prev
      Set Arg3 = Arg2.Prev
      Set Arg4 = Arg3.Prev
      Set Arg5 = Arg4.Prev
      Set Arg6 = Arg5.Prev
      Set Arg7 = Arg6.Prev
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
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
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg8.Arg, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
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
      LetSet(Lambda_Run) = Application.Run(This.Fn, Arg9.Arg, Arg8.Arg, Arg7.Arg, Arg6.Arg, Arg5.Arg, Arg4.Arg, Arg3.Arg, Arg2.Arg, Arg1.Arg, Arg0.Arg)
      
    Case Else
      Err.Raise 10000 + This.Args.Size, "Lambda.Run()", "Not Implemented: Too Many Arguments (" & This.Args.Size & ")"
  End Select
End Function

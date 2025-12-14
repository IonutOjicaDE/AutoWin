Attribute VB_Name = "CommandsLoop"
Option Explicit
Private Const MODULE_NAME As String = "CommandsLoop"

Enum ForDoLoopsType
  NoLoop = 0
  LoopFor
  LoopForEach
  DoInfinite
  DoWhile
  DoUntil
  LoopWhile
  LoopUntil
  LoopInfinite
  LoopGoSub
  LoopGoTo
End Enum

Private Type LoopInfo
  LoopType     As ForDoLoopsType
  ident        As String  ' User-defined name from Arg1
  counterCell  As Range   ' Cell reference of Arg2 counter (For loops) or condition (Do loops)
  counterValue As Long    ' For loops
  startValue   As Long    ' For loops
  endValue     As Long    ' For loops
  step         As Long    ' For loops
  colCounterValue As Long ' ForEach loops
  colStartValue   As Long ' ForEach loops
  colEndValue     As Long ' ForEach loops
  colStep         As Long ' ForEach loops
End Type

Private loopStack()       As LoopInfo ' Stack to track loops
Private loopStackSize     As Long     ' Tracks the number of active loops
Private loopStackCapacity As Long     ' Tracks current allocated array size

Private tmpL As Long, tmpS As String, tmpR As Range


Public Function RegisterCommandsLoop()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "sub", Array("SubToGoSub", "Sub", _
    MODULE_NAME, "Place a Sub to execute usind Go Sub command." & loopListSub, _
    "SubName", "Enter the name of the Sub to be executed.")
  commandMap.Add "gosub", Array("GoSubs", "Go Sub", _
    MODULE_NAME, "Execute a Sub.", _
    "SubName", "Enter the name of the Sub to be executed." & loopListSub)

  commandMap.Add "ifthenexecutesub", Array("CommandIfThenExecuteSub", "If Then Execute Sub", MODULE_NAME, _
    "Execute the Sub according to a condition.", _
    "Condition", "Condition that should return true or false.", _
    "SubName if True", "What line or label should be next if condition is true." & loopListSub, _
    "SubName if False", "What line or label should be next if condition is false." & loopListSub)

  commandMap.Add "endsub", Array("EndSubs", "End Sub", _
    MODULE_NAME, "Terminate a Sub.", _
    "SubName", "Optional argument, you can enter it for checking purposes." & loopListSub)


  commandMap.Add "label", Array("LabelToGoTo", "Label", _
    MODULE_NAME, "Place a Label to return to usind Go To command." & loopListLabel, _
    "Label", "Specify the label of this line, to jump to it using the Go To command.")
  commandMap.Add "goto", Array("GoToLabels", "Go To", _
    MODULE_NAME, "Continue the execution on another line or starting from a specified label", _
    "Line or label", "Enter new line number or label name where to continue the execution." & loopListLabel, _
    "Return wanted", "If left empty, then it will be possible to return to current line using Return command")
    
  commandMap.Add "return", Array("ReturnToGoto", "Return", _
    MODULE_NAME, "Return to a last Go To command", _
    "Label", "Optional you can return to a Go To command with a specified label." & loopListLabel)


  commandMap.Add "do", Array("DoStart", "Do", _
    MODULE_NAME, "Start of a loop (you can insert an infinite loop or specify the exit condition on the Loop command)." & loopListLoop, _
    "Ident", "Enter the name of the loop, to identify it later")
    
  commandMap.Add "dowhile", Array("DoStart", "Do While", _
    MODULE_NAME, "Start of a do loop and specify a condition that must be true while the loop will repeat." & loopListLoop, _
    "Ident", "Enter the name of the loop, to identify it later", _
    "Condition", "When false, then stop")
    
  commandMap.Add "dountil", Array("DoStart", "Do Until", _
    MODULE_NAME, "Start of a do loop and specify a condition that must be true to exit the loop." & loopListLoop, _
    "Ident", "Enter the name of the loop, to identify it later", _
    "Condition", "When true, then stop")
    
  commandMap.Add loopTypeLoop, Array("DoLoop", "Loop", _
    MODULE_NAME, "Finish of a do loop", _
    "Ident", "Enter the name of the loop, to identify it." & loopListLoop)
    
  commandMap.Add loopTypeLoop & "while", Array("DoLoop", "Loop While", _
    MODULE_NAME, "Finish of a do loop", _
    "Ident", "Enter the name of the loop, to identify it." & loopListLoop, _
    "Condition", "When true, then stop")
    
  commandMap.Add loopTypeLoop & "until", Array("DoLoop", "Loop Until", _
    MODULE_NAME, "Finish of a do loop", _
    "Ident", "Enter the name of the loop, to identify it." & loopListLoop, _
    "Condition", "When true, then stop")
    
  commandMap.Add "exitdo", Array("ExitDo", "Exit Do", _
    MODULE_NAME, "Stop the current Do loop and jump right after the loop", _
    "Ident", "Enter the name of the loop to exit." & loopListLoop)


  commandMap.Add "for", Array("ForNormalStart", "For", _
    MODULE_NAME, "Start a For-Next loop. Rembember to add also the Next command where the loop exit." & loopListFor, _
    "Ident", "Enter the name of the loop, to identify it (optional).", _
    "Counter", "Here will be written on each start of the loop the current value.", _
    "Start", "Value to start with.", _
    "Finish", "Value to finish with.", _
    "Step", "Step value that will be added to the Counter after each loop (optional; if empty then +1 if Start<=Finish, else -1).")
    
  commandMap.Add "foreach", Array("ForEachStart", "For Each", _
    MODULE_NAME, "Loop trough each cell of a range of cells. Rembember to add also the Next command where the loop exit." & loopListFor, _
    "Ident", "Enter the name of the loop, to identify it (optional).", _
    "Counter", "Here will be written on each start of the loop the current cell address. Any content will be overwritten.", _
    "Start", "Cell Address or reference to a cell using a formula to start the looping.", _
    "Finish", "Cell Address or reference to a cell using a formula to finish the looping.")
    
  commandMap.Add loopTypeNext, Array("ForNext", "Next", _
    MODULE_NAME, "Finish of a For loop or a For Each loop. This is the place where it will be decided if the loop continues or it will exit.", _
    "Ident", "Enter the name of the loop, to identify it. If empty then the last loop will be considered." & loopListFor)
    
  commandMap.Add "exitfor", Array("ExitFor", "Exit For", _
    MODULE_NAME, "Stop the current For loop and jump right after the loop", _
    "Ident", "Enter the name of the loop to exit. If empty then the last loop will be considered." & loopListFor)


  loopStackSize = 0&
  loopStackCapacity = 0&

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsLoop", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsLoop()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsLoop", Err.Number, Err.Source, Err.Description, Erl
End Function




Public Function SubToGoSub(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  SubToGoSub = True
End Function

Public Function GoSubs(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (required)
  On Error GoTo eh
  Call ExecuteSub(CStr(currentRowArray(1&, ColAArg1)), "Arg1")

NormalExecution:
  GoSubs = True
  Exit Function
eh:
  GoSubs = False
  RaiseError MODULE_NAME & ".GoSubs", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function CommandIfThenExecuteSub(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: Condition that should return true or false
' Arg2: What line or label should be next if condition is true
' Arg3: What line or label should be next if condition is false
  On Error GoTo eh
  If GetBoolean(CStr(currentRowArray(1, ColAArg1))) Then
    Call ExecuteSub(CStr(currentRowArray(1, ColAArg1 + 1)), "Arg2")

  Else ' Arg1 is false
    Call ExecuteSub(CStr(currentRowArray(1, ColAArg1 + 2)), "Arg3")

  End If

NormalExecution:
  CommandIfThenExecuteSub = True
  Exit Function
eh:
  CommandIfThenExecuteSub = False
  RaiseError MODULE_NAME & ".CommandIfThenExecuteSub", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Private Function ExecuteSub(SubName As String, ArgName As String) As Boolean
  On Error GoTo eh
  Dim tmpLoop As LoopInfo
  With tmpLoop
    .ident = CleanString(SubName)

    If Len(.ident) > 0 Then
      Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeSub, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)

      If Not tmpR Is Nothing Then

        tmpL = tmpR.Row
        Do
          If StrComp(CleanString(tmpR.Offset(0, 1).Text), .ident, vbTextCompare) = 0 Then
            .LoopType = LoopGoSub
            .counterValue = currentRow
            Set .counterCell = currentRowRange(1&, ColAArg1)
            If Not PushLoopStack(tmpLoop) Then
              ExecuteSub = False
              RaiseError MODULE_NAME & ".ExecuteSub", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 1
              Exit Function
            End If
            currentRow = tmpR.Row
            GoTo NormalExecution
          End If
          Set tmpR = shAuto.Columns(ColACommand).FindNext(tmpR)
        Loop Until tmpR.Row = tmpL
      End If

      ExecuteSub = False
      RaiseError MODULE_NAME & ".ExecuteSub", Err.Number, Err.Source, "No Sub with the ident found: " & ArgName & "=[" & SubName & "]", Erl, 2
      Exit Function
    Else

      ExecuteSub = False
      RaiseError MODULE_NAME & ".ExecuteSub", Err.Number, Err.Source, "Name of the Sub is needed in " & ArgName, Erl, 3
      Exit Function
    End If
  End With

NormalExecution:
  ExecuteSub = True
  Exit Function
eh:
  ExecuteSub = False
  RaiseError MODULE_NAME & ".ExecuteSub", Err.Number, Err.Source, Err.Description, Erl
End Function


Public Function EndSubs(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional)
  On Error GoTo eh
  If loopStackSize = 0& Then GoTo NormalStopExecution ' it is the main Sub, so stop execution


  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                            ' we have an ident of a Sub to be returned
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopGoSub) And _
          loopStack(tmpL).ident = ident Then _
            Exit For                                 ' found the GoSub with the same ident
    Next tmpL
  Else                                               ' no ident of a GoSub to be returned provided; then let the last GoSub be
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopGoSub) Then _
            Exit For                                 ' found the last GoSub
    Next tmpL
  End If


  If tmpL = 0& Then GoTo NormalStopExecution          ' no GoSub found, so it is the main Sub, so stop execution


  currentRow = loopStack(tmpL).counterValue

  Do While loopStackSize >= tmpL: PopLoopStack: Loop ' remove any preceding loops

  GoTo NormalExecution

NormalStopExecution:
  stopExecutionRequired = True
NormalExecution:
  EndSubs = True
  Exit Function
eh:
  EndSubs = False
  RaiseError MODULE_NAME & ".EndSubs", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function




Public Function LabelToGoTo(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  LabelToGoTo = True
End Function

Public Function GoToLabels(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (required)
' Arg2 = empty if return wanted or enter text if no return wanted
  On Error GoTo eh
  Dim tmpLoop As LoopInfo
  With tmpLoop
    .ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))

    If Len(.ident) > 0 Then

      If IsNumber(.ident) Then
        tmpL = CLng(.ident)
        If tmpL = currentRow Then
          GoToLabels = False
          RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, "Go to same row? It will be an infinite loop!", Erl, 1, ExecutingTroughApplicationRun
          Exit Function
        End If
        .LoopType = LoopGoTo
        .counterValue = currentRow
        Set .counterCell = currentRowRange(1&, ColAArg1)
        If Len(currentRowArray(1&, ColAArg1 + 1)) > 0 Then
          If Not PushLoopStack(tmpLoop) Then
            GoToLabels = False
            RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 2, ExecutingTroughApplicationRun
            Exit Function
          End If
        End If
        currentRow = tmpL
        GoTo NormalExecution

      Else ' ident is not a number, but a text
        Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeLabel, LookIn:=xlValues, LookAt:=xlWhole, SearchDirection:=xlNext, MatchCase:=False)
  
        If Not tmpR Is Nothing Then
  
          tmpL = tmpR.Row
          Do
            If StrComp(CleanString(tmpR.Offset(0, 1).Text), .ident, vbTextCompare) = 0 Then
              .LoopType = LoopGoTo
              .counterValue = currentRow
              Set .counterCell = currentRowRange(1&, ColAArg1)
              If Len(currentRowArray(1&, ColAArg1 + 1)) > 0 Then
                If Not PushLoopStack(tmpLoop) Then
                  GoToLabels = False
                  RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 3, ExecutingTroughApplicationRun
                  Exit Function
                End If
              End If
              currentRow = tmpR.Row
              GoTo NormalExecution
            End If
            Set tmpR = shAuto.Columns(ColACommand).FindNext(tmpR)
          Loop Until tmpR.Row = tmpL
        End If

        GoToLabels = False
        RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, "No Label with the ident found: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 4, ExecutingTroughApplicationRun
        Exit Function
      End If

    Else

      GoToLabels = False
      RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, "Name of the Label is needed in Arg1", Erl, 5, ExecutingTroughApplicationRun
      Exit Function
    End If
  End With

NormalExecution:
  GoToLabels = True
  Exit Function
eh:
  GoToLabels = False
  RaiseError MODULE_NAME & ".GoToLabels", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function



Public Function ReturnToGoto(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional)
  On Error GoTo eh
  If loopStackSize = 0& Then GoTo NormalStopExecution ' it is the main Sub, so stop execution


  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                            ' we have an ident of a GoTo to be returned
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopGoTo) And _
          loopStack(tmpL).ident = ident Then _
            Exit For                                 ' found the GoTo with the same ident
    Next tmpL
  Else                                               ' no ident of a GoTo to be returned provided; then let the last GoTo be
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopGoTo) Then _
            Exit For                                 ' found the last GoTo
    Next tmpL
  End If


  If tmpL = 0& Then GoTo NormalStopExecution          ' no GoTo found, so it is the main Sub, so stop execution
  
  currentRow = loopStack(tmpL).counterValue

  Do While loopStackSize >= tmpL: PopLoopStack: Loop ' remove any preceding loops


  GoTo NormalExecution

NormalStopExecution:
  stopExecutionRequired = True
NormalExecution:
  ReturnToGoto = True
  Exit Function
eh:
  ReturnToGoto = False
  RaiseError MODULE_NAME & ".ReturnToGoto", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function





Public Function DoStart(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional), Arg2 = condition or empty
  On Error GoTo eh
  Dim tmpLoop As LoopInfo

  With tmpLoop
    .ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
    .counterValue = currentRow
    tmpS = currentRowArray(1&, ColACommand)
    Set .counterCell = currentRowRange(1&, ColAArg1 + 1&)
    
    If InStr(tmpS, "while", vbTextCompare) > 0 Then
      .LoopType = DoWhile                ' when false, then stop
      If Not .counterCell Then
        If Not SkipToMatchingNext(tmpLoop) Then
          DoStart = False
          RaiseError MODULE_NAME & ".DoStart", Err.Number, Err.Source, "Loop is ending, without even to begin, but matching next not found.", Erl, 1, ExecutingTroughApplicationRun
          Exit Function
        End If
        GoTo NormalExecution
      End If
    ElseIf InStr(tmpS, "until", vbTextCompare) > 0 Then
      .LoopType = DoUntil                ' when true, then stop
      If .counterCell Then
        If Not SkipToMatchingNext(tmpLoop) Then
          DoStart = False
          RaiseError MODULE_NAME & ".DoStart", Err.Number, Err.Source, "Loop is ending, without even to begin, but matching next not found.", Erl, 2, ExecutingTroughApplicationRun
          Exit Function
        End If
        GoTo NormalExecution
      End If
    Else
      .LoopType = DoInfinite
    End If
  End With

  If Not PushLoopStack(tmpLoop) Then
    DoStart = False
    RaiseError MODULE_NAME & ".DoStart", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 3, ExecutingTroughApplicationRun
    Exit Function
  End If

NormalExecution:
  DoStart = True
  Exit Function
eh:
  DoStart = False
  RaiseError MODULE_NAME & ".DoStart", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function




Public Function ExitDo(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional)
  On Error GoTo eh
  If loopStackSize = 0& Then
    ExitDo = False
    RaiseError MODULE_NAME & ".ExitDo", Err.Number, Err.Source, "No Do loop found to exit", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                                ' we have an ident of a Do to be Loop-ed
    For tmpL = loopStackSize To 1& Step -1&
      Select Case loopStack(tmpL).LoopType
        Case DoInfinite, DoWhile, DoUntil, LoopWhile, LoopUntil, LoopInfinite
          If loopStack(tmpL).ident = ident Then Exit For ' found the Do with the same ident
      End Select
    Next tmpL
  Else                                                   ' no ident of a Do to be Loop-ed provided; then let the last Do be
    For tmpL = loopStackSize To 1& Step -1&
      Select Case loopStack(tmpL).LoopType
        Case DoInfinite, DoWhile, DoUntil, LoopWhile, LoopUntil, LoopInfinite
          Exit For                                       ' found the last Do
      End Select
    Next tmpL
  End If

  If tmpL = 0& Then
    ExitDo = False
    RaiseError MODULE_NAME & ".ExitDo", Err.Number, Err.Source, "Do loop found to exit", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  If Not SkipToMatchingNext(loopStack(tmpL)) Then
    ExitDo = False
    RaiseError MODULE_NAME & ".ExitDo", Err.Number, Err.Source, "Do loop is exiting, but matching next not found.", Erl, 3, ExecutingTroughApplicationRun
    Exit Function
  End If

  Do While loopStackSize >= tmpL: PopLoopStack: Loop    ' remove any preceding loops along with the found loop

NormalExecution:
  ExitDo = True
  Exit Function
eh:
  ExitDo = False
  RaiseError MODULE_NAME & ".ExitDo", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function




Public Function DoLoop(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional), Arg2 = condition or empty
  On Error GoTo eh
  If loopStackSize = 0& Then
    DoLoop = False
    RaiseError MODULE_NAME & ".DoLoop", Err.Number, Err.Source, "No Do loop was started", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                                ' we have an ident of a Do to be Loop-ed
    For tmpL = loopStackSize To 1& Step -1&
      Select Case loopStack(tmpL).LoopType
        Case DoInfinite, DoWhile, DoUntil, LoopWhile, LoopUntil, LoopInfinite
          If loopStack(tmpL).ident = ident Then Exit For ' found the Do with the same ident
      End Select
    Next tmpL
  Else                                                   ' no ident of a Do to be Loop-ed provided; then let the last Do be
    For tmpL = loopStackSize To 1& Step -1&
      Select Case loopStack(tmpL).LoopType
        Case DoInfinite, DoWhile, DoUntil, LoopWhile, LoopUntil, LoopInfinite
          Exit For                                       ' found the last Do
      End Select
    Next tmpL
  End If


  If tmpL = 0& Then
    DoLoop = False
    RaiseError MODULE_NAME & ".DoLoop", Err.Number, Err.Source, "Do loop was started", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  Do While loopStackSize > tmpL: PopLoopStack: Loop      ' remove any preceding loops


  With loopStack(tmpL)

    If .LoopType = DoInfinite Then
      tmpS = currentRowArray(1&, ColACommand)
      Set .counterCell = currentRowRange(1&, ColAArg1 + 1&)
      
      If InStr(tmpS, "while", vbTextCompare) > 0 Then
        .LoopType = LoopWhile
      ElseIf InStr(tmpS, "until", vbTextCompare) > 0 Then
        .LoopType = LoopUntil
      Else
        .LoopType = LoopInfinite
      End If
    End If

    Select Case .LoopType
      Case DoWhile, LoopWhile ' when false, then stop
        If Not .counterCell Then PopLoopStack Else currentRow = .counterValue
      Case DoUntil, LoopUntil ' when true, then stop
        If .counterCell Then PopLoopStack Else currentRow = .counterValue
      Case LoopInfinite
        currentRow = .counterValue
    End Select

  End With

NormalExecution:
  DoLoop = True
  Exit Function
eh:
  DoLoop = False
  RaiseError MODULE_NAME & ".DoLoop", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function






Public Function ForNormalStart(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional)
' Arg2 = empty (counter), Arg3 = min value, Arg4 = max value
' Arg5 = step value (optional, if empty then +1 if Start<=Finish, else -1)
  On Error GoTo eh
  Dim tmpLoop As LoopInfo
  Dim startValueS As String, endValueS As String, stepS As String
  startValueS = currentRowArray(1&, ColAArg1 + 2&)
  endValueS = currentRowArray(1&, ColAArg1 + 3&)
  stepS = currentRowArray(1&, ColAArg1 + 4&)
  
  With tmpLoop
    .ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
    .LoopType = LoopFor
    .counterValue = startValueS
    Set .counterCell = currentRowRange(1&, ColAArg1 + 1&)
    
    If IsNumber(startValueS) And IsNumber(endValueS) Then
      .startValue = CLng(startValueS): .endValue = CLng(endValueS)

      If IsNumber(stepS) Then
        .step = CLng(stepS)

        If .step > 0& Then
          If .startValue > .endValue Then
            If Not SkipToMatchingNext(tmpLoop) Then
              ForNormalStart = False
              RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "For loop is ending, without even to begin, but matching next not found.", Erl, 1, ExecutingTroughApplicationRun
              Exit Function
            End If
            GoTo NormalExecution
          End If

        ElseIf .step < 0& Then
          If .startValue < .endValue Then
            If Not SkipToMatchingNext(tmpLoop) Then
              ForNormalStart = False
              RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "For loop is ending, without even to begin, but matching next not found.", Erl, 2, ExecutingTroughApplicationRun
              Exit Function
            End If
            GoTo NormalExecution
          End If

        Else ' StepL=0
          ' What loop is this, with a Step of 0 ? It would be an infinite loop. So skip the loop. This case should be verified beforehand.
          If Not SkipToMatchingNext(tmpLoop) Then
            ForNormalStart = False
            RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "For loop will be skipped because Step=0, meaning an infinite loop, but matching next not found.", Erl, 3, ExecutingTroughApplicationRun
            Exit Function
          End If
          GoTo NormalExecution
        End If

      ElseIf Len(stepS) > 0& Then ' StepS exists but is not a number. This case should be verified beforehand.
        ForNormalStart = False
        RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "startValue and endValue are not numbers: Arg4=[" & stepS & "]", Erl, 4, ExecutingTroughApplicationRun
        Exit Function

      Else ' StepS is empty, so StepL defaults to +1
'        If .startValue > .endValue Then
'          If Not SkipToMatchingNext(tmpLoop) Then
'            ForNormalStart = False
'            RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "For loop is ending, without even to begin, but matching next not found.", Erl, 5, ExecutingTroughApplicationRun
'            Exit Function
'          End If
'          GoTo NormalExecution
'        End If
'        .step = 1&

        .step = IIf(.startValue >= .endValue, 1&, -1&)
      End If

    Else ' startValue and endValue exists but are not numbers. This case should be verified beforehand.
      ForNormalStart = False
      RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "startValue and endValue are not numbers: Arg2=[" & startValueS & "] , Arg3=[" & endValueS & "]", Erl, 6, ExecutingTroughApplicationRun
      Exit Function
    End If
  End With
  If Not PushLoopStack(tmpLoop) Then
    ForNormalStart = False
    RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 7, ExecutingTroughApplicationRun
    Exit Function
  End If

NormalExecution:
  ForNormalStart = True
  Exit Function
eh:
  ForNormalStart = False
  RaiseError MODULE_NAME & ".ForNormalStart", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function



' Prompt for ChatGPT:
'ajuta-ma te rog sa implementez corect comanda ForEachStart, care accepta urmatoarele argumente:
'
'Arg1=ident (este textul care identifica bucla)
'Arg2=empty (aici va fi scrisa adresa celulei curente, in formatul binecunoscut A1, A2, B1, B2 ...)
'Arg3=starting cell (aici va fi introdusa celula de pornire; formatul acceptat este cel binecunoscut (A1, A2, B1, B2...) sau chiar o formula care indica spre o singura celula)
'Arg4=ending cell (aici va fi introdusa celula de final; formatul acceptat este cel binecunoscut (A1, A2, B1, B2...) sau chiar o formula care indica spre o singura celula)
'
'Ca info pe care sa o consideri, tipul LoopInfo l-am extins astfel:
'
'Private Type LoopInfo
'  LoopType     As ForDoLoopsType
'  ident        As String ' User-defined name from Arg1
'  counterValue As Long   ' For loops, row in ForEach loops => aceasta valoare va fi la inceput egala cu startValue
'  startValue   As Long   ' For loops, row in ForEach loops => aceasta valoare va fi linia celulei mentionata in Arg3
'  endValue     As Long   ' For loops, row in ForEach loops => aceasta valoare va fi linia celulei mentionata in Arg4
'  step         As Long   ' For loops, row in ForEach loops => aceasta valoare va fi +1 daca endValue>startValue sau -1 altfel
'  colCounterValue As Long ' col in ForEach loops => aceasta valoare va fi la inceput egala cu colStartValue
'  colStartValue   As Long ' col in ForEach loops => aceasta valoare va fi coloana celulei mentionata in Arg3
'  colEndValue     As Long ' col in ForEach loops => aceasta valoare va fi coloana celulei mentionata in Arg4
'  colStep         As Long ' col in ForEach loops => aceasta valoare va fi +1 daca colEndValue>colStartValue sau -1 altfel
'  counterCell  As Range  ' Cell reference for Arg2 counter (For loops) or condition (Do loops) => in aceasta celula va fi scrisa la inceput adresa celulei mentionate de Arg3
'End Type
'
'Eu adaug aici functia ForNormalStart - presupun ca ar fi super daca ForEachStart ar fi asemanatoare cu ForNormalStart.
'
'Te rog sa te concentrezi deocamdata doar pe functia ForEachStart, adaugand comentarii in cod pe limba engleza.

Public Function ForEachStart(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional), Arg2 = empty (counter), Arg3 = start cell, Arg4 = end cell
  On Error GoTo eh
  Dim tmpLoop As LoopInfo
  Dim startCell As Range, endCell As Range
  
  ' Resolve start and end cells, including handling formulas
  Set startCell = GetReferencedCell(currentRowRange(1&, ColAArg1 + 2&))
  Set endCell = GetReferencedCell(currentRowRange(1&, ColAArg1 + 3&))
  
  ' Validate cell references
  If startCell Is Nothing Or endCell Is Nothing Then
  ' If either start or end cell is invalid, skip the loop
    Call SkipToMatchingNext(tmpLoop)
    GoTo FaultExecution
  End If
  
  ' Initialize loop info
  With tmpLoop
    .ident = CleanString(CStr(currentRowArray(1&, ColAArg1))) ' Store loop name
    .LoopType = LoopForEach
    Set .counterCell = currentRowRange(1&, ColAArg1 + 1&) ' Set counter cell reference
    
    ' Extract row and column indices
    .startValue = startCell.Row
    .endValue = endCell.Row
    .colStartValue = startCell.Column
    .colEndValue = endCell.Column
    
    ' Set initial counter values
    .counterValue = .startValue
    .colCounterValue = .colStartValue
    
    ' Determine step values based on start and end position
    If .endValue > .startValue Then
      .step = 1
    ElseIf .endValue < .startValue Then
      .step = -1
    Else
      .step = 0 ' No movement in rows
    End If
    
    If .colEndValue > .colStartValue Then
      .colStep = 1
    ElseIf .colEndValue < .colStartValue Then
      .colStep = -1
    Else
      .colStep = 0 ' No movement in columns
    End If
    
    ' If both step values are zero, the loop has no iteration range, so skip it
    If .step = 0 And .colStep = 0 Then
      Call SkipToMatchingNext(tmpLoop)
      GoTo FaultExecution
    End If
    
    ' Write the initial counter cell address
    .counterCell.Value = startCell.Address(False, False) ' Store as "A1" format
  End With
  
  ' Push loop onto stack for tracking
  If Not PushLoopStack(tmpLoop) Then
    ForEachStart = False
    RaiseError MODULE_NAME & ".ForEachStart", Err.Number, Err.Source, "To many loops in the queue: " & loopStackSize, Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

NormalExecution:
  ForEachStart = True
  Exit Function
FaultExecution:
  ForEachStart = False
  Exit Function
eh:
  ForEachStart = False
  RaiseError MODULE_NAME & ".ForEachStart", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function





Public Function ForNext(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
' Arg1 = ident (optional)
  If loopStackSize = 0& Then GoTo FaultExecution


  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                           ' we have an ident of a For to be Next-ed
    For tmpL = loopStackSize To 1& Step -1&
      If ((loopStack(tmpL).LoopType = LoopFor) Or _
          (loopStack(tmpL).LoopType = LoopForEach)) And _
           loopStack(tmpL).ident = ident Then _
             Exit For                               ' found the For with the same ident
    Next tmpL
  Else                                              ' no ident of a For to be Next-ed provided; then let the last For be
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopFor) Or _
         (loopStack(tmpL).LoopType = LoopForEach) Then _
           Exit For                                 ' found the last For
    Next tmpL
  End If


  If tmpL = 0& Then GoTo FaultExecution             ' no For loop found

  Do While loopStackSize > tmpL: PopLoopStack: Loop ' remove any preceding loops


  With loopStack(tmpL)


    If .LoopType = LoopFor Then                  ' Handling LoopFor (Numeric For Loops)

      .counterValue = .counterValue + .step      ' update counter
      .counterCell.Value = .counterValue         ' update cell value

      ' check loop condition
      If (.step > 0 And .counterValue <= .endValue) Or _
         (.step < 0 And .counterValue >= .endValue) Then
           currentRow = .counterCell.Row         ' return to loop start
      Else
        PopLoopStack                             ' remove loop from stack
      End If



    ElseIf .LoopType = LoopForEach Then          ' Handling LoopForEach (Iteration Over Cells)
      Dim NextRow As Long, nextCol As Long

      ' Determine next row and column based on iteration order
      If startWithRows Then                      ' First iterate over rows, then move to the next column
        NextRow = .counterValue + .step
        nextCol = .colCounterValue
      Else                                       ' First iterate over columns, then move to the next row
        NextRow = .counterValue
        nextCol = .colCounterValue + .colStep
      End If

      ' Check if row iteration should continue
      If (.step > 0 And NextRow <= .endValue) Or _
         (.step < 0 And NextRow >= .endValue) Then ' Update row counter

        .counterValue = NextRow
        .counterCell.Value = Cells(NextRow, nextCol).Address(False, False) ' Store address in A1 format
      
      ' If rows are finished, move to the next column (if startWithRows is True)
      ElseIf startWithRows Then                    ' Move to the next column
          
          nextCol = .colCounterValue + .colStep

          ' Check if column iteration should continue
          If (.colStep > 0 And nextCol <= .colEndValue) Or _
             (.colStep < 0 And nextCol >= .colStartValue) Then
            
            ' Reset row counter and increment column counter
            .counterValue = .startValue
            .colCounterValue = nextCol
            .counterCell.Value = Cells(.startValue, nextCol).Address(False, False)
          
          Else                                     ' If no more rows and no more columns, exit loop
            PopLoopStack
            GoTo NormalExecution
          End If

      ' If columns are finished, move to the next row (if startWithRows is False)
      Else                                         ' Move to the next row

        NextRow = .counterValue + .step

        ' Check if row iteration should continue
        If (.step > 0 And NextRow <= .endValue) Or _
           (.step < 0 And NextRow >= .startValue) Then

          ' Reset column counter and increment row counter
          .colCounterValue = .colStartValue
          .counterValue = NextRow
          .counterCell.Value = Cells(NextRow, .colStartValue).Address(False, False)

        Else                                       ' If no more columns and no more rows, exit loop
          PopLoopStack
          GoTo NormalExecution
        End If
      End If
  
      currentRow = .counterCell.Row                ' Return to loop start
  
  
    End If

  End With
  
NormalExecution:
  ForNext = True
  Exit Function
FaultExecution:
  ForNext = False
  Exit Function
eh:
  ForNext = False
  RaiseError MODULE_NAME & ".ForNext", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function






Public Function ExitFor(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = ident (optional)
  On Error GoTo eh
  If loopStackSize = 0& Then GoTo FaultExecution
  
  Dim ident As String: ident = CleanString(CStr(currentRowArray(1&, ColAArg1)))
  If Len(ident) > 0& Then                               ' we have an ident of a For to be Exit-ed
    For tmpL = loopStackSize To 1& Step -1&
      If ((loopStack(tmpL).LoopType = LoopFor) Or _
          (loopStack(tmpL).LoopType = LoopForEach)) And _
           loopStack(tmpL).ident = ident Then _
             Exit For                                   ' found the For with the same ident
    Next tmpL
  Else                                                  ' no ident of a For to be Exit-ed provided; then let the last For be
    For tmpL = loopStackSize To 1& Step -1&
      If (loopStack(tmpL).LoopType = LoopFor) Or _
         (loopStack(tmpL).LoopType = LoopForEach) Then _
           Exit For                                     ' found the last For
    Next tmpL
  End If

  If tmpL = 0& Then GoTo FaultExecution                 ' no For loop found

  If Not SkipToMatchingNext(loopStack(tmpL)) Then _
    GoTo FaultExecution                                 ' if no Next found, then we cannot Exit

  Do While loopStackSize >= tmpL: PopLoopStack: Loop    ' remove any preceding loops along with the found loop

NormalExecution:
  ExitFor = True
  Exit Function
FaultExecution:
  ExitFor = False
  Exit Function
eh:
  ExitFor = False
  RaiseError MODULE_NAME & ".ExitFor", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function






Private Function PushLoopStack(ByRef newLoop As LoopInfo) As Boolean
  On Error GoTo eh
  PushLoopStack = True
  If loopStackCapacity = 0& Then ' Initialize stack if necessary
    ReDim loopStack(1 To loopStackBlock)
    loopStackCapacity = loopStackBlock
  ElseIf loopStackSize >= loopStackCapacity Then ' Expand array if needed
    loopStackCapacity = loopStackCapacity + loopStackBlock
    ReDim Preserve loopStack(1& To loopStackCapacity)

    If loopStackSize > loopStackMaxSize Then
      ' Too many nested loop - warning on "out of memory" error - perhaps is a mistake - let user decide
      If AskNextStep("WARNING:" & vbCrLf & vbCrLf & _
                     "There are " & loopStackSize & "nested loops." & vbCrLf & vbCrLf & _
                     "Do you want to continue execution and receive a new warning when there will be " & loopStackSize + loopStackBlock & " nested loops?", _
                     vbYesNo, "Warning on too many nested loops") = vbNo Then _
                         PushLoopStack = False
    End If
  End If

  ' Add new loop to stack
  loopStackSize = loopStackSize + 1&
  loopStack(loopStackSize) = newLoop

done:
  Exit Function
eh:
  PushLoopStack = False
  RaiseError MODULE_NAME & ".PushLoopStack", Err.Number, Err.Source, Err.Description, Erl
End Function

Private Function PopLoopStack() As LoopInfo
  On Error GoTo eh
  If loopStackSize = 0& Then ' Stack is empty, so exit
    PopLoopStack.LoopType = NoLoop
    Exit Function
  End If

  ' Return last stored loop
  PopLoopStack = loopStack(loopStackSize)
  loopStackSize = loopStackSize - 1&

  ' Shrink array if necessary (optional)
  If loopStackSize < loopStackCapacity - loopStackBlock Then
    loopStackCapacity = loopStackCapacity - loopStackBlock
    If loopStackCapacity Then
      ReDim Preserve loopStack(1& To loopStackCapacity)
    Else
      Erase loopStack
    End If
  End If

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PopLoopStack", Err.Number, Err.Source, Err.Description, Erl
End Function

Private Function SkipToMatchingNext(ByRef currentLoop As LoopInfo) As Boolean
' return false if no Next found
' https://forum.ozgrid.com/forum/index.php?thread/96677-two-column-combination-search/&postID=735616#post735616
' Dim currentLoop As LoopInfo: currentLoop = loopStack(loopStackSize)
  On Error GoTo eh
  With currentLoop
    Dim ident As String: ident = .ident
    Dim firstFound As Long
    Select Case .LoopType
      Case LoopFor, LoopForEach
        Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeNext, After:=currentRowRange(1&, ColAArg1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If tmpR Is Nothing Then GoTo FaultExecution
        If tmpR.Row < currentLoop.counterCell.Row Then GoTo FaultExecution
      
        If Len(ident) = 0& Then GoTo NormalExecution
      
        firstFound = tmpR.Row
        Do
          If CleanString(tmpR.Offset(0, 1).Value) = ident Then GoTo NormalExecution
      
          Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeNext, After:=tmpR, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
          If tmpR Is Nothing Then GoTo FaultExecution
          If tmpR.Row < firstFound Then GoTo FaultExecution
      
        Loop
      Case DoInfinite, DoWhile, DoUntil, LoopWhile, LoopUntil, LoopInfinite
        Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeLoop, After:=currentRowRange(1&, ColAArg1), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
        If tmpR Is Nothing Then GoTo FaultExecution
        If tmpR.Row < currentLoop.counterCell.Row Then GoTo FaultExecution

        If Len(ident) = 0& Then GoTo NormalExecution

        firstFound = tmpR.Row
        Do
          If CleanString(tmpR.Offset(0, 1).Value) = ident Then GoTo NormalExecution
      
          Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeLoop, After:=tmpR, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
          If tmpR Is Nothing Then GoTo FaultExecution
          If tmpR.Row < firstFound Then GoTo FaultExecution

        Loop
    End Select
  End With

FaultExecution:
  SkipToMatchingNext = False
  Exit Function
NormalExecution:
  currentRow = tmpR.Row
  SkipToMatchingNext = True
  Exit Function
eh:
  SkipToMatchingNext = False
  RaiseError MODULE_NAME & ".SkipToMatchingNext", Err.Number, Err.Source, Err.Description, Erl
End Function

Private Function GetReferencedCell(ByRef targetCell As Range) As Range
  On Error GoTo eh
  If targetCell Is Nothing Then _
    RaiseError MODULE_NAME & ".GetReferencedCell", Err.Number, Err.Source, "The targetCell is Nothing and is required.", Erl, 1

  ' If the cell exists, check if it has a formula
  If targetCell.HasFormula Then
    ' Extract the referenced address from the formula
    Dim referencedAddress As String
    referencedAddress = Replace(targetCell.Formula, "=", "") ' Remove "="
    
    ' Check if it's a valid single-cell reference
    On Error Resume Next
    Set GetReferencedCell = shAuto.Range(referencedAddress)
    On Error GoTo eh
    
    ' If the extracted reference is invalid, return the cell from content
    If GetReferencedCell Is Nothing Then
      ' Attempt to get the cell from content
      On Error Resume Next
      Set GetReferencedCell = shAuto.Range(targetCell.Value2)
      On Error GoTo eh
      
      If GetReferencedCell Is Nothing Then _
        RaiseError MODULE_NAME & ".GetReferencedCell", Err.Number, Err.Source, _
          "The cell could not be identified from the formula : targetCell.Formula=[" & CStr(targetCell.Formula) & "] or from value targetCell.Value2=[" & CStr(targetCell.Value2) & "]", Erl, 2

    End If
  Else
    ' Attempt to get the cell from content
    On Error Resume Next
    Set GetReferencedCell = shAuto.Range(targetCell.Value2)
    On Error GoTo eh
    
    If GetReferencedCell Is Nothing Then _
      RaiseError MODULE_NAME & ".GetReferencedCell", Err.Number, Err.Source, "The cell could not be identified from value: targetCell.Value2=[" & CStr(targetCell.Value2) & "]", Erl, 3
  
  End If

done:
  Exit Function
eh:
  Set GetReferencedCell = Nothing
  RaiseError MODULE_NAME & ".GetReferencedCell", Err.Number, Err.Source, Err.Description, Erl
End Function


Attribute VB_Name = "CommandsCondition"
Option Explicit
Private Const MODULE_NAME As String = "CommandsCondition"
Private tmpL As Long, tmpS As String, tmpR As Range

Public Sub RegisterCommandsCondition()
  On Error GoTo eh
  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "stop", Array("CommandStop", "Stop", MODULE_NAME, "Stops the execution.")

  commandMap.Add "end", Array("CommandStop", "End", MODULE_NAME, "Ends the execution.")

  commandMap.Add "skip", Array("CommandSkip", "Skip", MODULE_NAME, "How many further lines not to be executed. Current line excluded. The line to be executed is excluded.", _
    "Number of Lines", "How many lines to be skipped. If currentRow is 2 and command is Skip 3, next command to execute is on line 6.")

  commandMap.Add "ifthenskip", Array("CommandIfThenSkip", "If Then Skip", MODULE_NAME, "Skip lines according to a condition.", _
    "Condition", "Condition that should return true or false.", _
    "Skip Lines if True", "How many lines to be skipped if condition is true.", _
    "Skip Lines if False", "How many lines to be skipped if condition is false.")
  
  commandMap.Add "ifthengoto", Array("CommandIfThenGoTo", "If Then GoTo", MODULE_NAME, "Set next line or label to be executed according to a condition.", _
    "Condition", "Condition that should return true or false.", _
    "Line if True", "What line or label should be next if condition is true." & loopListLabel, _
    "Line if False", "What line or label should be next if condition is false." & loopListLabel)

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsCondition", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsCondition()
  On Error GoTo eh


done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsCondition", Err.Number, Err.Source, Err.description, Erl
End Sub


Public Function CommandStop(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  stopExecutionRequired = True
  CommandStop = True
End Function

Public Function CommandSkip(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: How many lines to be skipped. If currentRow is 2 and command is Skip 3, next command to execute is on line 6.
  On Error GoTo eh
  Call SkipLines(CStr(currentRowArray(1, ColAArg1)), "Arg1")

done:
  CommandSkip = True
  Exit Function
eh:
  CommandSkip = False
  RaiseError MODULE_NAME & ".CommandSkip", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function CommandIfThenSkip(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: Condition that should return true or false
' Arg2: How many lines to be skipped if condition is true
' Arg3: How many lines to be skipped if condition is false
  On Error GoTo eh
  If GetBoolean(CStr(currentRowArray(1, ColAArg1))) Then
    SkipLines CStr(currentRowArray(1, ColAArg1 + 1)), "Arg2"

  Else ' Arg1 is false
    SkipLines CStr(currentRowArray(1, ColAArg1 + 2)), "Arg3"

  End If

NormalExecution:
  CommandIfThenSkip = True
  Exit Function
eh:
  CommandIfThenSkip = False
  RaiseError MODULE_NAME & ".CommandIfThenSkip", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function CommandIfThenGoTo(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: Condition that should return true or false
' Arg2: What line or label should be next if condition is true
' Arg3: What line or label should be next if condition is false
  On Error GoTo eh
  If GetBoolean(CStr(currentRowArray(1, ColAArg1))) Then
    GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 1)), "Arg2"

  Else ' Arg1 is false
    GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 2)), "Arg3"

  End If

NormalExecution:
  CommandIfThenGoTo = True
  Exit Function
eh:
  CommandIfThenGoTo = False
  RaiseError MODULE_NAME & ".CommandIfThenGoTo", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function



Public Function SkipLines(Lines As String, ArgName As String, Optional EmptyIsError As Boolean = False) As Boolean
  On Error GoTo eh
  If Len(Lines) = 0 Then
    If EmptyIsError Then
      SkipLines = False
      RaiseError MODULE_NAME & ".SkipLines", Err.Number, Err.Source, ArgName & " must contain a line number and is empty.", Erl, 1
      Exit Function
    End If
    GoTo NormalExecution

  ElseIf IsNumber(Lines) Then
    currentRow = currentRow + CLng(Lines)
    GoTo NormalExecution

  Else
    SkipLines = False
    RaiseError MODULE_NAME & "SkipLines", Err.Number, Err.Source, _
      "Number of lines to skip needs to be valid number: " & ArgName & "=[" & Lines & "]", Erl, 2
    Exit Function
  End If

NormalExecution:
  SkipLines = True
  Exit Function
eh:
  SkipLines = False
  RaiseError MODULE_NAME & ".SkipLines", Err.Number, Err.Source, Err.description, Erl
End Function

Public Function GotoLineOrLabel(LineOrLabel As String, ArgName As String, Optional EmptyIsError As Boolean = False) As Boolean
  On Error GoTo eh
  If Len(LineOrLabel) = 0 Then
    If EmptyIsError Then
      GotoLineOrLabel = False
      RaiseError MODULE_NAME & ".GotoLineOrLabel", Err.Number, Err.Source, ArgName & " must contain a line number or a label and is empty.", Erl, 1
      Exit Function
    End If
    GoTo NormalExecution

  ElseIf IsNumber(LineOrLabel) Then
    currentRow = CLng(LineOrLabel) - 1&
    GoTo NormalExecution

  Else ' LineOrLabel is not a number, but a text
    Set tmpR = shAuto.Columns(ColACommand).Find(loopTypeLabel, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

    If Not tmpR Is Nothing Then

      tmpL = tmpR.Row
      Do
        If StrComp(CleanString(tmpR.Offset(0, 1).Text), LineOrLabel, vbTextCompare) = 0 Then
          currentRow = tmpR.Row
          GoTo NormalExecution
        End If
        Set tmpR = shAuto.Columns(ColACommand).FindNext(tmpR)
      Loop Until tmpR.Row = tmpL
    End If

    GotoLineOrLabel = False
    RaiseError MODULE_NAME & ".GotoLineOrLabel", Err.Number, Err.Source, "No Label with the ident found: " & ArgName & "=[" & LineOrLabel & "]", Erl, 2
    Exit Function
  End If

NormalExecution:
  GotoLineOrLabel = True
  Exit Function
eh:
  GotoLineOrLabel = False
  RaiseError MODULE_NAME & ".GotoLineOrLabel", Err.Number, Err.Source, Err.description, Erl
End Function

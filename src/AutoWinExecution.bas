Attribute VB_Name = "AutoWinExecution"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinExecution"

Public Sub ExecuteAutomation()
  On Error GoTo eh
  Call ufAutoWin.Show(vbModal) ' Get Sub to start execution - wait to close the window
  If ufAutoWin.SelectedLine < 1 Then GoTo done Else currentRow = ufAutoWin.SelectedLine
  
  Do Until stopExecutionRequired
    Call CenterViewToCurrentRow
    Call SaveCurrentRowValues

    Static command As String: command = CleanString(CStr(currentRowArray(1, ColACommand)))
    If commandMap.Exists(command) Then
      emptyRows = 0
      StatusUpdate StatusNOW
      If Not HandleBeforeChecks Then GoTo StopExecution
      Static cmdInfo As Variant
      cmdInfo = commandMap(command)

' https://stackoverflow.com/questions/42124252/application-run-with-error-trapping
' http://www.cpearson.com/excel/errorhandling.htm
      'Call RecordFocusedControlState(cmdInfo(cmdFunctionName))
      If Application.Run(cmdInfo(cmdFunctionName), True) Then
        StatusUpdate StatusOK
      Else
        If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
        StatusUpdate StatusNOK
        stopExecutionRequired = True
      End If

    Else ' commandMap.exists(command)
      If tooManyEmptyRows Then GoTo StopExecution
      StatusUpdate StatusSKIP
      
      If Not HandleBeforeChecks Then GoTo StopExecution
    End If

    currentRow = currentRow + 1
  Loop

StopExecution:
  If currentRow > 0 Then If Left(currentRowArray(1, ColAStatus), 1) = StatusNOW Then StatusUpdate StatusNOK

done:
  Application.StatusBar = False
  Exit Sub
eh:
  Application.StatusBar = False
  RaiseError MODULE_NAME & ".ExecuteAutomation", Err.Number, Err.Source, Err.description, Erl
End Sub


Private Function HandleBeforeChecks() As Boolean
110  On Error GoTo eh
120  If Not WaitWindowToActivate Then
130    HandleBeforeChecks = False
140    RaiseError MODULE_NAME & ".HandleBeforeChecks", Err.Number, Err.Source, "No window with the mentioned name found => execution will stop.", Erl, 1
150    Exit Function
160  End If
165  'Err.Raise 2000
170  If Not WaitColorUnderCursor Then
180    HandleBeforeChecks = False
190    RaiseError MODULE_NAME & ".HandleBeforeChecks", Err.Number, Err.Source, "The color under cursor is not the same as mentioned => execution will stop.", Erl, 2
200    Exit Function
210  End If

220  Static sleepTime As Long
230  If IsNumber(CStr(currentRowArray(1, ColAPause))) Then sleepTime = maxLong(CLng(currentRowArray(1, ColAPause)), minWaitTime) Else sleepTime = minWaitTime
240  shAuto.Calculate
250  Application.ScreenUpdating = False
260  Application.ScreenUpdating = True
270  'ActiveSheet.EnableCalculation = False
280  'ActiveSheet.EnableCalculation = True
290  If MySleep(sleepTime) Then
300    ' Mouse moved
310    HandleBeforeChecks = False
320    RaiseError MODULE_NAME & ".HandleBeforeChecks", Err.Number, Err.Source, "Mouse was moved => execution will stop.", Erl, 3
330    Exit Function
340  End If

350  HandleBeforeChecks = True

done:
360  Exit Function
eh:
370  HandleBeforeChecks = False
380  RaiseError MODULE_NAME & ".HandleBeforeChecks", Err.Number, Err.Source, Err.description, Erl
End Function

Private Function tooManyEmptyRows() As Boolean
  On Error GoTo eh
  emptyRows = emptyRows + 1: tooManyEmptyRows = emptyRows > maxEmptyRows

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".tooManyEmptyRows", Err.Number, Err.Source, Err.description, Erl
End Function

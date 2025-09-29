Attribute VB_Name = "AutoWinValidation"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinValidation"

Public Function ValidateAutomationSheet() As Boolean
  On Error GoTo eh
  ' Checking Caption
  Set currentRowRange = shAuto.Cells(1, ColAStatus).Resize(1, ColAComment - ColAStatus + 1)
  currentRowArray = currentRowRange.Value2
  
  ValidateAutomationSheet = _
    (currentRowArray(1, ColAStatus) = "Status") And _
    (currentRowArray(1, ColACommand) = "Command") And _
    (currentRowArray(1, ColAArg1 + 0) = "Arg1") And _
    (currentRowArray(1, ColAArg1 + 1) = "Arg2") And _
    (currentRowArray(1, ColAArg1 + 2) = "Arg3") And _
    (currentRowArray(1, ColAArg1 + 3) = "Arg4") And _
    (currentRowArray(1, ColAArg1 + 4) = "Arg5") And _
    (currentRowArray(1, ColAArg1 + 5) = "Arg6") And _
    (currentRowArray(1, ColAArg1 + 6) = "Arg7") And _
    (currentRowArray(1, ColAArg1 + 7) = "Arg8") And _
    (currentRowArray(1, ColAArg1 + 8) = "Arg9") And _
    (currentRowArray(1, ColAArg1 + 9) = "Arg10") And _
    (currentRowArray(1, ColAWindow) = "WindowName before") And _
    (currentRowArray(1, ColAColor) = "ColorUnderMouse before") And _
    (currentRowArray(1, ColAPause) = "Pause before") And _
    (currentRowArray(1, ColAKeybd) = "KeybdCode") And _
    (currentRowArray(1, ColAonError) = "On Error") And _
    (currentRowArray(1, ColAComment) = "Comment")

  If Not ValidateAutomationSheet Then errorDescription = "Heading on Automation sheet is not complete."

done:
  Exit Function
eh:
  ValidateAutomationSheet = False
  RaiseError MODULE_NAME & ".ValidateAutomationSheet", Err.Number, Err.Source, Err.description, Erl
End Function


Attribute VB_Name = "AutoWinMain"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinMain"

Public Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "M\n14"
  On Error GoTo eh
  Call InitializeSettings

  If Not ValidateAutomationSheet() Then
    MsgBox "Error: Automation values are incomplete or invalid!" & vbCrLf & errorDescription, vbCritical
    GoTo done
  End If

  Call ClearStatusColumn

  Call ExecuteAutomation
  
  Call PrepareExit

done:
  Exit Sub
eh:
  DisplayError MODULE_NAME & ".Main", Err.Source, Err.description, Erl
End Sub

Public Sub ShowCommands()
Attribute ShowCommands.VB_ProcData.VB_Invoke_Func = "N\n14"
  On Error GoTo eh
  Call InitializeSettings

  Call ufCommand.Show(vbModal)

done:
  Exit Sub
eh:
  DisplayError MODULE_NAME & ".ShowCommands", Err.Source, Err.description, Erl
End Sub

Public Sub MoveMouseToXY()
Attribute MoveMouseToXY.VB_ProcData.VB_Invoke_Func = "Y\n14"
  On Error GoTo eh
  Call InitializeSettings

  currentRow = ActiveCell.Row
  Call SaveCurrentRowValues
  Call GetColorFromPoint
  Call MoveMouse

done:
  Exit Sub
eh:
  DisplayError MODULE_NAME & ".MoveMouseToXY", Err.Source, Err.description, Erl
End Sub

Public Sub ReadMouseToXY()
Attribute ReadMouseToXY.VB_ProcData.VB_Invoke_Func = "X\n14"
  On Error GoTo eh
  Call InitializeSettings

  currentRow = ActiveCell.Row
  Call SaveCurrentRowValues
  Call ReadMouseXY
  Call GetColorFromPoint
  Call RecordControlStateUndeMouse("MoveMouse")

done:
  Exit Sub
eh:
  DisplayError MODULE_NAME & ".MoveMouseToXY", Err.Source, Err.description, Erl
End Sub


Attribute VB_Name = "CommandsExecute"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsExecution"


Public Function RegisterCommandsExecute()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "runprogram", Array("RunProgram", "Run Program", _
    MODULE_NAME, "Start an application", _
    "Application", "ca be 'notepad', 'iexplorer' or 'C:\Program Files\Notepad++\notepad++.exe'", _
    "Argument", "Here specify the argument for the program (can be '/G')")

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsExecute", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsExecute()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsExecute", Err.Number, Err.Source, Err.Description, Erl
End Function


Public Function RunProgram(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
' https://stackoverflow.com/questions/20917355/how-do-you-run-a-exe-with-parameters-using-vbas-shell
' Arg1 is the program (ca be "notepad", "iexplorer" or "C:\Program Files\Notepad++\notepad++.exe")
' Arg2 is the argument for the program (can be "/G")
  If Len(currentRowArray(1, ColAArg1 + 1)) > 0 Then
    Shell """" & currentRowArray(1, ColAArg1) & """ """ & currentRowArray(1, ColAArg1 + 1) & """", vbNormalFocus
  Else
    Shell """" & currentRowArray(1, ColAArg1) & """", vbNormalFocus
  End If
done:
  RunProgram = True
  Exit Function
eh:
  RunProgram = False
  RaiseError MODULE_NAME & ".RunProgram", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

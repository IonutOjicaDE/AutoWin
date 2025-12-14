Attribute VB_Name = "CommandsMouse"
Option Explicit
Private Const MODULE_NAME As String = "CommandsMouse"

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As LongPtr)


Private Enum MOUSEEVENTF_Constants 'mouse_event; dwFlags ; https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mouse_event
  MOUSEEVENTF_LEFTDOWN = &H2&
  MOUSEEVENTF_LEFTUP = &H4&
  MOUSEEVENTF_RIGHTDOWN = &H8&
  MOUSEEVENTF_RIGHTUP = &H10&
  MOUSEEVENTF_MIDDLEDOWN = &H20&
  MOUSEEVENTF_MIDDLEUP = &H40&
End Enum


Private StopByMouseMove  As Boolean
Private PauseByMouseMove As Byte '0=not paused, 1=paused with StopByMouseMove=true before, 2=paused with StopByMouseMove=false before

Private MouseNow As POINTAPI
Private MousePos As POINTAPI

Public Function RegisterCommandsMouse()
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "mousemovedetection", Array("MoveDetection", "Mouse Move Detection", _
    MODULE_NAME, "", _
    "State", "{{Start / Stop / Pause / Resume}}")
  commandMap.Add "readmousexy", Array("ReadMouseXY", "Read Mouse XY", _
    MODULE_NAME, "Read mouse position in X and Y", _
    "X-Value", "X Value of the current mouse position", _
    "Y-Value", "Y Value of the current mouse position")
  commandMap.Add "movemouse", Array("MoveMouse", "Move Mouse", _
    MODULE_NAME, "Move mouse to position X, Y", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")

  commandMap.Add "leftclick", Array("LeftClick", "Left Click", _
    MODULE_NAME, "Move mouse to position X, Y then simulate a mouse left click", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "leftdoubleclick", Array("LeftDoubleClick", "Left Double Click", _
    MODULE_NAME, "Move mouse to position X, Y then simulate a mouse left double click", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "rightclick", Array("RightClick", "Right Click", _
    MODULE_NAME, "Move mouse to position X, Y then simulate a mouse right click", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "middleclick", Array("MiddleClick", "Middle Click", _
    MODULE_NAME, "Move mouse to position X, Y then simulate a mouse middle click", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")

  commandMap.Add "leftclickdown", Array("LeftClickDown", "Left Click Down", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the pressing holding down mouse left button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "leftclickup", Array("LeftClickUp", "Left Click Up", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the releasing mouse left button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "rightclickdown", Array("RightClickDown", "Right Click Down", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the pressing holding down mouse right button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "rightclickup", Array("RightClickUp", "Right Click Up", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the releasing mouse right button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "middleclickdown", Array("MiddleClickDown", "Middle Click Down", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the pressing holding down mouse middle button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")
  commandMap.Add "middleclickup", Array("MiddleClickUp", "Middle Click Up", _
    MODULE_NAME, "Move mouse to position X, Y then simulate the releasing mouse middle button", _
    "X-Value", "X Value to move the mouse", _
    "Y-Value", "Y Value to move the mouse")

'MoveDetectionStart
  StopByMouseMove = True
  PauseByMouseMove = 0
  GetCursorPos MousePos

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsMouse", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsMouse()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsMouse", Err.Number, Err.Source, Err.Description, Erl
End Function



Public Function MoveDetection(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveDetection = True
  Select Case CleanString(CStr(currentRowArray(1, ColAArg1)))
    Case "start":
      StopByMouseMove = True
      PauseByMouseMove = 0
      GetCursorPos MousePos
    Case "stop":
      StopMoveDetection
    Case "pause":
      PauseByMouseMove = IIf(StopByMouseMove, 1, 2)
      StopByMouseMove = False
    Case "resume":
      StopByMouseMove = PauseByMouseMove = 1
      PauseByMouseMove = 0
      If StopByMouseMove Then GetCursorPos MousePos
    Case Else:
      MoveDetection = False
  End Select
done:
  Exit Function
eh:
  MoveDetection = False
  RaiseError MODULE_NAME & ".MoveDetection", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Sub StopMoveDetection(): StopByMouseMove = False: PauseByMouseMove = 0: End Sub


Public Function ReadMouseXY(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  GetCursorPos MousePos
  currentRowRange(1, ColAArg1 + 0).Value = MousePos.x: currentRowRange(1, ColAArg1 + 1).Value = MousePos.y
done:
  ReadMouseXY = True
  Exit Function
eh:
  ReadMouseXY = False
  RaiseError MODULE_NAME & ".ReadMouseXY", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function MoveMouse(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
'https://stackoverflow.com/questions/13896658/sendinput-vb-basic-example
'SetCursorPos => is deprecated; SendInput should be used instead
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Or IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then
    GetCursorPos MousePos
    With MousePos
      If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then .x = CLng(currentRowArray(1, ColAArg1 + 0))
      If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then .y = CLng(currentRowArray(1, ColAArg1 + 1))
      If shakeMovementDist Then
        SetCursorPos .x + shakeMovementDist, .y + shakeMovementDist
        Sleep shakeMovementWait
      End If
      SetCursorPos .x, .y
      Sleep maxLong(waitAfterMove + IIf(shakeMovementDist, -shakeMovementWait, 0&), 1&)
    End With
    MoveMouse = True
  Else
    MoveMouse = False
  End If
done:
  'Call RecordControlStateUndeMouse("MoveMouse") ' uncomment this line to have more informations about windows <<<<<<<<<<<<<<<
  Exit Function
eh:
  MoveMouse = False
  RaiseError MODULE_NAME & ".MoveMouse", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function LeftClick(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
done:
  LeftClick = True
  Exit Function
eh:
  LeftClick = False
  RaiseError MODULE_NAME & ".LeftClick", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function LeftDoubleClick(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
done:
  LeftDoubleClick = True
  Exit Function
eh:
  LeftDoubleClick = False
  RaiseError MODULE_NAME & ".LeftDoubleClick", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function RightClick(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
done:
  RightClick = True
  Exit Function
eh:
  RightClick = False
  RaiseError MODULE_NAME & ".RightClick", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function MiddleClick(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
  Sleep waitAfterClick
  mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
done:
  MiddleClick = True
  Exit Function
eh:
  MiddleClick = False
  RaiseError MODULE_NAME & ".MiddleClick", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function LeftClickDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
done:
  LeftClickDown = True
  Exit Function
eh:
  LeftClickDown = False
  RaiseError MODULE_NAME & ".LeftClickDown", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function LeftClickUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
done:
  LeftClickUp = True
  Exit Function
eh:
  LeftClickUp = False
  RaiseError MODULE_NAME & ".LeftClickUp", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function RightClickDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
done:
  RightClickDown = True
  Exit Function
eh:
  RightClickDown = False
  RaiseError MODULE_NAME & ".RightClickDown", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function RightClickUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
done:
  RightClickUp = True
  Exit Function
eh:
  RightClickUp = False
  RaiseError MODULE_NAME & ".RightClickUp", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function MiddleClickDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0
done:
  MiddleClickDown = True
  Exit Function
eh:
  MiddleClickDown = False
  RaiseError MODULE_NAME & ".MiddleClickDown", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function MiddleClickUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MoveMouse
  mouse_event MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0
done:
  MiddleClickUp = True
  Exit Function
eh:
  MiddleClickUp = False
  RaiseError MODULE_NAME & ".MiddleClickUp", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function



Public Function MouseMovedByUser(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
'other posibilities to allow User to stop AutoWin:
'https://www.informit.com/articles/article.aspx?p=366892&seqNum=4

  GetCursorPos MouseNow
  If StopByMouseMove Then
'option to be added: save active window, then activate Excel window
'if user wish to continue, then revert the previous active window, wait 200ms then continue
    If Not AreSamePointsLong(MouseNow.x, MouseNow.y, MousePos.x, MousePos.y, positionTolerance) Then _
      MouseMovedByUser = _
        AskNextStep("You have moved the mouse." & vbCrLf & vbCrLf & "Do you want to STOP the running of the automation?", vbYesNo, "Stop the automation?") = vbYes
  Else
    MouseMovedByUser = False
  End If
  Sleep minWaitTime
  GetCursorPos MouseNow
  MousePos.x = MouseNow.x: MousePos.y = MouseNow.y
done:
  Exit Function
eh:
  MouseMovedByUser = False
  RaiseError MODULE_NAME & ".MouseMovedByUser", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

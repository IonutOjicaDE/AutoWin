Attribute VB_Name = "AutoWinRecord"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinRecord"

Public Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
Public Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As LongPtr, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As LongPtr) As LongPtr
Public Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As LongPtr, ByVal wParam As LongPtr, lParam As Any) As LongPtr

Public MouseHookHandle As LongPtr, KeyHookHandle As LongPtr


'https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/index
'https://docs.microsoft.com/de-de/windows/win32/api/winuser/ns-winuser-kbdllhookstruct?redirectedfrom=MSDN
Private Type KBDLLHOOKSTRUCT ' {
  VKcode      As Long ' virtual key code in range 1 .. 254
  ScanCode    As Long ' hardware code
  flags       As Long ' bit 4: alt key was pressed
  time        As Long ' in ms since computer was started
  dwExtraInfo As Long
End Type  ' }

'******************************************************
Private Enum WM_Constants
'* System messages for Keys that we want to trace
  WM_KEYFIRST = &H100& ' 256
  WM_KEYDOWN = &H100&  ' 256
  WM_KEYUP = &H101&    ' 257
  WM_CHAR = &H102& 'https://docs.microsoft.com/en-us/windows/win32/inputdev/using-keyboard-input
  WM_DEADCHAR = &H103&
  WM_SYSKEYDOWN = &H104&
  WM_SYSKEYUP = &H105&
  WM_SYSCHAR = &H106&
  WM_SYSDEADCHAR = &H107&
  WM_KEYLAST = &H108&
  
'* System messages for Mouse that we want to trace
  WM_MOUSEMOVE = &H200
  WM_MOUSEWHEEL = &H20A
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
End Enum
'******************************************************

Private Const HC_ACTION = 0


'±3 pixels position diference will be considered as no move of mouse
Private Const MousePositionTolerance As Long = 3


'* Type to hold Mouse Hook information
Private Type MOUSELLHOOKSTRUCT
  pt As POINTAPI
  MouseData As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private pressedKey As String, pressedKeyRow As Long
Private tmpS As String, tmpL As Long, tmpR As Range

Private v1 As MOUSELLHOOKSTRUCT, v2 As KBDLLHOOKSTRUCT


Private previousRowRange As Range, previousRowArray As Variant, previousLastArg As Long
Private previous2RowRange As Range, previous2RowArray As Variant, previous2LastArg As Long
Private windowTitle As String



Public Function Record(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  MouseHookHandle = 0
  KeyHookHandle = 0

  If ufAutoWin.SelectedLine < 1 Then GoTo done Else currentRow = ufAutoWin.SelectedLine

  Call ClearKeyPressColumn

  currentRow = currentRow + 1: Call SaveCurrentRowValues
  currentRowRange(1, ColACommand).Value = "Set Keyboard Layout"
  Call GetKeybLayout

  currentRow = currentRow + 1: Call SaveCurrentRowValues

  'ActiveWindow.WindowState = xlMinimized
  ufRecorder.Show
  'ActiveWindow.WindowState = xlNormal

done:
  Application.StatusBar = False
  Record = True
  Exit Function
eh:
  Application.StatusBar = False
  Record = False
  RaiseError MODULE_NAME & ".Record", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function MouseHookProc(ByVal lngCode As Long, ByVal wParam As LongPtr, ByRef lParam As MOUSELLHOOKSTRUCT) As LongPtr
'https://social.msdn.microsoft.com/Forums/sqlserver/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba
  On Error GoTo eh
  If (lngCode = HC_ACTION) Then
    v1 = lParam
    Select Case wParam
      Case WM_MOUSEWHEEL 'not recorded yet !!!
        If lParam.MouseData > 0 Then
          pressedKey = " Wheel up"
        Else
          pressedKey = " Wheel down"
        End If
      Case WM_LBUTTONDOWN
        Call RecordMouseDown("Left")
      Case WM_LBUTTONUP
        Call RecordMouseUp("Left")
      Case WM_RBUTTONDOWN
        Call RecordMouseDown("Right")
      Case WM_RBUTTONUP
        Call RecordMouseUp("Right")
      Case WM_MBUTTONDOWN
        Call RecordMouseDown("Middle")
      Case WM_MBUTTONUP
        Call RecordMouseUp("Middle")
    End Select
  End If
  MouseHookProc = CallNextHookEx(0, lngCode, wParam, ByVal lParam)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".MouseHookProc", Err.Number, Err.Source, Err.Description, Erl, , True
End Function

Public Function KeyHookProc(ByVal lngCode As Long, ByVal wParam As LongPtr, ByRef lParam As KBDLLHOOKSTRUCT) As LongPtr
'https://renenyffenegger.ch/notes/development/languages/VBA/Win-API/examples/SetWindowsHookEx/index
  On Error GoTo eh
  'Debug.Print "KeyHookProc lngCode=" & lngCode & ", wParam=" & wParam
  If (lngCode = HC_ACTION) Then
    v2 = lParam
    Select Case wParam
      Case WM_KEYDOWN, WM_SYSKEYDOWN
        Call RecordFocusedControlState("KeyDown")
        Call RecordKeyDown("Key Down")
      Case WM_KEYUP, WM_SYSKEYUP
        Call RecordFocusedControlState("KeyUp")
        Call RecordKeyUp("Key Up")
    End Select
  End If
  KeyHookProc = CallNextHookEx(0, lngCode, wParam, ByVal lParam)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".KeyHookProc", Err.Number, Err.Source, Err.Description, Erl, , True
End Function



Private Sub RecordMouseDown(Text As String)
  On Error GoTo eh

  Call AttachThreadToWindowFromPoint(v1.pt.x, v1.pt.y) ' is needed so we can get the handle of a focused window in another app
  windowTitle = GetWindowTitleFromPoint(v1.pt.x, v1.pt.y)
  If windowTitle = ufRecorder.Caption Then Exit Sub

  Call SavePreviousRowValues

  If Len(windowTitle) > 0 Then
    If previousRowArray(1, ColAWindow) <> windowTitle Then
    'If .Cells(currentRow - 1, ColAWindow).Value <> tmpS Then

      currentRowRange(1, ColACommand).Value = "Activate Window by Name"
      currentRowRange(1, ColAPause).Value = 200
      currentRowRange(1, ColAArg1).Value = windowTitle

      currentRow = currentRow + 1: Call SaveCurrentRowValues

      currentRowRange(1, ColACommand).Value = "Set Window Position"
      currentRowRange(1, ColAPause).Value = 500
      currentRowRange(1, ColAArg1).Value = windowTitle
      Call GetWindowPosition

      currentRow = currentRow + 1: Call SaveCurrentRowValues
  End If: End If

  currentRowRange(1, ColACommand).Value = Text & IIf(Len(Text) > 6, "", " click Down")
  currentRowRange(1, ColAArg1 + 0).Value = v1.pt.x
  currentRowRange(1, ColAArg1 + 1).Value = v1.pt.y
  currentRowRange(1, ColAWindow).Value = windowTitle

  Call SaveCurrentRowValues
  currentRowRange(1, ColAColor).Value = GetColorFromPointAsHex(v1.pt.x, v1.pt.y)
  Call WritePauseCol

  currentRow = currentRow + 1: Call SaveCurrentRowValues
  Call CenterViewToCurrentRow

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RecordMouseDown", Err.Number, Err.Source, Err.Description, Erl
End Sub

Private Sub RecordMouseUp(Text As String)
  On Error GoTo eh
  Text = Text & " click Up"

  Call SavePreviousRowValues

  If Left(Text, 1) = Left(previousRowArray(1, ColACommand), 1) Then 'L, R or M
    If AreSamePointsLong(CLng(previousRowArray(1, ColAArg1 + 0)), _
                         CLng(previousRowArray(1, ColAArg1 + 1)), _
                         v1.pt.x, v1.pt.y, positionTolerance) Then

        previousRowRange(1, ColACommand).Value = Left(Text, Len(Text) - 3) ' Remove " Up" and will remain "Left Click" or "Right Click" or "Middle Click"
        Exit Sub
  End If: End If

  Call RecordMouseDown(Text)

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RecordMouseUp", Err.Number, Err.Source, Err.Description, Erl
End Sub




Private Sub RecordKeyDown(Text As String)
  On Error GoTo eh

  Call StopMoveDetection
  windowTitle = GetActiveWindowTitle
  If windowTitle = ufRecorder.Caption Then Exit Sub

  Set tmpR = shKey.Columns(ColKeyCodeDec).Find(v2.VKcode, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
  If tmpR Is Nothing Then Exit Sub ' code not found

  pressedKeyRow = tmpR.Row
  pressedKey = shKey.Cells(pressedKeyRow, ColKeyName).Text


  Dim keybLayout As String: keybLayout = GetKeybLayoutAsString

  Dim keybLayoutOfKey As String: keybLayoutOfKey = shKey.Cells(pressedKeyRow, ColKeyPressed).Value
  If Len(keybLayout) > 0 And Len(keybLayoutOfKey) > 0 Then
    If InStr(keybLayout, keybLayoutOfKey) > 0 Then
      Exit Sub ' the key is already pressed, on the same keyboard layout, so do not record anything
    End If
  End If


  Call SavePreviousRowValues


  If previousRowArray(1, ColACommand) = Text Then ' previous action name is also "Key Down"
    For tmpL = 1 To 9 ' Find first free argument
      If Len(previousRowArray(1, ColAArg1 + tmpL)) = 0 Then Exit For
    Next

    If tmpL <= 9 Then ' There is a free argument => add this key there
      previousRowRange(1, ColAArg1 + tmpL).Value = pressedKey
      Exit Sub
    Else
      ' All arguments are full (this would mean that 10 Keys are pressed at once => obviously not normal) - will not be implemented
    End If
  End If

  shKey.Cells(pressedKeyRow, ColKeyPressed).Value = keybLayout ' meaning: key is pressed

  If Len(windowTitle) > 0 Then
    If previousRowArray(1, ColAWindow) <> windowTitle Then

      'AttachThreadToActiveWindow ' is needed so we can get the handle of a focused window in another app
      currentRowRange(1, ColACommand).Value = "Activate Window by Name"
      currentRowRange(1, ColAPause).Value = 200
      currentRowRange(1, ColAArg1).Value = windowTitle

      currentRow = currentRow + 1: Call SaveCurrentRowValues

      currentRowRange(1, ColACommand).Value = "Set Window Position"
      currentRowRange(1, ColAPause).Value = 500
      currentRowRange(1, ColAArg1).Value = windowTitle
      Call GetWindowPosition

      currentRow = currentRow + 1: Call SaveCurrentRowValues
  End If: End If

  currentRowRange(1, ColACommand).Value = Text
  currentRowRange(1, ColAArg1).Value = pressedKey
  currentRowRange(1, ColAKeybd).Value = keybLayout
  currentRowRange(1, ColAWindow).Value = windowTitle

  Call WritePauseCol

  currentRow = currentRow + 1: Call SaveCurrentRowValues
  Call CenterViewToCurrentRow

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RecordKeyDown", Err.Number, Err.Source, Err.Description, Erl
End Sub

Private Sub RecordKeyUp(Text As String)
  On Error GoTo eh

  Set tmpR = shKey.Columns(ColKeyCodeDec).Find(v2.VKcode, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
  If tmpR Is Nothing Then Exit Sub ' code not found

  pressedKeyRow = tmpR.Row
  pressedKey = shKey.Cells(pressedKeyRow, ColKeyName).Text
  
  shKey.Cells(pressedKeyRow, ColKeyPressed).Value = "" 'meaning key is released

  windowTitle = GetActiveWindowTitle

  Call SavePreviousRowValues


  If previousLastArg >= 0 Then ' there is also a previous command
    Select Case CleanString(CStr(previousRowArray(1, ColACommand)))
  
      Case "keydown":
        If previousRowArray(1, ColAArg1 + previousLastArg) = pressedKey Then
          ' last pressed key is this key that was now released => so replace to Key Press
  
          If InStr(pressedKey, "CONTROL") > 0 Or InStr(pressedKey, "CTRL") > 0 Then ' the position of the mouse will be registered and the CTRL key press will be omited
            If previousLastArg = 0 Then ' there is only the Ctrl pressed key; the command will be changed to Move Mouse
              currentRow = currentRow - 1: Call SaveCurrentRowValues
              currentRowRange(1, ColACommand).Value = "Move Mouse"
              Call ReadMouseXY
              currentRowRange(1, ColAWindow).Value = windowTitle
              Call WritePauseCol
              currentRow = currentRow + 1: Call SaveCurrentRowValues
              GoTo done
  
            ElseIf previousLastArg > 0 Then ' there are more pressed keys in the command; the last Ctrl key will be removed from Key Down and a new Move Mouse will be added
              previousRowRange(1, ColAArg1 + previousLastArg).Value = ""
              currentRowRange(1, ColACommand).Value = "Move Mouse"
              Call ReadMouseXY
              currentRowRange(1, ColAWindow).Value = windowTitle
              Call WritePauseCol
              currentRow = currentRow + 1: Call SaveCurrentRowValues
              GoTo done
  
            End If
          End If
  
          If previousLastArg = 0 Then ' there is a single pressed key; the command will be changed to Key Press

            shAuto.Cells(currentRow - 1, ColACommand).Value = "Key Press"
            GoTo done

          ElseIf previousLastArg > 0 Then ' there are more pressed keys in the command; this key will be removed from Key Down and moved to Key Press
          
            currentRowRange(1, ColACommand).Value = "Key Press"
            currentRowRange(1, ColAArg1).Value = pressedKey
            previousRowRange(1, ColAArg1 + previousLastArg).Value = ""
            currentRowRange(1, ColAWindow).Value = windowTitle
            Call WritePauseCol
            currentRow = currentRow + 1: Call SaveCurrentRowValues
            GoTo done
          
          Else ' tmpl<0 should not be possible, it would mean that the command Key Down has no argument
            ' not implemented :D
            RaiseError MODULE_NAME & ".RecordKeyUp", Err.Number, Err.Source, "Key " & pressedKey & " was released. Previous command is Key Down with no arguments - this is impossible. This was not implemented.", Erl, 1
            Exit Sub
  
          End If
  
        Else ' last pressed key is not this key that was now released
          ' simply record current Key Up command
  
        End If
  
      Case "keyup":
        
        If previousLastArg = 9 Then ' There are no more free arguments, so just add a new Key Up event => this should not be possible, meaning there were more than 10 Keys pressed at once
          ' not implemented :D
          RaiseError MODULE_NAME & ".RecordKeyUp", Err.Number, Err.Source, "Key " & pressedKey & " was released. Previous command is Key Up with all 10 arguments full - this is impossible, that 10+1 Keys are pressed then released. This was not implemented.", Erl, 2
          Exit Sub
  
        Else ' There is a free argument to add current Key
          previousRowRange(1, ColAArg1 + previousLastArg).Value = pressedKey
          GoTo done
  
        End If
  
      Case "keypress":

        If previousLastArg = 9 Then ' There are no more free arguments, so just add a new Key Up event => this should not be possible, meaning there were more than 10 Keys pressed at once
          ' not implemented :D
          RaiseError MODULE_NAME & ".RecordKeyUp", Err.Number, Err.Source, "Key " & pressedKey & " was released. Previous command is Key Press with all 10 arguments full - this is impossible, that 10+1 Keys are pressed then released. This was not implemented.", Erl, 3
          Exit Sub
        
        Else ' There is a free argument to add current Key
          If previous2LastArg >= 0 Then ' there is also a 2nd previous command
            
            If CleanString(CStr(previous2RowArray(1, ColACommand))) = "keydown" Then ' 2nd previous command is Key Down and previous command is Key Press
            
              If previous2RowArray(1, ColAArg1 + previous2LastArg) = pressedKey Then ' last pressed key is this key that was now released => so move it to Key Press, in the first position
                
                If previous2LastArg = 0 Then ' there is only one argument in Key Down, so change to Key Press and move all arguments of Key Press here
                  previous2RowRange(1, ColACommand).Value = "Key Press"
                  
                  For tmpL = previousLastArg To 0 Step -1
                    previous2RowRange(1, ColAArg1 + tmpL + 1).Value = previousRowArray(1, ColAArg1 + tmpL)
                  Next

                  previousRowRange.ClearContents
                  currentRow = currentRow - 1: Call SaveCurrentRowValues
                  GoTo done
                
                Else ' there are more arguments in Key Down
                
                  ' first move all arguments of the Key Press one place further
                  For tmpL = previousLastArg To 0 Step -1
                    previousRowRange(1, ColAArg1 + tmpL + 1).Value = previousRowArray(1, ColAArg1 + tmpL)
                  Next

                  'move the key from Key Down to Key Press
                  previousRowRange(1, ColAArg1).Value = pressedKey
                  previous2RowRange(1, ColAArg1 + previous2LastArg).Value = ""
                  GoTo done

                End If
                
              Else ' last pressed key is not this key that was now released
                ' simply record current Key Up command
              
              End If
            
            Else ' 2nd previous command is not Key Down command
              ' simply record current Key Up command
              
            End If
            
          Else ' there is not a 2nd previous command
            ' simply record current Key Up command
            
          End If
          
        End If
      
      
      Case Else: ' previous command is another command
        ' simply record current Key Up command

    End Select

  Else ' there is no previous command
    ' simply record current Key Up command

  End If

  currentRowRange(1, ColACommand).Value = "Key Up"
  currentRowRange(1, ColAArg1).Value = pressedKey
  currentRowRange(1, ColAWindow).Value = windowTitle
  Call WritePauseCol
  currentRow = currentRow + 1: Call SaveCurrentRowValues

done:
  Call CenterViewToCurrentRow
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RecordKeyUp", Err.Number, Err.Source, Err.Description, Erl
End Sub









Private Function WritePauseCol()
  On Error GoTo eh
  currentRowRange(1, ColAPause).Value = EndTimer

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".WritePauseCol", Err.Number, Err.Source, Err.Description, Erl
End Function



Private Function SavePreviousRowValues()
  On Error GoTo eh

  If currentRow - 1 >= startRow Then ' there is a previous command
    Set previousRowRange = shAuto.Cells(currentRow - 1, ColAStatus).Resize(1, ColAComment - ColAStatus + 1)
    previousRowArray = previousRowRange.Value2
  
    For previousLastArg = 9 To 0 Step -1 ' Find last argument
      If Len(previousRowArray(1, ColAArg1 + previousLastArg)) <> 0 Then Exit For
    Next

    If currentRow - 2 >= startRow Then ' there is a 2nd previous command
      Set previous2RowRange = shAuto.Cells(currentRow - 2, ColAStatus).Resize(1, ColAComment - ColAStatus + 1)
      previous2RowArray = previous2RowRange.Value2

      For previous2LastArg = 9 To 0 Step -1 ' Find last argument
        If Len(previous2RowArray(1, ColAArg1 + previous2LastArg)) <> 0 Then Exit For
      Next
      
    Else ' this is the second command
      previous2LastArg = -1
    End If

  Else ' this is the first command
    previousLastArg = -1
    previous2LastArg = -1
  End If

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".SavePreviousRowValues", Err.Number, Err.Source, Err.Description, Erl
End Function






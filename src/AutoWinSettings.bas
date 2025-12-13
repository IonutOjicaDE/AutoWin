Attribute VB_Name = "AutoWinSettings"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinSettings"

Public Const debugLevel = 0 '1=basic, 5=verbose => how much information to be displayed in the Immediate Window


Public commandMap As Object          ' Dictionary to store commands
' === Sheet Automation, CodeName ShAuto      ===
Public Const cmdFunctionName    As Long = 0& ' Name of VBA function
Public Const cmdDisplayName     As Long = 1& ' Pretty name with spaces
Public Const cmdCategory        As Long = 2& ' Category (General, Mouse, Window, etc.)
Public Const cmdDescription     As Long = 3& ' Command description
Public Const cmdArgName1        As Long = 4& ' Argument names
Public Const cmdArgDescription1 As Long = 5& ' Argument descriptions
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)


Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#If Win64 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If


' === sheet Automation, first row with commands                                                           ===
                                  Public Const startRow     As Long = 2&
Public currentRow   As Long ' here is current executed command


' === Maximum number of continuous empty rows to stop AutoWin                                             ===
Public maxEmptyRows As Long:      Public Const maxEmptyRowsDefault As Long = 10&
Public emptyRows    As Long

' === sheet Automation, CodeName ShAuto                                                                   ===
' === Set CurrentRowRange = shAuto.Cells(CurrentRow, ColAStatus).Resize(1, ColAComment - ColAStatus + 1)  ===
' === CurrentRowArray     = CurrentRowRange.Value2                                                        ===
' === CurrentRowRange.Cells(1, ColAStatus).Value  = CurrentRowArray(1, ColAStatus)                        ===
' === CurrentRowRange.Cells(1, ColACommand).Value = CurrentRowArray(1, ColACommand)                       ===
' === if CurrentRowArray(1, ColAArg1) is Date, can be retrieved using CDate(CurrentRowArray(1, ColAArg1)) ===
Public currentRowRange       As Range
Public currentRowArray       As Variant


Public stopExecutionRequired As Boolean ' default False
Public errorNumber           As Long
Public errorSource           As String
Public errorDescription      As String


' === minimum waiting time between commands, if nothing is specified, in ms           ===
Public minWaitTime   As Long:     Public Const minWaitTimeDefault       As Long = 50&
' === during waiting, check interval for moving mouse by user to break AutoWin, in ms ===
Public waitTimeSplit As Long:     Public Const waitTimeSplitDefault     As Long = 500&


' === maximum duration to wait for the foreground window title to match, in ms        ===
Public windowCheckMax   As Long:  Public Const windowCheckMaxDefault    As Long = 5000&
' === foreground window title will be checked each 200ms                              ===
Public windowCheckSplit As Long:  Public Const windowCheckSplitDefault  As Long = 200&


' === maximum duration to wait for the color under cursor to match, in ms             ===
Public colorCheckMax As Long:     Public Const colorCheckMaxDefault     As Long = 5000&
' === color under cursor will be checked each 200ms                                   ===
Public colorCheckSplit As Long:   Public Const colorCheckSplitDefault   As Long = 100&
' === color tolerance between detected and expected color, applied to each channel    ===
Public maxColorTolerance As Long: Public Const maxColorToleranceDefault As Long = 10&
' === read color on the further mouse commands, before moving mouse                   ===
Public colorReadFirst As Boolean: Public Const colorReadFirstDefault    As Boolean = False


' === waiting time after a keypress, in ms                                            ===
Public waitAfterKeyPress As Long: Public Const waitAfterKeyPressDefault As Long = 20&

' === waiting time after a mouse move, in ms                                          ===
Public waitAfterMove As Long:     Public Const waitAfterMoveDefault     As Long = 50&
' === waiting time after a click, in ms                                               ===
Public waitAfterClick As Long:    Public Const waitAfterClickDefault    As Long = 20&
' === position tolerance to consider that mouse was not moved by user, in pixels      ===
Public positionTolerance As Long: Public Const positionToleranceDefault As Long = 10&


' === ForEach loops will start looping trough rows in a rectangular range             ===
Public startWithRows As Boolean:  Public Const startWithRowsDefault     As Boolean = True
' === increase loops_array in blocks, for performance reasons                         ===
                                  Public Const loopStackBlock           As Long = 10&
' === number of nested loops after warning will be triggered                          ===
Public loopStackMaxSize As Long:  Public Const loopStackMaxSizeDefault  As Long = 50&


' === Mouse moved in steps:                                                           ===
' 1. Move mouse to Target + ShakeMovementDist                                         ===
' 2. Wait ShakeMovementWait (Intermediate mouse position, to avoid bugs on automation)===
' 3. Move mouse to Target                                                             ===
' 4. Wait WaitAfterMove - ShakeMovementWait                                           ===
Public shakeMovementDist As Long: Public Const shakeMovementDistDefault As Long = 2&
Public shakeMovementWait As Long: Public Const shakeMovementWaitDefault As Long = 10&

Public Const recordText    As String = "record"
Public Const loopTypeNext  As String = "next"
Public Const loopTypeLoop  As String = "loop"
Public Const loopTypeSub   As String = "sub"
Public Const loopTypeLabel As String = "label"

Public Const loopListSub   As String = "{list_sub}"   ' sub
Public Const loopListLabel As String = "{list_label}" ' label
Public Const loopListFor   As String = "{list_for}"   ' for, foreach
Public Const loopListLoop  As String = "{list_loop}"  ' do, dowhile, dountil

' Shortcodes for KeyPress sheet
Public Const keyPressName     As String = "{key_name}"      ' maps to ColKeyName
Public Const keyPressSendKeys As String = "{key_sendkeys}"  ' maps to ColKeySendKeys

Public Const loopListFlagNone  As Long = 0&
Public Const loopListFlagSub   As Long = 1&
Public Const loopListFlagLabel As Long = 2&
Public Const loopListFlagFor   As Long = 4&
Public Const loopListFlagLoop  As Long = 8&


Public ApplicationStatusBar As String

'#############################################
'#######                               #######
'#######       InitializeSettings      #######
'#######                               #######
'#############################################

Public Function InitializeSettings()
  On Error GoTo eh
  Call InitializeParameters
  Call InitializeCommandMap

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".InitializeSettings", Err.Number, Err.Source, Err.Description, Erl
End Function

Private Function InitializeParameters()
  On Error GoTo eh
  currentRow = startRow
  maxEmptyRows = maxEmptyRowsDefault
  emptyRows = 0&
  
  minWaitTime = minWaitTimeDefault
  waitTimeSplit = waitTimeSplitDefault
  
  windowCheckMax = windowCheckMaxDefault
  windowCheckSplit = windowCheckSplitDefault
  
  colorCheckMax = colorCheckMaxDefault
  colorCheckSplit = colorCheckSplitDefault
  maxColorTolerance = maxColorToleranceDefault
  colorReadFirst = colorReadFirstDefault

  waitAfterKeyPress = waitAfterKeyPressDefault
  
  waitAfterMove = waitAfterMoveDefault
  waitAfterClick = waitAfterClickDefault
  positionTolerance = positionToleranceDefault
  shakeMovementDist = shakeMovementDistDefault
  shakeMovementWait = shakeMovementWaitDefault
  
  startWithRows = startWithRowsDefault
  loopStackMaxSize = loopStackMaxSizeDefault
  
  'Call SaveCurrentRowValues
  
  stopExecutionRequired = False
  errorNumber = 0&
  errorSource = vbNullString
  errorDescription = vbNullString
  
  ApplicationStatusBar = vbNullString

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".InitializeParameters", Err.Number, Err.Source, Err.Description, Erl
End Function

Private Function InitializeCommandMap()
  On Error GoTo eh
  If Not commandMap Is Nothing Then Set commandMap = Nothing
  Set commandMap = CreateObject("Scripting.Dictionary")

  Call RegisterCommandsCondition
  Call RegisterCommandsDisplay
  Call RegisterCommandsExecute
  Call RegisterCommandsFile
  Call RegisterCommandsKeyboard
  Call RegisterCommandsLoop
  Call RegisterCommandsMouse
  Call RegisterCommandsOfficeExcel
  Call RegisterCommandsOfficeOutlook
  Call RegisterCommandsOfficeWord
  Call RegisterCommandsScreenColor
  Call RegisterCommandsSendMessage
  Call RegisterCommandsWindow

  commandMap.Add recordText, Array("Record", "Record", _
    MODULE_NAME, "Start recording on current row")


  commandMap.Add "setmaxemptyrows", Array("SetMaxEmptyRows", "Set Max Empty Rows", _
    MODULE_NAME, "Maximum number of continuous empty rows to stop AutoWin", _
    "New Value", "Default is " & maxEmptyRowsDefault)


  commandMap.Add "setminimumwaitingtime", Array("SetMinWaitTime", "Set Minimum Waiting Time", _
    MODULE_NAME, "Minimum waiting time between commands, if nothing is specified, in ms", _
    "New Value", "Default is " & minWaitTimeDefault)
  
  commandMap.Add "setwaitingtick", Array("SetWaitTimeSplit", "Set Waiting Tick", _
    MODULE_NAME, "During waiting, check interval for Moving Mouse by user to break AutoWin, in ms", _
    "New Value", "Default is " & waitTimeSplitDefault)


  commandMap.Add "setmaxwaitforwindowcheck", Array("SetWindowCheckMax", "Set max Wait for Window Check", _
    MODULE_NAME, "Maximum duration to wait for the Foreground Window Title to match, in ms", _
    "New Value", "Default is " & windowCheckMaxDefault)
  
  commandMap.Add "setwindowchecktick", Array("SetWindowCheckSplit", "Set Window Check Tick", _
    MODULE_NAME, "During waiting, check interval for the Foreground Window Title to match, in ms", _
    "New Value", "Default is " & windowCheckSplitDefault)


  commandMap.Add "setmaxwaitforcolorcheck", Array("SetColorCheckMax", "Set max Wait for Color Check", _
    MODULE_NAME, "Maximum duration to wait for the Color Under Cursor to match, in ms", _
    "New Value", "Default is " & colorCheckMaxDefault)
  
  commandMap.Add "setcolorchecktick", Array("SetColorCheckSplit", "Set Color Check Tick", _
    MODULE_NAME, "During waiting, check interval for the Color Under Cursor to match, in ms", _
    "New Value", "Default is " & colorCheckSplitDefault)
  
  commandMap.Add "setmaxcolortolerance", Array("SetMaxColorTolerance", "Set max Color Tolerance", _
    MODULE_NAME, "Color tolerance between detected and expected color, applied to each channel", _
    "New Value", "Default is " & maxColorToleranceDefault)

  commandMap.Add "setcolorreadfirst", Array("SetColorReadFirst", "Set Color Read First", _
    MODULE_NAME, "Read color on the further mouse commands, before moving mouse.", _
    "New Value", "Default is " & colorReadFirstDefault & " {{True/False}}")


  commandMap.Add "setwaitafterkeypress", Array("SetWaitAfterKeyPress", "Set Wait After Key Press", _
    MODULE_NAME, "Waiting time after a Keypress, in ms", _
    "New Value", "Default is " & waitAfterKeyPressDefault)


  commandMap.Add "setwaitaftermousemove", Array("SetWaitAfterMove", "Set Wait After Mouse Move", _
    MODULE_NAME, "Waiting time after a Mouse Move, in ms", _
    "New Value", "Default is " & waitAfterMoveDefault)
  
  commandMap.Add "setwaitaftermouseclick", Array("SetWaitAfterClick", "Set Wait After Mouse Click", _
    MODULE_NAME, "Waiting time after a Mouse Click, in ms", _
    "New Value", "Default is " & waitAfterClickDefault)
  
  commandMap.Add "setmousepositiontolerance", Array("SetPositionTolerance", "Set Mouse Position Tolerance", _
    MODULE_NAME, "Position tolerance to consider that mouse was not moved by user, in pixels", _
    "New Value", "Default is " & positionToleranceDefault)
  
  commandMap.Add "setmouseshakemovementdist", Array("SetShakeMovementDist", "Set Mouse Shake Movement Dist", _
    MODULE_NAME, "Mouse moved in steps:" & vbCrLf & _
                 "1. Move mouse to Target + ShakeMovementDist" & vbCrLf & _
                 "2. Wait ShakeMovementWait (Intermediate mouse position, to avoid bugs on automation)" & vbCrLf & _
                 "3. Move mouse to Target" & vbCrLf & _
                 "4. Wait WaitAfterMove - ShakeMovementWait", _
    "New Value", "Default is " & shakeMovementDistDefault)

  commandMap.Add "setmouseshakemovementwait", Array("SetShakeMovementWait", "Set Mouse Shake Movement Wait", _
    MODULE_NAME, "Mouse moved in steps:" & vbCrLf & _
                 "1. Move mouse to Target + ShakeMovementDist" & vbCrLf & _
                 "2. Wait ShakeMovementWait (Intermediate mouse position, to avoid bugs on automation)" & vbCrLf & _
                 "3. Move mouse to Target" & vbCrLf & _
                 "4. Wait WaitAfterMove - ShakeMovementWait", _
    "New Value", "Default is " & shakeMovementWaitDefault)


  commandMap.Add "setforeachstartwithrows", Array("SetStartWithRows", "Set ForEach Start With Rows", _
    MODULE_NAME, "ForEach loops will start looping trough:" & vbCrLf & _
                 "True = rows in a rectangular range" & vbCrLf & _
                 "False = columns in a rectangular range", _
    "New Value", "Default is " & startWithRowsDefault & " {{True/False}}")
  
  commandMap.Add "setloopstackmaxsize", Array("SetLoopStackMaxSize", "Set Loop Stack Max Size", _
    MODULE_NAME, "Number of nested loops after a warning will be triggered, to prevent an Out Of Memory error", _
    "New Value", "Default is " & loopStackMaxSizeDefault)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".InitializeCommandMap", Err.Number, Err.Source, Err.Description, Erl
End Function

'#############################################
'#######                               #######
'#######          PrepareExit          #######
'#######                               #######
'#############################################

Public Function PrepareExit()
  On Error GoTo eh
  Call PrepareExitCommandsCondition
  Call PrepareExitCommandsDisplay
  Call PrepareExitCommandsExecute
  Call PrepareExitCommandsFile
  Call PrepareExitCommandsKeyboard
  Call PrepareExitCommandsLoop
  Call PrepareExitCommandsMouse
  Call PrepareExitCommandsOfficeExcel
  Call PrepareExitCommandsOfficeOutlook
  Call PrepareExitCommandsOfficeWord
  Call PrepareExitCommandsScreenColor
  Call PrepareExitCommandsSendMessage
  Call PrepareExitCommandsWindow

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExit", Err.Number, Err.Source, Err.Description, Erl
End Function

'#############################################
'#######                               #######
'#######             Other             #######
'#######                               #######
'#############################################

Public Sub test()
Dim a As String, b As Long, c As String
a = CStr(Cells(30, 4).Value)
Debug.Print "a=" & a & " is numeric=" & IsNumber(a)
b = InStr(a, ",")
Debug.Print b
c = Replace(a, ",", ".")
Debug.Print "c=" & c & " is numeric=" & IsNumber(c)
End Sub

Public Function SaveCurrentRowValues()
  On Error GoTo eh
  Set currentRowRange = shAuto.Cells(currentRow, ColAStatus).Resize(1, ColAComment - ColAStatus + 1)
  currentRowArray = currentRowRange.Value2

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".SaveCurrentRowValues", Err.Number, Err.Source, Err.Description, Erl
End Function

'#############################################
'#######                               #######
'#######         Set Parameters        #######
'#######                               #######
'#############################################

Public Function SetMaxEmptyRows(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    maxEmptyRows = CLng(currentRowArray(1, ColAArg1 + 0))
    SetMaxEmptyRows = True
    Exit Function
  End If

  SetMaxEmptyRows = False
  RaiseError MODULE_NAME & ".SetMaxEmptyRows", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetMaxEmptyRows = False
  RaiseError MODULE_NAME & ".SetMaxEmptyRows", Err.Number, Err.Source, Err.Description, Erl
End Function
  
Public Function SetMinWaitTime(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    minWaitTime = CLng(currentRowArray(1, ColAArg1 + 0))
    SetMinWaitTime = True
    Exit Function
  End If

  SetMinWaitTime = False
  RaiseError MODULE_NAME & ".SetMinWaitTime", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetMinWaitTime = False
  RaiseError MODULE_NAME & ".SetMinWaitTime", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetWaitTimeSplit(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    waitTimeSplit = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWaitTimeSplit = True
    Exit Function
  End If

  SetWaitTimeSplit = False
  RaiseError MODULE_NAME & ".SetWaitTimeSplit", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWaitTimeSplit = False
  RaiseError MODULE_NAME & ".SetWaitTimeSplit", Err.Number, Err.Source, Err.Description, Erl
End Function
  
Public Function SetWindowCheckMax(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    windowCheckMax = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWindowCheckMax = True
    Exit Function
  End If

  SetWindowCheckMax = False
  RaiseError MODULE_NAME & ".SetWindowCheckMax", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWindowCheckMax = False
  RaiseError MODULE_NAME & ".SetWindowCheckMax", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetWindowCheckSplit(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    windowCheckSplit = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWindowCheckSplit = True
    Exit Function
  End If

  SetWindowCheckSplit = False
  RaiseError MODULE_NAME & ".SetWindowCheckSplit", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWindowCheckSplit = False
  RaiseError MODULE_NAME & ".SetWindowCheckSplit", Err.Number, Err.Source, Err.Description, Erl
End Function
  
  
Public Function SetColorCheckMax(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    colorCheckMax = CLng(currentRowArray(1, ColAArg1 + 0))
    SetColorCheckMax = True
    Exit Function
  End If

  SetColorCheckMax = False
  RaiseError MODULE_NAME & ".SetColorCheckMax", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetColorCheckMax = False
  RaiseError MODULE_NAME & ".SetColorCheckMax", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetColorCheckSplit(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    colorCheckSplit = CLng(currentRowArray(1, ColAArg1 + 0))
    SetColorCheckSplit = True
    Exit Function
  End If

  SetColorCheckSplit = False
  RaiseError MODULE_NAME & ".SetColorCheckSplit", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetColorCheckSplit = False
  RaiseError MODULE_NAME & ".SetColorCheckSplit", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetMaxColorTolerance(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    maxColorTolerance = CLng(currentRowArray(1, ColAArg1 + 0))
    SetMaxColorTolerance = True
    Exit Function
  End If

  SetMaxColorTolerance = False
  RaiseError MODULE_NAME & ".SetMaxColorTolerance", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetMaxColorTolerance = False
  RaiseError MODULE_NAME & ".SetMaxColorTolerance", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetColorReadFirst(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  colorReadFirst = GetBoolean(CStr(currentRowArray(1, ColAArg1 + 0)))
  SetColorReadFirst = True
End Function


Public Function SetWaitAfterKeyPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    waitAfterKeyPress = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWaitAfterKeyPress = True
    Exit Function
  End If

  SetWaitAfterKeyPress = False
  RaiseError MODULE_NAME & ".SetWaitAfterKeyPress", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWaitAfterKeyPress = False
  RaiseError MODULE_NAME & ".SetWaitAfterKeyPress", Err.Number, Err.Source, Err.Description, Erl
End Function
  
Public Function SetWaitAfterMove(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    waitAfterMove = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWaitAfterMove = True
    Exit Function
  End If

  SetWaitAfterMove = False
  RaiseError MODULE_NAME & ".SetWaitAfterMove", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWaitAfterMove = False
  RaiseError MODULE_NAME & ".SetWaitAfterMove", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetWaitAfterClick(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    waitAfterClick = CLng(currentRowArray(1, ColAArg1 + 0))
    SetWaitAfterClick = True
    Exit Function
  End If

  SetWaitAfterClick = False
  RaiseError MODULE_NAME & ".SetWaitAfterClick", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetWaitAfterClick = False
  RaiseError MODULE_NAME & ".SetWaitAfterClick", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetPositionTolerance(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    positionTolerance = CLng(currentRowArray(1, ColAArg1 + 0))
    SetPositionTolerance = True
    Exit Function
  End If

  SetPositionTolerance = False
  RaiseError MODULE_NAME & ".SetPositionTolerance", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetPositionTolerance = False
  RaiseError MODULE_NAME & ".SetPositionTolerance", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetShakeMovementDist(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    shakeMovementDist = CLng(currentRowArray(1, ColAArg1 + 0))
    SetShakeMovementDist = True
    Exit Function
  End If

  SetShakeMovementDist = False
  RaiseError MODULE_NAME & ".SetShakeMovementDist", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetShakeMovementDist = False
  RaiseError MODULE_NAME & ".SetShakeMovementDist", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function SetShakeMovementWait(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    shakeMovementWait = CLng(currentRowArray(1, ColAArg1 + 0))
    SetShakeMovementWait = True
    Exit Function
  End If

  SetShakeMovementWait = False
  RaiseError MODULE_NAME & ".SetShakeMovementWait", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetShakeMovementWait = False
  RaiseError MODULE_NAME & ".SetShakeMovementWait", Err.Number, Err.Source, Err.Description, Erl
End Function

Public Function SetStartWithRows(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  startWithRows = GetBoolean(CStr(currentRowArray(1, ColAArg1 + 0)))
  SetStartWithRows = True
End Function
Public Function SetLoopStackMaxSize(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then
    loopStackMaxSize = CLng(currentRowArray(1, ColAArg1 + 0))
    SetLoopStackMaxSize = True
    Exit Function
  End If

  SetLoopStackMaxSize = False
  RaiseError MODULE_NAME & ".SetLoopStackMaxSize", Err.Number, Err.Source, "Argument need to be valid number: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 1

done:
  Exit Function
eh:
  SetLoopStackMaxSize = False
  RaiseError MODULE_NAME & ".SetLoopStackMaxSize", Err.Number, Err.Source, Err.Description, Erl
End Function


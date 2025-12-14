Attribute VB_Name = "CommandsScreenColor"
Option Explicit
Private Const MODULE_NAME As String = "CommandsScreenColor"

Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'https://www.vbarchiv.net/tipps/tipp_1830-pixelfarbe-unter-dem-mauszeiger-bildschirmweit.html
'https://rosettacode.org/wiki/Color_of_a_screen_pixel#VBA

Private ColorTemp As Long
Private WindowDC  As Long
Private i         As Long
Private tmpL      As Long
Private tmpS      As String
Private tmpR      As Range
Private MousePos  As POINTAPI

Private Const colorFormatLong As Long = 1 ' white = 16 777 215, black = 0, red = 255, green = 65 280, blue = 16 711 680
Private Const colorFormatRGB  As Long = 2 ' white = 255, 255, 255 , black = 0 0 0 , red = 255,0,0 , green = 0,255,0
Private Const colorFormatHEX  As Long = 3 ' white = #FFFFFF = #FFF, black = #000000 = #000, red = #FF0000 = #f00, green = #00ff00 = #0f0

Public Function RegisterCommandsScreenColor()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "getcolorundercursor", Array("GetColorUnderCursor", "Get Color Under Cursor", _
    MODULE_NAME, "Retrieves the color of the pixel that is under the cursor.", _
    "x", "Here will be written the x-value of the mouse position.", _
    "y", "Here will be written the y-value of the mouse position.", _
    "Color", "Here will be written the color of the pixel in Hex format.")
  commandMap.Add "getcolorfrompoint", Array("GetColorFromPoint", "Get Color From Point", _
    MODULE_NAME, "Retrieves the color of the pixel from the given coordinates.", _
    "x", "X-value of the pixel to retrieve the color.", _
    "y", "Y-value of the pixel to retrieve the color.", _
    "Color", "Here will be written the color of the pixel in Hex format.")


  commandMap.Add "ifcolorthenskip", Array("CommandIfColorThenSkip", "If Color Then Skip", MODULE_NAME, _
    "Skip lines according to the pixel color of the pixel position.", _
    "x", "X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position.", _
    "y", "Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position.", _
    "Color", "Color of the pixel in Hex format; if missing or invalid, then stop with error.", _
    "Skip Lines if Color match", "How many lines to be skipped if the color matches.", _
    "Skip Lines if Color does not match", "How many lines to be skipped if the color does not matches.")

  commandMap.Add "ifcolorthengoto", Array("CommandIfColorThenGoTo", "If Color Then GoTo", MODULE_NAME, _
    "Set next line or label to be executed according to the pixel color of the pixel position.", _
    "x", "X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position.", _
    "y", "Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position.", _
    "Color", "Color of the pixel in Hex format; if missing or invalid, then stop with error.", _
    "Lines if Color match", "What line or label should be next if the color matches." & loopListLabel, _
    "Lines if Color does not match", "What line or label should be next if the color does not matches." & loopListLabel)


  commandMap.Add "ifwaitcolorthenskip", Array("CommandIfWaitColorThenSkip", "If Wait Color Then Skip", MODULE_NAME, _
    "Skip lines according to the pixel color of the pixel position. If color does not match, wait (Set max Wait for Color Check) for the color to match.", _
    "x", "X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position.", _
    "y", "Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position.", _
    "Color", "Color of the pixel in Hex format; if missing or invalid, then stop with error.", _
    "Skip Lines if Color match", "How many lines to be skipped if the color matches.", _
    "Skip Lines if Color does not match", "How many lines to be skipped if the color does not matches.")

  commandMap.Add "ifwaitcolorthengoto", Array("CommandIfWaitColorThenGoTo", "If Wait Color Then GoTo", MODULE_NAME, _
    "Set next line or label to be executed according to the pixel color of the pixel position. If color does not match, wait (Set max Wait for Color Check) for the color to match.", _
    "x", "X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position.", _
    "y", "Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position.", _
    "Color", "Color of the pixel in Hex format; if missing or invalid, then stop with error.", _
    "Lines if Color match", "What line or label should be next if the color matches." & loopListLabel, _
    "Lines if Color does not match", "What line or label should be next if the color does not matches." & loopListLabel)

  WindowDC = GetWindowDC(0)
done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsScreenColor", Err.Number, Err.Source, Err.Description, Erl
End Function

Public Function PrepareExitCommandsScreenColor()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsScreenColor", Err.Number, Err.Source, Err.Description, Erl
End Function


Private Function IsValidClickCommand() As Boolean
  tmpS = CleanString(CStr(currentRowArray(1, ColACommand)))

  If InStr(tmpS, "mouse") = 0 And InStr(tmpS, "click") = 0 Then Exit Function
  If Not IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then Exit Function
  If Not IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then Exit Function

  IsValidClickCommand = True
End Function




Public Function CommandIfColorThenSkip(Optional ExecutingTroughApplicationRun As Boolean = False)
' Arg1: X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position
' Arg2: Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position
' Arg3: Color of the pixel in Hex format; if missing or invalid, then stop with error
' Arg4: How many lines to be skipped if the color matches
' Arg5: How many lines to be skipped if the color does not matches

  On Error GoTo eh

  Static colorText As String: colorText = Trim(CStr(currentRowArray(1, ColAArg1 + 2)))
  If Len(colorText) = 0 Then
    CommandIfColorThenSkip = False
    RaiseError MODULE_NAME & "CommandIfColorThenSkip", Err.Number, Err.Source, _
      "Color needs to be entered in Arg3", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Static colorFormat As Byte
  Static colorLng As Long

  colorLng = DetectColorFormatAndExtractColor(colorText, colorFormat)
  If colorLng < 0 Then
    ' Color format invalid
    CommandIfColorThenSkip = False
    RaiseError MODULE_NAME & "CommandIfColorThenSkip", Err.Number, Err.Source, _
      "Color needs to be valid color code: Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  GetCursorPos MousePos
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))

  tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
  ' Check if color matches within tolerance
  If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
    ' Color matches within tolerance
    SkipLines CStr(currentRowArray(1, ColAArg1 + 3)), "Arg4"
  Else
    ' Color does not matches within tolerance
    SkipLines CStr(currentRowArray(1, ColAArg1 + 4)), "Arg5"
  End If

NormalExecution:
  CommandIfColorThenSkip = True
  Exit Function
eh:
  CommandIfColorThenSkip = False
  RaiseError MODULE_NAME & ".CommandIfColorThenSkip", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function CommandIfColorThenGoTo(Optional ExecutingTroughApplicationRun As Boolean = False)
' Arg1: X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position
' Arg2: Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position
' Arg3: Color of the pixel in Hex format; if missing or invalid, then stop with error
' Arg4: What line or label should be next if the color matches
' Arg5: What line or label should be next if the color does not matches

  On Error GoTo eh

  Static colorText As String: colorText = Trim(CStr(currentRowArray(1, ColAArg1 + 2)))
  If Len(colorText) = 0 Then
    CommandIfColorThenGoTo = False
    RaiseError MODULE_NAME & "CommandIfColorThenGoTo", Err.Number, Err.Source, _
      "Color needs to be entered in Arg3", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Static colorFormat As Byte
  Static colorLng As Long

  colorLng = DetectColorFormatAndExtractColor(colorText, colorFormat)
  If colorLng < 0 Then
    ' Color format invalid
    CommandIfColorThenGoTo = False
    RaiseError MODULE_NAME & "CommandIfColorThenGoTo", Err.Number, Err.Source, _
      "Color needs to be valid color code: Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  GetCursorPos MousePos
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))

  tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
  ' Check if color matches within tolerance
  If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
    ' Color matches within tolerance
    GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 3)), "Arg4"
  Else
    ' Color does not matches within tolerance
    GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 4)), "Arg5"
  End If

NormalExecution:
  CommandIfColorThenGoTo = True
  Exit Function
eh:
  CommandIfColorThenGoTo = False
  RaiseError MODULE_NAME & ".CommandIfColorThenGoTo", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function CommandIfWaitColorThenSkip(Optional ExecutingTroughApplicationRun As Boolean = False)
' Arg1: X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position
' Arg2: Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position
' Arg3: Color of the pixel in Hex format; if missing or invalid, then stop with error
' Arg4: How many lines to be skipped if the color matches
' Arg5: How many lines to be skipped if the color does not matches

  On Error GoTo eh

  Static colorText As String: colorText = Trim(CStr(currentRowArray(1, ColAArg1 + 2)))
  If Len(colorText) = 0 Then
    CommandIfColorThenSkip = False
    RaiseError MODULE_NAME & "CommandIfWaitColorThenSkip", Err.Number, Err.Source, _
      "Color needs to be entered in Arg3", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Static colorFormat As Byte
  Static colorLng As Long

  colorLng = DetectColorFormatAndExtractColor(colorText, colorFormat)
  If colorLng < 0 Then
    ' Color format invalid
    CommandIfColorThenSkip = False
    RaiseError MODULE_NAME & "CommandIfWaitColorThenSkip", Err.Number, Err.Source, _
      "Color needs to be valid color code: Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  GetCursorPos MousePos
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))

  Static timeToCheckAndWait As Long: timeToCheckAndWait = colorCheckMax
  Do Until timeToCheckAndWait <= 0
    tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
    ' Check if color matches within tolerance
    If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
    
      ' Color matches within tolerance
      SkipLines CStr(currentRowArray(1, ColAArg1 + 3)), "Arg4"
      GoTo NormalExecution

    End If
    Sleep minLong(timeToCheckAndWait, colorCheckSplit)
    timeToCheckAndWait = timeToCheckAndWait - colorCheckSplit
    DoEvents
  Loop

  ' Color does not matches within tolerance
  SkipLines CStr(currentRowArray(1, ColAArg1 + 4)), "Arg5"

NormalExecution:
  CommandIfColorThenSkip = True
  Exit Function
eh:
  CommandIfColorThenSkip = False
  RaiseError MODULE_NAME & ".CommandIfWaitColorThenSkip", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function CommandIfWaitColorThenGoTo(Optional ExecutingTroughApplicationRun As Boolean = False)
' Arg1: X-value of the pixel to retrieve the color; if missing or invalid, then use current mouse X-Position
' Arg2: Y-value of the pixel to retrieve the color; if missing or invalid, then use current mouse Y-Position
' Arg3: Color of the pixel in Hex format; if missing or invalid, then stop with error
' Arg4: What line or label should be next if the color matches
' Arg5: What line or label should be next if the color does not matches

  On Error GoTo eh

  Static colorText As String: colorText = Trim(CStr(currentRowArray(1, ColAArg1 + 2)))
  If Len(colorText) = 0 Then
    CommandIfColorThenGoTo = False
    RaiseError MODULE_NAME & "CommandIfWaitColorThenGoTo", Err.Number, Err.Source, _
      "Color needs to be entered in Arg3", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  Static colorFormat As Byte
  Static colorLng As Long

  colorLng = DetectColorFormatAndExtractColor(colorText, colorFormat)
  If colorLng < 0 Then
    ' Color format invalid
    CommandIfColorThenGoTo = False
    RaiseError MODULE_NAME & "CommandIfWaitColorThenGoTo", Err.Number, Err.Source, _
      "Color needs to be valid color code: Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

  GetCursorPos MousePos
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))

  Static timeToCheckAndWait As Long: timeToCheckAndWait = colorCheckMax
  Do Until timeToCheckAndWait <= 0
    tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
    ' Check if color matches within tolerance
    If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
    
      ' Color matches within tolerance
      GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 3)), "Arg4"
      GoTo NormalExecution

    End If
    Sleep minLong(timeToCheckAndWait, colorCheckSplit)
    timeToCheckAndWait = timeToCheckAndWait - colorCheckSplit
    DoEvents
  Loop

  ' Color does not matches within tolerance
  GotoLineOrLabel CStr(currentRowArray(1, ColAArg1 + 4)), "Arg5"

NormalExecution:
  CommandIfColorThenGoTo = True
  Exit Function
eh:
  CommandIfColorThenGoTo = False
  RaiseError MODULE_NAME & ".CommandIfWaitColorThenGoTo", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function



Public Function WaitColorUnderCursor(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Static colorText As String: colorText = Trim(CStr(currentRowArray(1, ColAColor)))
  If Len(colorText) = 0 Then
    WaitColorUnderCursor = True
    Exit Function
  End If
  Static colorFormat As Byte
  Static colorLng As Long

  colorLng = DetectColorFormatAndExtractColor(colorText, colorFormat)
  If colorLng < 0 Then
    ' Color format invalid
    WaitColorUnderCursor = False
    Exit Function
  End If

  ' Get cursor position
  GetCursorPos MousePos
  If IsValidClickCommand Then
    If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
    If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))
  End If


  Static timeToCheckAndWait As Long

  If colorReadFirst Then
    timeToCheckAndWait = colorCheckMax
    Do Until timeToCheckAndWait <= 0
      tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
      If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
        WaitColorUnderCursor = True
        Exit Function
      End If
      Sleep minLong(timeToCheckAndWait, colorCheckSplit)
      timeToCheckAndWait = timeToCheckAndWait - colorCheckSplit
      DoEvents
    Loop

    tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
    currentRowRange(1, ColAColor).Value2 = ConvertLongToHex(tmpL)

    WaitColorUnderCursor = True
    Exit Function
  End If

  Do
    timeToCheckAndWait = colorCheckMax
    Do Until timeToCheckAndWait <= 0
      tmpL = GetPixel(WindowDC, MousePos.x, MousePos.y)
      If IsColorWithinTolerance(tmpL, colorLng, maxColorTolerance) Then
        WaitColorUnderCursor = True
        Exit Function
      End If
      Sleep minLong(timeToCheckAndWait, colorCheckSplit)
      timeToCheckAndWait = timeToCheckAndWait - colorCheckSplit
      DoEvents
    Loop
    
    ' Convert detected color to user input format
    Select Case colorFormat
      Case colorFormatLong:
        tmpS = CStr(tmpL)
      Case colorFormatRGB:
        tmpS = ConvertLongToRGB(tmpL)
      Case colorFormatHEX:
        tmpS = ConvertLongToHex(tmpL)
    End Select
    
    ' Prompt user with detected color
    Select Case AskNextStep("Color of pixel (" & MousePos.x & " ," & MousePos.y & ") is expected to be " & colorText & " but is " & tmpS & ". Do you want to retry the checking?", vbAbortRetryIgnore, "Color does not match")
      Case vbIgnore:
        WaitColorUnderCursor = True
        Exit Function
      Case vbAbort:
        WaitColorUnderCursor = False
        Exit Function
      Case vbRetry:
        ' Check again
    End Select
  Loop While True

  WaitColorUnderCursor = False
  Exit Function
eh:
  WaitColorUnderCursor = False
  RaiseError MODULE_NAME & ".WaitColorUnderCursor", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Private Function DetectColorFormatAndExtractColor(ByRef colorText As String, ByRef format As Byte) As Long
  If IsNumeric(colorText) Then
    ' Format numeric Long (existent)
    DetectColorFormatAndExtractColor = CLng(colorText)
    format = colorFormatLong
  ' Check the limits
    If DetectColorFormatAndExtractColor < 0 Or DetectColorFormatAndExtractColor > 16777215 Then
      DetectColorFormatAndExtractColor = -1
      Exit Function
    End If
  ElseIf Left(colorText, 1) = "#" Then
    ' Format hexazecimal #RRGGBB sau #RRGGBBAA (we ignor AA)
    DetectColorFormatAndExtractColor = ConvertHexToLong(colorText)
    format = colorFormatHEX
  ElseIf colorText Like "*[0-9]*,*[0-9]*,*[0-9]*" Or colorText Like "*[0-9]* [0-9]* [0-9]*" Then
    ' Format R,G,B or R G B
    DetectColorFormatAndExtractColor = ConvertRGBToLong(colorText)
    format = colorFormatRGB
  Else
    ' Unknown format
    DetectColorFormatAndExtractColor = -1
  End If
End Function
Private Function ConvertRGBToLong(colorText As String) As Long
  Dim parts() As String
  Dim r As Long, g As Long, b As Long

  ' Detect separator: space or comma
  If InStr(colorText, ",") > 0 Then
    parts = Split(colorText, ",")
  Else
    parts = Split(colorText, " ")
  End If

  ' Check if we have exactly 3 components
  If UBound(parts) <> 2 Then
    ConvertRGBToLong = -1
    Exit Function
  End If

  On Error GoTo wrongFormat
  r = CLng(Trim(parts(0)))
  g = CLng(Trim(parts(1)))
  b = CLng(Trim(parts(2)))
  On Error GoTo 0

  ' Check the limits
  If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then
    ConvertRGBToLong = -1
    Exit Function
  End If

  ' Convert RGB to Long format
  ConvertRGBToLong = RGB(r, g, b)
  Exit Function
wrongFormat:
  ConvertRGBToLong = -1
End Function

Private Function ConvertHexToLong(colorText As String) As Long
  Dim r As Long, g As Long, b As Long
  Dim hexColor As String

  ' Remove # if present
  hexColor = Replace(colorText, "#", "")

  ' Convert to uppercase to ensure case insensitivity
  hexColor = UCase(hexColor)

  ' Expand short hex formats (e.g., #FFF ? #FFFFFF)
  Select Case Len(hexColor)
    Case 3  ' #RGB ? #RRGGBB
      hexColor = Mid(hexColor, 1, 1) & Mid(hexColor, 1, 1) & _
                 Mid(hexColor, 2, 1) & Mid(hexColor, 2, 1) & _
                 Mid(hexColor, 3, 1) & Mid(hexColor, 3, 1)
    Case 4  ' #RGBA ? #RRGGBB (ignoring A)
      hexColor = Mid(hexColor, 1, 1) & Mid(hexColor, 1, 1) & _
                 Mid(hexColor, 2, 1) & Mid(hexColor, 2, 1) & _
                 Mid(hexColor, 3, 1) & Mid(hexColor, 3, 1)
    Case 6  ' #RRGGBB ? No change
      ' Already in correct format
    Case 8  ' #RRGGBBAA ? Ignore alpha (last two characters)
      hexColor = Left(hexColor, 6)
    Case Else
      ' Invalid length, return error (-1)
      ConvertHexToLong = -1
      Exit Function
  End Select

  ' Ensure the final length is exactly 6 characters
  If Len(hexColor) <> 6 Then
    ConvertHexToLong = -1
    Exit Function
  End If

  On Error GoTo wrongFormat
  ' Convert hex components to decimal values
  r = CLng("&H" & Mid(hexColor, 1, 2))
  g = CLng("&H" & Mid(hexColor, 3, 2))
  b = CLng("&H" & Mid(hexColor, 5, 2))
  On Error GoTo 0

  ' Convert RGB to Long format
  ConvertHexToLong = RGB(r, g, b)
  Exit Function
wrongFormat:
  ConvertHexToLong = -1
End Function

Private Function ConvertLongToRGB(colorLng As Long) As String
  Dim r As Long, g As Long, b As Long

  ' Extract RGB components from Long format
  r = colorLng And 255           ' Extract Red (Least Significant Byte)
  g = (colorLng \ 256) And 255   ' Extract Green
  b = (colorLng \ 65536) And 255 ' Extract Blue

  ' Return in "R,G,B" format
  ConvertLongToRGB = CStr(r) & "," & CStr(g) & "," & CStr(b)
End Function

Private Function ConvertLongToHex(colorLng As Long) As String
  Dim r As Long, g As Long, b As Long

  ' Extract RGB components from Long format
  r = colorLng And 255           ' Extract Red (Least Significant Byte)
  g = (colorLng \ 256) And 255   ' Extract Green
  b = (colorLng \ 65536) And 255 ' Extract Blue

  ' Format as a hex string
  ConvertLongToHex = "#" & Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function

Private Function IsColorWithinTolerance(detectedColor As Long, expectedColor As Long, tolerance As Long) As Boolean
  Dim rD As Long, gD As Long, bD As Long
  Dim rE As Long, gE As Long, bE As Long

  ' Extract RGB components from detected color
  rD = detectedColor And 255
  gD = (detectedColor \ 256) And 255
  bD = (detectedColor \ 65536) And 255

  ' Extract RGB components from expected color
  rE = expectedColor And 255
  gE = (expectedColor \ 256) And 255
  bE = (expectedColor \ 65536) And 255

  ' Check if detected color is within the allowed range
  If rD >= rE - tolerance And rD <= rE + tolerance And _
     gD >= gE - tolerance And gD <= gE + tolerance And _
     bD >= bE - tolerance And bD <= bE + tolerance Then
    IsColorWithinTolerance = True
  Else
    IsColorWithinTolerance = False
  End If
End Function


Public Function GetColorUnderCursor(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: Here will be written the x-value of the mouse position
' Arg2: Here will be written the y-value of the mouse position
' Arg3: Here will be written the color of the pixel in Hex format
  On Error GoTo eh
  GetCursorPos MousePos
  currentRowRange(1, ColAArg1 + 0).Value = MousePos.x
  currentRowRange(1, ColAArg1 + 1).Value = MousePos.y
  currentRowRange(1, ColAArg1 + 2).Value = ConvertLongToHex(GetPixel(WindowDC, MousePos.x, MousePos.y))

done:
  GetColorUnderCursor = True
  Exit Function
eh:
  GetColorUnderCursor = False
  RaiseError MODULE_NAME & ".GetColorUnderCursor", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function GetColorFromPoint(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1: X-value of the pixel to retrieve the color
' Arg2: Y-value of the pixel to retrieve the color
' Arg3: Here will be written the color of the pixel in Hex format
  On Error GoTo eh
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Or IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then
    GetCursorPos MousePos
    If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) Then MousePos.x = CLng(currentRowArray(1, ColAArg1 + 0))
    If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then MousePos.y = CLng(currentRowArray(1, ColAArg1 + 1))
    currentRowRange(1, ColAArg1 + 2).Value = ConvertLongToHex(GetPixel(WindowDC, MousePos.x, MousePos.y))
  Else
    GetColorFromPoint = False
    RaiseError MODULE_NAME & ".GetColorFromPoint", Err.Number, Err.Source, "Please enter the pixel coordinates as numbers: x=Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] y=Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

done:
  GetColorFromPoint = True
  Exit Function
eh:
  GetColorFromPoint = False
  RaiseError MODULE_NAME & ".GetColorFromPoint", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function GetColorUnderCursorAsHex() As String
  On Error GoTo eh
  GetCursorPos MousePos
  GetColorUnderCursorAsHex = ConvertLongToHex(GetPixel(WindowDC, MousePos.x, MousePos.y))

done:
  Exit Function
eh:
  GetColorUnderCursorAsHex = ""
  RaiseError MODULE_NAME & ".GetColorUnderCursorAsHex", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function GetColorFromPointAsHex(x As Long, y As Long) As String
  On Error GoTo eh
  GetColorFromPointAsHex = ConvertLongToHex(GetPixel(WindowDC, x, y))

done:
  Exit Function
eh:
  GetColorFromPointAsHex = ""
  RaiseError MODULE_NAME & ".GetColorFromPointAsHex", Err.Number, Err.Source, Err.Description, Erl
End Function


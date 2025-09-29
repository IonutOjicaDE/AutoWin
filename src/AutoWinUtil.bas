Attribute VB_Name = "AutoWinUtil"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinUtil"
Private tmpL As Long
#If Win64 Then
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#End If

'#############################################
Public Sub StatusUpdate(s As String)
  On Error GoTo eh
'Public Const StatusNOW   As String = ">"
'Public Const StatusOK    As String = "+" ' alternative ChrW(9786) 'HEX 263A = DEC 9786
'Public Const StatusNOK   As String = "!"
'Public Const StatusSKIP  As String = "-"
'Public Const StatusPause As String = "p"

  Static tmpS As String, tmpL As Long
  tmpS = currentRowArray(1, ColAStatus)
  If Len(tmpS) <= 1 Then
    If s = StatusNOW Then currentRowArray(1, ColAStatus) = s & "0" Else currentRowArray(1, ColAStatus) = s & "1"
  Else
    tmpS = Mid(tmpS, 2)
    If IsNumber(tmpS) Then
      If s = StatusOK Or s = StatusSKIP Then tmpL = tmpS + 1 Else tmpL = tmpS
      currentRowArray(1, ColAStatus) = s & tmpL
    Else
      If s = StatusNOW Then currentRowArray(1, ColAStatus) = s & "0" Else currentRowArray(1, ColAStatus) = s & "1"
    End If
  End If

  Static lActiveRow As Long
  If currentRow = lActiveRow Then
    ApplicationStatusBar = s & " " & Mid(ApplicationStatusBar, 3)
  Else
    ApplicationStatusBar = s & " " & currentRow
    If Len(currentRowArray(1, ColACommand)) <> 0 Then
      ApplicationStatusBar = ApplicationStatusBar & ": " & currentRowArray(1, ColACommand)
      tmpS = "("
      For tmpL = 0 To 9
        If Len(currentRowArray(1, ColAArg1 + tmpL)) <> 0 Then
          If Len(tmpS) = 1 Then tmpS = tmpS & ", " & Mid(currentRowArray(1, ColAArg1 + tmpL), 1, 10) Else tmpS = tmpS & ", "
        End If
      Next
      ApplicationStatusBar = ApplicationStatusBar & tmpS & ")"
    End If
  End If
  currentRowRange(1, ColAStatus).Value = currentRowArray(1, ColAStatus)
  Application.StatusBar = ApplicationStatusBar

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".StatusUpdate", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub ClearStatusColumn()
  On Error GoTo eh
  tmpL = GetLastLineOnColumn(shAuto, ColAStatus)
  If tmpL > startRow Then
    With shAuto.Cells(startRow, ColAStatus).Resize(tmpL, 1)
      .ClearContents
      .NumberFormat = "@"
    End With
  End If

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".ClearStatusColumn", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub ClearKeyPressColumn()
  On Error GoTo eh
  tmpL = GetLastLineOnColumn(shKey, ColKeyName)
  If tmpL > 1 Then
    With shKey.Cells(1, ColKeyPressed).Resize(GetLastLineOnColumn(shKey, ColKeyName), 1)
      .ClearContents
      .NumberFormat = "@"
    End With
  End If

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".ClearStatusColumn", Err.Number, Err.Source, Err.description, Erl
End Sub

'#############################################
Public Function MySleep(ByRef ms As Long) As Boolean
'valoarea returnata = MouseMoved
  On Error GoTo eh
  Dim msCurrent As Long: msCurrent = ms
  Do Until msCurrent <= 0
    Application.StatusBar = ApplicationStatusBar & " " & String((ms - msCurrent) / waitTimeSplit, "#") & String(msCurrent / waitTimeSplit, "~")
    Sleep minLong(msCurrent, waitTimeSplit)
    msCurrent = msCurrent - waitTimeSplit
    DoEvents
    If MouseMovedByUser Then
      MySleep = True
      Exit Function
    End If
  Loop
  MySleep = False

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".MySleep", Err.Number, Err.Source, Err.description, Erl
End Function

'#############################################
'Public Sub Dep(s As String, Optional nivel As Byte = 1, Optional s2 As String, Optional nivel2 As Byte = 2)
'  On Error GoTo eh
'  Static Text As String
'  Text = vbNullString
'  If debugLevel >= nivel Then Text = " §" & nivel & ": " & s
'  If debugLevel >= nivel2 And Len(s2) > 0 Then Text = Text & " §" & nivel2 & ": " & s2
'  If Len(Text) > 0 Then
'    Debug.Print Now & Text
'    'Application.StatusBar = Text
'  End If
'
'done:
'  Exit Sub
'eh:
'  RaiseError MODULE_NAME & ".Dep", Err.Number, Err.Source, Err.description, Erl
'End Sub
'Public Sub EndWithMessage(text As String)
'  MsgBox text, vbOKOnly
'  End
'End Sub
'#############################################
'Public Function getCellColumn(ByRef text As String) As Long
'  On Error GoTo eh
'  Dim tmpR As Range
'  Set tmpR = shAuto.Range("1:1").Find(text, LookIn:=xlValues, LookAt:=xlPart)
'  If tmpR Is Nothing Then EndWithMessage "Column name """ & text & """ not found. Terminating the application."
'  getCellColumn = tmpR.Column
'
'done:
'  Exit Function
'eh:
'  RaiseError MODULE_NAME & ".getCellColumn", Err.Number, Err.Source, Err.description, Erl
'End Function
'Public Sub PrintError(s As String)
'  If Err.Number <> 0 Then
'    Debug.Print s & ": " & Err.Source & ": " & Err.Number & ": " & Err.description
'    Err.Clear
'  End If
'End Sub

'#############################################
Public Function BetweenLng(minL As Long, checked As Long, maxL As Long) As Boolean
  BetweenLng = (checked >= minL) And (checked <= maxL)
End Function
Public Function inRangeLng(checked As Long, center As Long, delta As Long) As Boolean
  On Error GoTo eh
  inRangeLng = Abs(checked - center) <= delta

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".inRangeLng", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function minLong(a As Long, b As Long) As Long
  minLong = IIf(a < b, a, b)
End Function
Public Function maxLong(a As Long, b As Long) As Long
  maxLong = IIf(a > b, a, b)
End Function
Public Function AreSamePointsLong(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long, ByRef tol As Long) As Boolean
  On Error GoTo eh
  AreSamePointsLong = inRangeLng(x1, x2, tol) And inRangeLng(y1, y2, tol)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".AreSamePointsLong", Err.Number, Err.Source, Err.description, Erl
End Function


'#############################################
Public Function BetweenInt(minI As Integer, checked As Integer, maxI As Integer) As Boolean
  BetweenInt = (checked >= minI) And (checked <= maxI)
End Function
Public Function inRangeInt(checked As Integer, center As Integer, delta As Integer) As Boolean
  On Error GoTo eh
  inRangeInt = Abs(checked - center) <= delta

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".inRangeInt", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function minInt(a As Integer, b As Integer) As Integer
  minInt = IIf(a < b, a, b)
End Function
Public Function maxInt(a As Integer, b As Integer) As Integer
  maxInt = IIf(a > b, a, b)
End Function
Public Function AreSamePointsInt(ByRef x1 As Integer, ByRef y1 As Integer, ByRef x2 As Integer, ByRef y2 As Integer, ByRef tol As Integer) As Boolean
  On Error GoTo eh
  AreSamePointsInt = inRangeInt(x1, x2, tol) And inRangeInt(y1, y2, tol)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".AreSamePointsInt", Err.Number, Err.Source, Err.description, Erl
End Function


'#############################################
Public Function BetweenDouble(minD As Double, checked As Double, maxD As Double) As Boolean
  BetweenDouble = (checked >= minD) And (checked <= maxD)
End Function
Public Function inRangeDouble(ByRef checked As Double, ByRef center As Double, ByRef delta As Double) As Boolean
  On Error GoTo eh
  inRangeDouble = Abs(checked - center) <= delta

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".inRangeDouble", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function minDouble(ByRef a As Double, ByRef b As Double) As Double
  minDouble = IIf(a < b, a, b)
End Function
Public Function maxDouble(ByRef a As Double, ByRef b As Double) As Double
  maxDouble = IIf(a > b, a, b)
End Function
Public Function AreSamePointsDouble(ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double, ByRef tol As Double) As Boolean
  On Error GoTo eh
  AreSamePointsDouble = inRangeDouble(x1, x2, tol) And inRangeDouble(y1, y2, tol)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".AreSamePointsDouble", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function RoundUp(ByRef d As Double) As Long
  On Error GoTo eh
  RoundUp = Round(d + 0.499999999999999)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RoundUp", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function RoundDown(ByRef d As Double) As Long
  On Error GoTo eh
  RoundDown = Round(d - 0.5)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RoundDown", Err.Number, Err.Source, Err.description, Erl
End Function


'#############################################
Public Function IsNumber(ByRef a As String) As Boolean
  IsNumber = (Len(a) > 0) And IsNumeric(a)
End Function


Public Function GetBoolean(ByRef s As String) As Boolean
  If IsNumber(s) Then
    GetBoolean = Val(s) <> 0
  Else
    GetBoolean = InStr(1, "true wahr adevarat yes ja da", s, vbTextCompare) > 0
  End If
End Function

Public Function ArraySize(Matrix As Variant, Optional Dimension As Long = 1&) As Long
  On Error GoTo eh
  If IsArray(Matrix) Then
    If Dimension = 1& Then
      ArraySize = UBound(Matrix) - LBound(Matrix) + 1&
    Else
      On Error Resume Next
      ArraySize = UBound(Matrix, Dimension) - LBound(Matrix, Dimension) + 1
      If Err.Number <> 0 Then
        ArraySize = -1&
        Err.Clear
      End If
    End If
    GoTo done
  End If
  ArraySize = -1&

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".ArraySize", Err.Number, Err.Source, Err.description, Erl
End Function

'#############################################
Public Function GetLastLine(ByRef sh As Worksheet) As Long
  If sh Is Nothing Then
    GetLastLine = 0&
  Else
    Dim c As Range: Set c = sh.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If c Is Nothing Then GetLastLine = 0& Else GetLastLine = c.Row
  End If
End Function
Public Function GetLastLineOnColumn(ByRef sh As Worksheet, ByRef ColumnNumber As Long) As Long
  If sh Is Nothing Then
    GetLastLineOnColumn = 0&
  Else
    If ColumnNumber > 0 And ColumnNumber <= sh.Columns.count Then
      Dim c As Range: Set c = sh.Columns(ColumnNumber).Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
      If c Is Nothing Then GetLastLineOnColumn = 0& Else GetLastLineOnColumn = c.Row
    Else
      GetLastLineOnColumn = 0&
    End If
  End If
End Function
Public Function GetLastColumn(ByRef sh As Worksheet) As Long
  If sh Is Nothing Then
    GetLastColumn = 0&
  Else
    Dim c As Range: Set c = sh.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If c Is Nothing Then GetLastColumn = 0& Else GetLastColumn = c.Column
  End If
End Function
Public Function GetLastColumnOnLine(ByRef sh As Worksheet, ByRef RowNumber As Long) As Long
  If sh Is Nothing Then
    GetLastColumnOnLine = 0&
  Else
    If RowNumber > 0 And RowNumber <= sh.Rows.count Then
      Dim c As Range: Set c = sh.Rows(RowNumber).Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
      If c Is Nothing Then GetLastColumnOnLine = 0& Else GetLastColumnOnLine = c.Column
    Else
      GetLastColumnOnLine = 0&
    End If
  End If
End Function


'#############################################
Public Function nTrim(ByRef s As String) As String
  Dim iPos As Long: iPos = InStr(s, Chr$(0))
  If iPos > 0 Then nTrim = Left$(s, iPos - 1) Else nTrim = s
End Function
Public Function CleanString(ByRef s As String) As String
  CleanString = LCase$(Replace$(s, " ", ""))
End Function


'#############################################
Public Sub StatusBarScheduleToRevert(Optional ByRef AfterSeconds As Long = 5)
  Application.OnTime Now + TimeSerial(0, 0, minLong(Abs(AfterSeconds), 59)), "StatusBarRevert"
End Sub
Public Sub StatusBarRevert()
  Application.StatusBar = False
End Sub


'#############################################
#If Win64 Then
  Public Function PointToLongLong(point As POINTAPI) As LongLong
' https://stackoverflow.com/questions/1070863/hidden-features-of-vba
' https://www.ms-office-forum.net/forum/showpost.php?s=825aa2b0304c7cf0e27bcabd7cfcdfde&p=1864214&postcount=9
    CopyMemory PointToLongLong, point, 8
  End Function
#End If

Attribute VB_Name = "ExcelSheetUtilites"

Public Sub CalculateNow()
  'Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.CalculateFull
  'Application.Calculation = xlCalculationAutomatic
  DoEvents
  DoEvents
  Application.ScreenUpdating = True
End Sub

Sub ShapePlacementMove()
' change all Shapes inside active Document Placement to xlMove ("Move but dont size with cells.")
' Ionut Ojica 15.06.2018
  Dim shBild As Shape
  Dim shSheet As Worksheet
  For Each shSheet In ActiveWorkbook.Worksheets
    For Each shBild In shSheet.Shapes
      'If shBild.DrawingObject.ShapeRange.Type = msoPicture (or msoGroup) Then
      shBild.Placement = xlMove ' xlFreeFloating, xlMove , xlMoveAndSize
    Next shBild
  Next shSheet
End Sub

Sub ListFormats()
' lists the cell format templates
' http://www.office-loesung.de/ftopic627990_0_0_asc.php
  Application.ScreenUpdating = False
'  Application.EnableEvents = False
  MsgBox (ActiveWorkbook.Styles.count)
  Dim i As Integer
  For i = ActiveWorkbook.Styles.count To 1 Step -1
    ActiveSheet.Cells(i, 1).Value = ActiveWorkbook.Styles(i).Name
  Next
  Application.ScreenUpdating = True
'  Application.EnableEvents = true
End Sub
Sub DeleteFormats()
' delete certain cell format templates
' http://www.office-loesung.de/ftopic627990_0_0_asc.php
  Application.ScreenUpdating = False
'  Application.EnableEvents = False
  Dim i As Integer
  Finish = 10 ' Replace Finish and Start with numbers
  Start = 1   ' Replace Finish and Start with numbers
  MsgBox (ActiveWorkbook.Styles.count)
  For i = Finish To Start Step -1
    With ActiveWorkbook
    .Styles(i).Delete
    End With
  Next
  MsgBox (ActiveWorkbook.Styles.count)
  Application.ScreenUpdating = True
'  Application.EnableEvents = true
End Sub

Sub DeleteAllCustomFormats()
'to be tested
  Application.ScreenUpdating = False
'  Application.EnableEvents = False
  Dim i As style
  On Error Resume Next ' the predefined Formats cannot be deleted
  MsgBox (ActiveWorkbook.Styles.count)
  For Each i In ActiveWorkbook.Styles
    i.Delete
  Next
  MsgBox (ActiveWorkbook.Styles.count)
  Application.ScreenUpdating = True
'  Application.EnableEvents = true
End Sub

Sub DeleteNumberFormats()
'to be tested; it takes too long
'http://www.office-loesung.de/ftopic530067_0_0_asc.php
  Dim NuFo As Object
  Dim tmpCell As Range
  Dim ws As Worksheet
  Dim i As Long
  Dim NF

  Set NuFo = CreateObject("scripting.dictionary")

  For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    Application.StatusBar = "Reading Number Format, Sheet " & i & " from " & ThisWorkbook.Worksheets.count
    For Each tmpCell In ws.UsedRange
      NuFo(tmpCell.NumberFormat) = 0
    Next
  Next

  Application.StatusBar = "Remove Number Format"
  i = 0
  For Each NF In NuFo.Keys
    On Error Resume Next
    ThisWorkbook.DeleteNumberFormat NumberFormat:=NF
    If Err = 0 Then i = i + 1
    On Error GoTo 0
  Next
  MsgBox i & " custom Number Formats removed."
  Application.StatusBar = False
End Sub

Public Function EvaluateString(strTextString As String, dummy As Range)
' https://exceloffthegrid.com/turn-string-formula-with-evalute/
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.volatile
  Application.Volatile 'This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell => dummy
  EvaluateString = Evaluate(strTextString)
End Function

Public Function EvaluateFunctionX(FunctionText As String, Xreplace As Range)
' https://exceloffthegrid.com/turn-string-formula-with-evalute/
' https://docs.microsoft.com/en-us/office/vba/api/excel.application.volatile
  Application.Volatile 'This method has no effect if it's not inside a user-defined function used to calculate a worksheet cell => dummy
  FunctionText = Replace(FunctionText, "x", Xreplace.Address, compare:=vbTextCompare)
  EvaluateFunctionX = Evaluate(FunctionText)
End Function

Sub Dename()
' Replace range names with cell references
' First select the cells to run onto
' https://www.excelforum.com/excel-formulas-and-functions/431148-replace-range-names-with-cell-references.html
  Dim Cell As Range
  ActiveSheet.TransitionFormEntry = True
  For Each Cell In Selection.SpecialCells(xlFormulas)
    Cell.Formula = Cell.Formula
  Next
  ActiveSheet.TransitionFormEntry = False
End Sub

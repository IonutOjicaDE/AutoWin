Attribute VB_Name = "CommandsOfficeExcel"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsOfficeExcel"

Public WorkingWorkbook      As Workbook
Public WorkingSheet         As Worksheet
Public WorkingCell          As Range

Private CountVisibleRows    As Long
Private CountVisibleColumns As Long
Private tmpS                As String


Public Sub RegisterCommandsOfficeExcel()
  On Error GoTo eh
  CountVisibleRows = ActiveWindow.VisibleRange.Rows.count
  CountVisibleColumns = ActiveWindow.VisibleRange.Columns.count
  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "openexceldocument", Array("OpenExcelDocument", "Open Excel Document", _
    MODULE_NAME, "Open an Excel document", _
    "Filename", "Filename including path and extension to be opened is to be provided")
    
  commandMap.Add "openexceldocumentwithdialog", Array("OpenExcelDocumentWithDialog", "Open Excel Document with Dialog", _
    MODULE_NAME, "Open an Excel document using a browsing window", _
    "Filename", "Filename (including path and extension) of the chosen file will be written here")
    
  commandMap.Add "setworkingexceldocument", Array("SetWorkingExcelDocument", "Set Working Excel Document", _
    MODULE_NAME, "Set as a working document a specified Excel document", "Excel document title", "Title of the Excel document. The document must be already opened. Can be inserted only a part of the title.")
    
  commandMap.Add "setworkingexcelsheet", Array("SetWorkingExcelSheet", "Set Working Excel Sheet", _
    MODULE_NAME, "Set a working worksheet from the specified working Excel document", "Name of Worksheet", "Name of the Worksheet. Can be inserted only a part of the title.")
    
  commandMap.Add "setworkingexcelcell", Array("SetWorkingExcelCell", "Set Working Excel Cell", _
    MODULE_NAME, "Set a working cell from the working sheet", "Row", "Row of the cell to activate", "Column", "Column of the cell to activate")
  
  
  commandMap.Add "copycelltoclipboard", Array("CopyCellToClipboard", "Copy Cell To Clipboard", _
    MODULE_NAME, "Copy a cell from the working sheet in the clipboard", "Row", "Row of the cell to copy", "Column", "Column of the cell to copy")
    
  commandMap.Add "writetextincell", Array("WriteTextInCell", "Write Text In Cell", _
    MODULE_NAME, "Write a text inside a cell", "Cell address", "Cell address in form B2 or C7 within working sheet", "Text", "Text to write in the cell")
    
  commandMap.Add "writeformulaincell", Array("WriteFormulaInCell", "Write Formula In Cell", _
    MODULE_NAME, "Write a formula inside a cell", "Cell address", "Cell address in form B2 or C7 within working sheet", "Formula", "Formula to write in the cell")

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsOfficeExcel", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsOfficeExcel()
  On Error GoTo eh

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsOfficeExcel", Err.Number, Err.Source, Err.description, Erl
End Sub


Public Function OpenExcelDocument(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=complete path of Excel document
  On Error GoTo eh
  Set WorkingWorkbook = Workbooks.Open(currentRowArray(1, ColAArg1))

done:
  OpenExcelDocument = True
  Exit Function
eh:
  OpenExcelDocument = False
  RaiseError MODULE_NAME & ".OpenExcelDocument", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function OpenExcelDocumentWithDialog(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' https://powerspreadsheets.com/vba-open-workbook/
' Arg1=write the complete path of opened Excel document
  On Error GoTo eh
  Dim t
  t = Application.GetOpenFilename(FileFilter:="Excel Files,*.xl*;*.xm*")
  If t <> False Then
    Set WorkingWorkbook = Workbooks.Open(t)
    currentRowRange(1, ColAArg1).Value = t
  Else
    OpenExcelDocumentWithDialog = False
    RaiseError MODULE_NAME & ".OpenExcelDocumentWithDialog", Err.Number, Err.Source, _
      "No document was chosen.", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

done:
  OpenExcelDocumentWithDialog = True
  Exit Function
eh:
  OpenExcelDocumentWithDialog = False
  RaiseError MODULE_NAME & ".OpenExcelDocumentWithDialog", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function SetWorkingExcelDocument(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=(partial) title of Excel document. The document must be already opened
  On Error GoTo eh
  tmpS = currentRowArray(1, ColAArg1)
  Dim t As Workbook
  For Each t In Workbooks
    If t.Name Like tmpS Then
      Set WorkingWorkbook = t
      SetWorkingExcelDocument = True
      Exit Function
    End If
  Next

  SetWorkingExcelDocument = False
  RaiseError MODULE_NAME & "SetWorkingExcelDocument", Err.Number, Err.Source, _
    "No Excel Document found, that contains: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] in its name.", Erl, 1, ExecutingTroughApplicationRun
  Exit Function
eh:
  SetWorkingExcelDocument = False
  RaiseError MODULE_NAME & ".SetWorkingExcelDocument", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function SetWorkingExcelSheet(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=name of Worksheet
  On Error GoTo eh
  
  If WorkingWorkbook Is Nothing Then
    SetWorkingExcelSheet = False
    RaiseError MODULE_NAME & "SetWorkingExcelSheet", Err.Number, Err.Source, _
      "WorkingWorkbook is not set. Please first use command [Set Working Excel Document]: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  tmpS = currentRowArray(1, ColAArg1)
  Dim t As Worksheet
  For Each t In WorkingWorkbook.Worksheets
    If t.Name Like tmpS Then
      Set WorkingSheet = t
      SetWorkingExcelSheet = True
      Exit Function
    End If
  Next

  SetWorkingExcelSheet = False
  RaiseError MODULE_NAME & "SetWorkingExcelSheet", Err.Number, Err.Source, _
    "No Worksheet found, that contains: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] in its name.", Erl, 2, ExecutingTroughApplicationRun
  Exit Function
eh:
  SetWorkingExcelSheet = False
  RaiseError MODULE_NAME & ".SetWorkingExcelSheet", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function SetWorkingExcelCell(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=row, Arg2=column
  On Error GoTo eh

  If WorkingSheet Is Nothing Then
    SetWorkingExcelCell = False
    RaiseError MODULE_NAME & "SetWorkingExcelCell", Err.Number, Err.Source, _
      "WorkingSheet is not set. Please first use command [Set Working Excel Sheet]: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then
    Set WorkingCell = WorkingSheet.Cells(currentRowArray(1, ColAArg1 + 0), currentRowArray(1, ColAArg1 + 1))
    SetWorkingExcelCell = True
    Exit Function
  End If

  SetWorkingExcelCell = False
  RaiseError MODULE_NAME & "SetWorkingExcelCell", Err.Number, Err.Source, _
    "Arguments need to be valid numbers of row and collumn: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 2, ExecutingTroughApplicationRun
  Exit Function
eh:
  SetWorkingExcelCell = False
  RaiseError MODULE_NAME & ".SetWorkingExcelCell", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
'Public Sub SendKeysFromCell(ByRef r As Range, ByRef MyKeysObj As keyKeys)
'    MyKeysObj.SendKeysToActiveWindow WorkingSheet.Cells(r.Value, r.Offset(0, 1).Value).Text
'End Sub
'Public Sub SendKeysFromWorkingCell(MyKeysObj As keyKeys)
'    MyKeysObj.SendKeysToActiveWindow WorkingCell.Text
'End Sub
Public Function CopyCellToClipboard(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=row, Arg2=column of the cell to copy the content
  On Error GoTo eh

  If WorkingSheet Is Nothing Then
    CopyCellToClipboard = False
    RaiseError MODULE_NAME & "CopyCellToClipboard", Err.Number, Err.Source, _
      "WorkingSheet is not set. Please first use command [Set Working Excel Sheet]: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) Then
    WorkingSheet.Cells(currentRowArray(1, ColAArg1 + 0), currentRowArray(1, ColAArg1 + 1)).Copy
    CopyCellToClipboard = True
    Exit Function
  End If

  CopyCellToClipboard = False
  RaiseError MODULE_NAME & "CopyCellToClipboard", Err.Number, Err.Source, _
    "Arguments need to be valid numbers of row and collumn: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 2, ExecutingTroughApplicationRun
  Exit Function
eh:
  CopyCellToClipboard = False
  RaiseError MODULE_NAME & ".CopyCellToClipboard", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function WriteTextInCell(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=cell name in form B2 or C7 within Working Excel Sheet, Arg2=text to write
  On Error GoTo eh
  WorkingSheet.Range(currentRowArray(1, ColAArg1 + 0)).Value = currentRowArray(1, ColAArg1 + 1)

done:
  WriteTextInCell = True
  Exit Function
eh:
  WriteTextInCell = False
  RaiseError MODULE_NAME & ".WriteTextInCell", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function WriteFormulaInCell(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=cell name in form B2 or C7 within Working Excel Sheet, Arg2=formula to write
  On Error GoTo eh
  WorkingSheet.Range(currentRowArray(1, ColAArg1 + 0)).Formula = currentRowArray(1, ColAArg1 + 1)

done:
  WriteFormulaInCell = True
  Exit Function
eh:
  WriteFormulaInCell = False
  RaiseError MODULE_NAME & ".WriteFormulaInCell", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function CenterViewToCurrentRow(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Application.GoTo shAuto.Cells(maxLong(1, currentRow - CountVisibleRows / 2), ColAStatus), True

done:
  CenterViewToCurrentRow = True
  Exit Function
eh:
  CenterViewToCurrentRow = False
  RaiseError MODULE_NAME & ".CenterViewToCurrentRow", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


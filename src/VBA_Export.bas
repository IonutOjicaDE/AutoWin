Attribute VB_Name = "VBA_Export"
' === Export all VBA components to text files ===
' Requires: Tools > References > Microsoft Visual Basic for Applications Extensibility
' And Excel Option: Trust access to the VBA project object model.
Option Explicit

Public Sub ExportVbaToFolder() '(Optional ByVal targetPath As String)
  ' Determine export folder next to the workbook if not provided
  Dim fso As Object, wbPath As String
  Set fso = CreateObject("Scripting.FileSystemObject")
  wbPath = ThisWorkbook.path
  Dim targetPath As String
  If Len(targetPath) = 0 Then targetPath = wbPath & "\src"
  If Not fso.FolderExists(targetPath) Then fso.CreateFolder targetPath

  Dim comp As VBIDE.VBComponent, filePath As String
  For Each comp In ThisWorkbook.VBProject.VBComponents
    filePath = targetPath & "\" & comp.Name & FileExtFor(comp)
    On Error Resume Next
    Kill filePath ' overwrite if exists
    On Error GoTo 0
    comp.Export filePath
  Next comp

  MsgBox "Exported VBA to: " & targetPath, vbInformation
End Sub

Private Function FileExtFor(ByVal comp As VBIDE.VBComponent) As String
  ' Map component type to file extension
  Select Case comp.Type
    Case vbext_ct_StdModule:    FileExtFor = ".bas"
    Case vbext_ct_ClassModule:  FileExtFor = ".cls"
    Case vbext_ct_MSForm:       FileExtFor = ".frm" ' .frx will be created automatically
    Case vbext_ct_Document:     FileExtFor = ".cls" ' sheet/ThisWorkbook code-behind
    Case Else:                  FileExtFor = ".bas"
  End Select
End Function


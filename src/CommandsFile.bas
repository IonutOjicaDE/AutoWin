Attribute VB_Name = "CommandsFile"
Option Explicit
Private Const MODULE_NAME As String = "CommandsFile"

Private FilesArray()      As String
Private FilesCurrentRow   As Integer
Private CurrentFileName   As Range
Private FoldersArray()    As String
Private FoldersCurrentRow As Integer
Private CurrentFolderName As Range

Private FileSystem As New FileSystemObject

Private tmpS As String

Public Function RegisterCommandsFile()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "foreachfileinfolder", Array("ForEachFileInFolder", "For Each File In Folder", _
    MODULE_NAME, "Loop trough ich file inside specified folder", _
    "Folder", "Specify folder (path) where to look for the files", _
    "Filename", "Here will be written the filenames one after the other on each loop")

  commandMap.Add "foreachfileinfoldernext", Array("ForEachFileInFolderNext", "For Each File In Folder Next", _
    MODULE_NAME, "Here ends the loop")


  commandMap.Add "foreachfolderinfolder", Array("ForEachFolderInFolder", "For Each Folder In Folder", _
    MODULE_NAME, "Loop trough ich folder inside specified folder", _
    "Folder", "Specify folder (path) where to look for the folders", _
    "SubFolder name", "Here will be written the subfolder names one after the other on each loop")

  commandMap.Add "foreachfolderinfoldernext", Array("ForEachFolderInFolderNext", "For Each Folder In Folder Next", _
    MODULE_NAME, "Here ends the loop")

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsFile", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsFile()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsFile", Err.Number, Err.Source, Err.Description, Erl
End Function




Public Function ForEachFileInFolder(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=Folder, Arg2=each File
  On Error GoTo eh
  ReDim FilesArray(100)
  tmpS = currentRowArray(1, ColAArg1)
  tmpS = tmpS & IIf(Right(tmpS, 1) = "\", "", "\")

  FilesCurrentRow = 0
  tmpS = Dir(tmpS & "\*" & "*")
  Do While Len(tmpS) > 0
    FilesArray(FilesCurrentRow) = currentRowArray(1, ColAArg1) & tmpS
    FilesCurrentRow = FilesCurrentRow + 1
    tmpS = Dir
  Loop
  
  If FilesCurrentRow > 0 Then currentRowRange(1, ColAArg1 + 1).Value = FilesArray(0)
  Set CurrentFileName = currentRowRange(1, ColAArg1 + 1)
  FilesCurrentRow = 0

done:
  ForEachFileInFolder = True
  Exit Function
eh:
  RaiseError MODULE_NAME & ".ForEachFileInFolder", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function ForEachFileInFolderNext(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' no Args, takes next File if any and jumps after ForEachFileInFolder
  On Error GoTo eh
  FilesCurrentRow = FilesCurrentRow + 1
  If Len(FilesArray(FilesCurrentRow)) > 0 Then
    CurrentFileName.Value = FilesArray(FilesCurrentRow)
    currentRow = CurrentFileName.Row
  Else

  End If

done:
  ForEachFileInFolderNext = True
  Exit Function
eh:
  RaiseError MODULE_NAME & ".ForEachFileInFolderNext", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function ForEachFolderInFolder(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=Folder, Arg2=each SubFolder
  On Error GoTo eh
  ReDim FoldersArray(100)
  tmpS = currentRowArray(1, ColAArg1)
  tmpS = tmpS & IIf(Right(tmpS, 1) = "\", "", "\")

  FoldersCurrentRow = 0
  Dim SubFolder As Folder
  For Each SubFolder In FileSystem.GetFolder(tmpS).SubFolders
    FoldersArray(FoldersCurrentRow) = SubFolder.path
    FoldersCurrentRow = FoldersCurrentRow + 1
  Next
  
  If FoldersCurrentRow > 0 Then currentRowRange(1, ColAArg1 + 1).Value = FoldersArray(0)
  Set CurrentFolderName = currentRowRange(1, ColAArg1 + 1).Value
  FoldersCurrentRow = 0

done:
  ForEachFolderInFolder = True
  Exit Function
eh:
  RaiseError MODULE_NAME & ".ForEachFolderInFolder", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function ForEachFolderInFolderNext(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' no Args, takes next SubFolder if any and jumps after ForEachFolderInFolder
  On Error GoTo eh
  FoldersCurrentRow = FoldersCurrentRow + 1
  If Len(FoldersArray(FoldersCurrentRow)) > 0 Then
    CurrentFolderName.Value = FoldersArray(FoldersCurrentRow)
    currentRow = CurrentFolderName.Row
  Else
  
  End If

done:
  ForEachFolderInFolderNext = True
  Exit Function
eh:
  RaiseError MODULE_NAME & ".ForEachFolderInFolderNext", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


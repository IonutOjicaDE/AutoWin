Attribute VB_Name = "CommandsOfficeWord"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsOfficeWord"
'Enumerations:
'https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211923(v=office.11)?redirectedfrom=MSDN


Private Enum WdStoryType
  wdMainTextStory = 1&
  wdFootnotesStory = 2&
  wdEndnotesStory = 3&
  wdCommentsStory = 4&
  wdTextFrameStory = 5&
  wdEvenPagesHeaderStory = 6&
  wdPrimaryHeaderStory = 7&
  wdEvenPagesFooterStory = 8&
  wdPrimaryFooterStory = 9&
  wdFirstPageHeaderStory = 10&
  wdFirstPageFooterStory = 11&
  wdFootnoteSeparatorStory = 12&
  wdFootnoteContinuationSeparatorStory = 13&
  wdFootnoteContinuationNoticeStory = 14&
  wdEndnoteSeparatorStory = 15&
  wdEndnoteContinuationSeparatorStory = 16&
  wdEndnoteContinuationNoticeStory = 17&
End Enum

Private Enum WdFindWrap 'for WordApp.Selection.Find
  wdFindStop = 0&
  wdFindContinue = 1&
  wdFindAsk = 2&
End Enum

Private Enum WdReplace 'WordApp.Selection.Find
  wdReplaceNone = 0&
  wdReplaceOne = 1&
  wdReplaceAll = 2&
End Enum

Private Enum WdSaveOptions 'for WordDoc.Close
  wdPromptToSaveChanges = -2&
  wdSaveChanges = -1&
  wdDoNotSaveChanges = 0&
End Enum

Private Enum WdOriginalFormat 'for WordDoc.Close
  wdWordDocument = 0&
  wdOriginalDocumentFormat = 1&
  wdPromptUser = 2&
End Enum

Public WordApp As Word.Application
Public WordDoc As Word.Document


Public Function RegisterCommandsOfficeWord()
  On Error GoTo eh

  Set WordDoc = Nothing
  Set WordApp = Nothing

  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "startword", Array("StartWord", "Start Word", _
    MODULE_NAME, "Start Word application", _
    "Visible", "Enter True if the application should be visible on screen, else the application will be started in background and will not be visible on screen.{{True/False}}")

  commandMap.Add "openworddocument", Array("OpenWordDocument", "Open Word Document", _
    MODULE_NAME, "Open a Word document", _
    "Path", "Filename including path and extension to be opened is to be provided; if Arg2 is mentioned, enter path in Arg1 and filename in Arg2", _
    "Filename", "Eventually filename, and path should be mentioned in Arg1")

  commandMap.Add "closeactiveworddocument", Array("CloseActiveWordDocument", "Close Active Word Document", _
    MODULE_NAME, "Close the active word document", _
    "Save", "Enter True if the document should be saved before closing.{{True/False}}")

  commandMap.Add "wordreplacetext", Array("WordReplaceText", "Word Replace Text", _
    MODULE_NAME, "Replace text inside active Word document", _
    "Current", "Current text version", _
    "New", "New text version")

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsOfficeWord", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsOfficeWord()
  On Error GoTo eh

  Set WordDoc = Nothing
  Set WordApp = Nothing

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsOfficeWord", Err.Number, Err.Source, Err.Description, Erl
End Function



Public Function StartWord(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=make window visible
  On Error Resume Next
  Set WordApp = GetObject("Word.Application") ' if Word is already running
  If Err.Number <> 0 Then ' Launch a new instance of Word
    Err.Clear
    On Error GoTo eh
    Set WordApp = CreateObject("Word.Application")
  Else
    On Error GoTo eh
  End If

  If ExecutingTroughApplicationRun Then _
    If GetBoolean(CStr(currentRowArray(1, ColAArg1))) Then _
      WordApp.Visible = True ' Make the application visible to the user

done:
  StartWord = True
  Exit Function
eh:
  StartWord = False
  RaiseError MODULE_NAME & ".StartWord", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function OpenWordDocument(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1 = Complete path to file, including filename and extension; if Arg2 is mentioned, enter path in Arg1 and filename in Arg2
' Arg2 = eventually filename, and path should be mentioned in Arg1
  On Error GoTo eh
  StartWord

  Dim f As String
  f = currentRowArray(1, ColAArg1)
  If Len(currentRowArray(1, ColAArg1 + 1)) > 0 Then f = f & IIf(Right(f, 1) = "\", "", "\") & currentRowArray(1, ColAArg1 + 1)
  Set WordDoc = WordApp.Documents.Open(Filename:=f)

done:
  OpenWordDocument = True
  Exit Function
eh:
  OpenWordDocument = False
  RaiseError MODULE_NAME & ".OpenWordDocument", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function CloseActiveWordDocument(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=make window visible
'https://docs.microsoft.com/en-us/office/vba/api/word.document.close(method)
  On Error GoTo eh
  WordDoc.Close SaveChanges:=wdSaveChanges, OriginalFormat:=wdOriginalDocumentFormat
  If ExecutingTroughApplicationRun Then
    WordDoc.Close SaveChanges:=IIf(GetBoolean(CStr(currentRowArray(1, ColAArg1))), wdSaveChanges, wdDoNotSaveChanges), OriginalFormat:=wdOriginalDocumentFormat
  Else
    WordDoc.Close SaveChanges:=wdSaveChanges, OriginalFormat:=wdOriginalDocumentFormat
  End If

done:
  CloseActiveWordDocument = True
  Exit Function
eh:
  CloseActiveWordDocument = False
  RaiseError MODULE_NAME & ".CloseActiveWordDocument", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function WordReplaceText(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=original text
' Arg2=new text
'https://wordmvp.com/FAQs/MacrosVBA/FindReplaceAllWithVBA.htm
  On Error GoTo eh
  Dim myStoryRange As Object
  
  'First search the main document using the Selection
  With WordApp.Selection.Find
    .Text = currentRowArray(1, ColAArg1)
    .Replacement.Text = currentRowArray(1, ColAArg1 + 1)
    .Forward = True
    .Wrap = wdFindContinue
    .format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    .Execute Replace:=wdReplaceAll
  End With
  
  'Now search all other stories using Ranges
  For Each myStoryRange In WordDoc.StoryRanges
    If myStoryRange.StoryType <> wdMainTextStory Then
      With myStoryRange.Find
        .Text = currentRowArray(1, ColAArg1)
        .Replacement.Text = currentRowArray(1, ColAArg1 + 1)
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
      End With
      Do While Not (myStoryRange.NextStoryRange Is Nothing)
        Set myStoryRange = myStoryRange.NextStoryRange
        With myStoryRange.Find
          .Text = currentRowArray(1, ColAArg1)
          .Replacement.Text = currentRowArray(1, ColAArg1 + 1)
          .Wrap = wdFindContinue
          .Execute Replace:=wdReplaceAll
        End With
      Loop
    End If
  Next myStoryRange
  
done:
  WordReplaceText = True
  Exit Function
eh:
  WordReplaceText = False
  RaiseError MODULE_NAME & ".WordReplaceText", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

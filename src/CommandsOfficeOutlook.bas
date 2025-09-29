Attribute VB_Name = "CommandsOfficeOutlook"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsOfficeOutlook"


Private Enum ol_Constants 'for OutlookApp.CreateItem https://docs.microsoft.com/en-us/office/vba/api/outlook.olitemtype
  olMailItem = 0&
  olAppointmentItem = 1&
  olContactItem = 2&
  olTaskItem = 3&
  olJournalItem = 4&
  olNoteItem = 5&
  olPostItem = 6&
  olDistributionListItem = 7&
End Enum

Private Enum olFormat_Constants 'for oMail.BodyFormat https://docs.microsoft.com/en-us/office/vba/api/outlook.olbodyformat
  olFormatUnspecified = 0&
  olFormatPlain = 1&
  olFormatHTML = 2&
  olFormatRichText = 3&
End Enum

Private OutlookApp As Outlook.Application
Private oMail As Outlook.MailItem
Private oAccount As Outlook.Account

Private tmpL As Long


Public Sub RegisterCommandsOfficeOutlook()
  On Error GoTo eh

  Set OutlookApp = Nothing
  Set oMail = Nothing
  Set oAccount = Nothing

  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "startoutlook", Array("StartOutlook", "Start Outlook", _
    MODULE_NAME, "Start Outlook application", _
    "Visible", "Enter True if the application should be visible on screen, else the application will be started in background and will not be visible on screen.{{True/False}}")
    
  commandMap.Add "sendemail", Array("SendEmail", "Send Email", _
    MODULE_NAME, "Send a simple email, eventual with attachments", _
    "To", "email address of the recipient or recipients separated by comma", _
    "Subject", "Enter the subject of the email", _
    "Body", "Multiline text is possible for the body of the email", _
    "Attachment1", "Full filename including path of the file to be attached to the email", _
    "Attachment2", "Full filename including path of the file to be attached to the email", _
    "Attachment3", "and so on")

  commandMap.Add "sendemailformated", Array("SendEmailFormated", "Send Email Formated", _
    MODULE_NAME, "Send an email with formated body, eventual with attachments", _
    "To", "email address of the recipient or recipients separated by comma", _
    "Subject", "Enter the subject of the email", _
    "Body", "Full filename including path of the Word file, where the content is saved", _
    "Attachment1", "Full filename including path of the file to be attached to the email", _
    "Attachment2", "Full filename including path of the file to be attached to the email", _
    "Attachment3", "and so on")
    
done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsOfficeOutlook", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsOfficeOutlook()
  On Error GoTo eh

  Set OutlookApp = Nothing
  Set oMail = Nothing
  Set oAccount = Nothing

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsOfficeOutlook", Err.Number, Err.Source, Err.description, Erl
End Sub



Public Function StartOutlook(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=make window visible
  On Error Resume Next
  Set OutlookApp = GetObject(, "Outlook.Application") ' if Outlook is already running
  If Err.Number <> 0 Then ' Launch a new instance of Outlook
    Err.Clear
    On Error GoTo eh
    Set OutlookApp = CreateObject("Outlook.Application")
  Else
    On Error GoTo eh
  End If

  If ExecutingTroughApplicationRun Then _
    If GetBoolean(CStr(currentRowArray(1, ColAArg1))) Then _
      OutlookApp.Visible = True ' Make the application visible to the user

done:
  StartOutlook = True
  Exit Function
eh:
  StartOutlook = False
  RaiseError MODULE_NAME & ".StartOutlook", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


'https://stackoverflow.com/questions/35609112/how-to-send-a-word-document-as-body-of-an-email-with-vba
Public Function SendEmail(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=To as email address
' Arg2=Subject
' Arg3=Body
' Arg4...9=Attachment files as complete path
  On Error GoTo eh
  StartOutlook
  
  Set oAccount = OutlookApp.Session.Accounts.Item(1)
  Set oMail = OutlookApp.CreateItem(olMailItem)
  
  With oMail
    .To = currentRowArray(1, ColAArg1 + 0)
    .Subject = currentRowArray(1, ColAArg1 + 1)
    .Body = currentRowArray(1, ColAArg1 + 2)

    For tmpL = 3 To 9
      If Len(currentRowArray(1, ColAArg1 + tmpL)) > 0 Then _
        .Attachments.Add currentRowArray(1, ColAArg1 + tmpL)
    Next
    
    '.Display
    'SendUsingAccount is new in Office 2007
    'Change Item(1)to the account number that you want to use
    .SendUsingAccount = oAccount
    
    .Send
  End With
  
done:
  SendEmail = True
  Exit Function
eh:
  SendEmail = False
  RaiseError MODULE_NAME & ".SendEmail", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function



Public Function SendEmailFormated(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=To as email address
' Arg2=Subject
' Arg3=Body content will be taken from the Word document using complete path
' Arg4...9=Attachment files as complete path
  On Error GoTo eh
  Dim editor As Object ' WordEditor

  StartOutlook

  StartWord
  
  On Error Resume Next
  Set WordDoc = WordApp.Documents(1)
  If Err.Number <> 0 Then
    Err.Clear
    On Error GoTo 0
    Set WordDoc = WordApp.Documents.Open(currentRowArray(1, ColAArg1 + 2))
  Else
    On Error GoTo 0
  End If
  
  WordDoc.Content.Copy
  Sleep 500

  Set oAccount = OutlookApp.Session.Accounts.Item(1)
  Set oMail = OutlookApp.CreateItem(olMailItem)
  
  'OutlookApp.Visible = True
  
  With oMail
    .To = currentRowArray(1, ColAArg1 + 0)
    .Subject = currentRowArray(1, ColAArg1 + 1)
    .Display

    .BodyFormat = olFormatRichText
    Set editor = .GetInspector.WordEditor
    editor.Content.Paste

    For tmpL = 3 To 9
      If Len(currentRowArray(1, ColAArg1 + tmpL)) > 0 Then _
        .Attachments.Add currentRowArray(1, ColAArg1 + tmpL)
    Next

    '.Display
    'SendUsingAccount is new in Office 2007
    'Change Item(1)to the account number that you want to use
    .SendUsingAccount = oAccount

    .Send
  End With
  
done:
  SendEmailFormated = True
  Exit Function
eh:
  SendEmailFormated = False
  RaiseError MODULE_NAME & ".SendEmailFormated", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

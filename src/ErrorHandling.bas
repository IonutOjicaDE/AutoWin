Attribute VB_Name = "ErrorHandling"
Option Explicit

Private Const SendLogToLogFile As Byte = 0
Private Const SendLogToImmediateWindow As Byte = 1
Private Const SendLogToMessageBox As Byte = 2

Private Const SendLogTo As Byte = SendLogToLogFile

Private Const LINE_NO_TEXT As String = "Line no: "
Dim AlreadyUsed As Boolean

' ExcelMacroMastery.com Error handling code
' https://excelmacromastery.com/vba-error-handling/

' Example of using:
'
' 1. Place DisplayError in the topmost sub at the bottom.
'    Replace the first paramter with the name of the sub.
' DisplayError "Module1.Topmost", Err.Source, Err.Description, Erl, customLineNumber
'
' 2. Place RaiseError in all the other subs at the bottom of each.
'    Replace the first paramter with the name of the sub.
' RaiseError "Module1.Level1", Err.Number, Err.Source, Err.Description, Erl, customLineNumber
'
'
' 3. The error handling in each sub should look like this
'
'  Sub subName()
'
'    On Error Goto eh
'
'    The main code of the sub here!!!!!
'
'  done:
'      Exit Sub
'  eh:
'      DisplayError "Module1.Topmost", Err.Source, Err.Description, Erl
'  End Sub
'

' Reraises an error and adds line number and current procedure name
Public Sub RaiseError(ByVal proc As String _
                    , ByVal errorNo As Long _
                    , ByVal src As String, ByVal desc As String, ByVal lineNoErl As Long, Optional ByVal lineNo As Long = 0, Optional ByVal ExecutingTroughApplicationRun As Boolean = False)
  'stopExecutionRequired = True
  Dim sSource As String
  
  ' If called for the first time then add line number
  If AlreadyUsed = False Then
  
    ' Add error line number if present
    If lineNoErl <> 0 Then
      sSource = vbNewLine & LINE_NO_TEXT & lineNoErl & " "
    ElseIf lineNo <> 0 Then
      sSource = vbNewLine & LINE_NO_TEXT & lineNo & " "
    End If
    
    ' Add procedure to source
    sSource = sSource & vbNewLine & proc
    AlreadyUsed = True
    
  Else
    ' If error has already been raised simply add on procedure name
    sSource = src & vbNewLine & proc
  End If
  
  ' Pause the code here when debugging
  ' (To Debug: "Tools->VBA Properties" from the menu.
  ' Add "Debugging=1" to the Conditional Compilation Arguments.)
  #If Debugging = 1 Then
  Debug.Assert False
  #End If
  
  ' Reraise the error so it will be caught in the caller procedure
  ' (Note: If the code stops here, make sure DisplayError has been
  ' placed in the topmost procedure)
  If ExecutingTroughApplicationRun Then
    ' https://stackoverflow.com/a/77416358
    ' This weird hack prevents the error information from getting lost when Application.Run returns:
    On Error GoTo -1
    errorNumber = errorNo
    errorSource = sSource
    errorDescription = desc
  Else
    Err.Raise errorNo, sSource, desc
  End If

End Sub

' Displays the error when it reaches the topmost sub
' Note: You can add a call to logging from this sub
Public Sub DisplayError(ByVal proc As String _
                      , ByVal src As String, ByVal desc As String, ByVal lineNoErl As Long, Optional ByVal lineNo As Long = 0)

  ' Check If the error happens in topmost sub
  If AlreadyUsed = False Then
    ' Reset string to remove "VBAProject" and add line number if it exists
    If lineNo <> 0 Then
      src = vbNewLine & LINE_NO_TEXT & lineNoErl
    ElseIf lineNo <> 0 Then
      src = vbNewLine & LINE_NO_TEXT & lineNo
    Else
      src = ""
    End If
  End If
  
  ' Build the final message
  Dim sMsg As String
  sMsg = vbNewLine & "#############################" & _
         vbNewLine & "The following error occurred:" & _
         vbNewLine & Err.Description & _
         vbNewLine & vbNewLine & "Error Location is:"
  sMsg = sMsg & src & vbNewLine & proc
  
  ' Display the message
  Log sMsg
  
  ' reset the boolean value
  AlreadyUsed = False

End Sub

Private Function Log(ByRef a_stringLogThis As String)
  ' concatenate date and what the user wants logged
  Dim l_stringLogStatement As String
  l_stringLogStatement = format(Now, "YYYY-MM-DD hh:mm:ss") & " " & a_stringLogThis
  Select Case SendLogTo
    Case SendLogToLogFile
      ' append (not write) to disk
      Logger l_stringLogStatement
'      Open Replace(ThisWorkbook.FullName, "xlsm", "log") For Append As #1
'      Print #1, l_stringLogStatement
'      Close #1
    Case SendLogToImmediateWindow
      ' send to TTY
      Debug.Print l_stringLogStatement
    Case SendLogToMessageBox
      MsgBox l_stringLogStatement, Title:="Error", Buttons:=vbCritical
  End Select
End Function
Public Function LogClear()
  Debug.Print ("Erasing the previous logs.")
  Open ThisWorkbook.path & "\Log.txt" For Output As #1
  Print #1, ""
  Close #1
End Function


Private Sub Logger(Text As String)
  On Error GoTo eh
  Dim sFilename As String
  sFilename = Replace(ThisWorkbook.FullName, "xlsm", "log")
  
  ' Archive file at certain size
  If FileExists(sFilename) Then
    If FileLen(sFilename) > 20000 Then
    FileCopy sFilename, Replace(sFilename, ".log", format(Now, "yyyymmdd-hhmmss.log"))
    Kill sFilename
    End If
  End If
  
  ' Open the file to write
  Dim filenumber As Variant
  filenumber = FreeFile
  Open sFilename For Append As #filenumber
  
  Print #filenumber, Text
  
  Close #filenumber
eh:
  #If Debugging = 1 Then
    Debug.Print "Error; " & Err.Number, Err.Source, Err.Description
    Debug.Assert False
    Exit Sub
    Resume
  #End If
End Sub

Private Function FileExists(fPath As String) As Boolean
  Dim sFileExists As String
  Let sFileExists = Dir(fPath)
  FileExists = sFileExists <> ""
End Function


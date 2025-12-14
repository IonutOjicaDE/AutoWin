Attribute VB_Name = "LineNumbersInVBA"
Option Explicit

'https://stackoverflow.com/questions/40731182/excel-vba-how-to-turn-on-line-numbers-in-code-editor

'You can make calls like this :
'
'Sub AddLineNumbers_vbLabelColon()
'    AddLineNumbers wbName:="EvaluateCall.xlsm", vbCompName:="ModLineNumbers_testDest", LabelType:=vbLabelColon, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
'End Sub
'
'Sub AddLineNumbers_vbLabelTab()
'    AddLineNumbers wbName:="EvaluateCall.xlsm", vbCompName:="ModLineNumbers_testDest", LabelType:=vbLabelTab, AddLineNumbersToEmptyLines:=True, AddLineNumbersToEndOfProc:=True, Scope:=vbScopeAllProc
'End Sub
'
'Sub RemoveLineNumbers_vbLabelColon()
'    RemoveLineNumbers wbName:="EvaluateCall.xlsm", vbCompName:="ModLineNumbers_testDest", LabelType:=vbLabelColon
'End Sub
'
'Sub RemoveLineNumbers_vbLabelTab()
'    RemoveLineNumbers wbName:="EvaluateCall.xlsm", vbCompName:="ModLineNumbers_testDest", LabelType:=vbLabelTab
'End Sub
'
'And as a reminder, here as some compile rules about line numbers:
'
'not allowed before a Sub/Function declaration statement
'not allowed outside of a proc
'not allowed on a line following a line continuation character "_" (underscore)
'not allowed to have more than one label/line number per code line ~~> Existing labels other than line numbers must be tested otherwise a compile error will occur trying to force a line number.
'not allowed to use characters that already have a special VBA meaning ~~> Allowed characters are [a-Z], [0-9], é, è, ô, ù, €, £, § and even ":" alone !
'compiler will trim any space before a label ~~> So if there is a label, the first char of the line is the first char of the label, it cannot be a space.
'appending a line number with a colon will result in having a space inserted between the ":" and the fist next char if there is none
'when appending a line number with a tab/space, there must be at least one space between the last digit and the first next char, compiler won't add it as it does for a label with a colon separator
'the .ReplaceLine method will overide the compile rules without displaying any compile error as it does in design mode when selecting a new line or when manually relaunching compilation
'the compiler is 'quicker than the VBA environment/system': for example, just after a line number with colon and without any space has been inserted with .ReplaceLine, if the .Lines property is called to get the new string, the space (between the colon character and the first character of the string) is already appended in that string !
'it is not possible to enter debug mode after a .ReplaceLine has been called (from within or outside the module it is editting), not till the code is running, and execution reset.


Public Enum vbLineNumbers_LabelTypes
  vbLabelColon    ' 0
  vbLabelTab      ' 1
End Enum

Public Enum vbLineNumbers_ScopeToAddLineNumbersTo
  vbScopeAllProc  ' 1
  vbScopeThisProc ' 2
End Enum

Public Sub AddLineNumbers(ByVal wbName As String, _
                          ByVal vbCompName As String, _
                          ByVal LabelType As vbLineNumbers_LabelTypes, _
                          ByVal AddLineNumbersToEmptyLines As Boolean, _
                          ByVal AddLineNumbersToEndOfProc As Boolean, _
                          ByVal Scope As vbLineNumbers_ScopeToAddLineNumbersTo, _
                 Optional ByVal thisProcName As String)

' USAGE RULES
' DO NOT MIX LABEL TYPES FOR LINE NUMBERS! IF ADDING LINE NUMBERS AS COLON TYPE, ANY LINE NUMBERS AS VBTAB TYPE MUST BE REMOVE BEFORE, AND RECIPROCALLY ADDING LINE NUMBERS AS VBTAB TYPE

  Dim i As Long
  Dim j As Long
  Dim procName As String
  Dim startOfProcedure As Long
  Dim lengthOfProcedure As Long
  Dim endOfProcedure As Long
  Dim strLine As String
  Dim bodyOfProcedure As Long, countOfProcedure As Long, prelinesOfProcedure As Long
  Dim InProcBodyLines As Boolean, PreviousIndentAdded As Long
  Dim temp_strLine As String, new_strLine As String

  With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule
    .CodePane.Window.Visible = False

    If Scope = vbScopeAllProc Then

      For i = 1 To .CountOfLines

        strLine = .Lines(i, 1)
        procName = .ProcOfLine(i, vbext_pk_Proc) ' Type d'argument ByRef incompatible ~~> Requires VBIDE library as a Reference for the VBA Project

        If procName <> vbNullString Then
          startOfProcedure = .ProcStartLine(procName, vbext_pk_Proc)
          bodyOfProcedure = .ProcBodyLine(procName, vbext_pk_Proc)
          countOfProcedure = .ProcCountLines(procName, vbext_pk_Proc)

          prelinesOfProcedure = bodyOfProcedure - startOfProcedure
          'postlineOfProcedure = ??? not directly available since endOfProcedure is itself not directly available.

          lengthOfProcedure = countOfProcedure - prelinesOfProcedure ' includes postlinesOfProcedure !
          'endOfProcedure = ??? not directly available, each line of the proc must be tested until the End statement is reached. See below.

          If endOfProcedure <> 0 And startOfProcedure < endOfProcedure And i > endOfProcedure Then GoTo NextLine

          If i = bodyOfProcedure Then InProcBodyLines = True

          If bodyOfProcedure < i And i < startOfProcedure + countOfProcedure Then
            If Not (.Lines(i - 1, 1) Like "* _") Then

              InProcBodyLines = False

              PreviousIndentAdded = 0

              If Trim(strLine) = "" And Not AddLineNumbersToEmptyLines Then GoTo NextLine

              If IsProcEndLine(wbName, vbCompName, i) Then
                endOfProcedure = i
                If AddLineNumbersToEndOfProc Then
                  Call IndentProcBodyLinesAsProcEndLine(wbName, vbCompName, LabelType, endOfProcedure)
                Else
                  GoTo NextLine
                End If
              End If

              If LabelType = vbLabelColon Then
                If HasLabel(strLine, vbLabelColon) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelColon)
                If Not HasLabel(strLine, vbLabelColon) Then
                  temp_strLine = strLine
                  .ReplaceLine i, CStr(i) & ":" & strLine
                  new_strLine = .Lines(i, 1)
                  If Len(new_strLine) = Len(CStr(i) & ":" & temp_strLine) Then
                    PreviousIndentAdded = Len(CStr(i) & ":")
                  Else
                    PreviousIndentAdded = Len(CStr(i) & ": ")
                  End If
                End If
              ElseIf LabelType = vbLabelTab Then
                If Not HasLabel(strLine, vbLabelTab) Then strLine = RemoveOneLineNumber(.Lines(i, 1), vbLabelTab)
                If Not HasLabel(strLine, vbLabelColon) Then
                  temp_strLine = strLine
                  .ReplaceLine i, CStr(i) & vbTab & strLine
                  PreviousIndentAdded = Len(strLine) - Len(temp_strLine)
                End If
              End If

            Else
              If Not InProcBodyLines Then
                If LabelType = vbLabelColon Then
                  .ReplaceLine i, Space(PreviousIndentAdded) & strLine
                ElseIf LabelType = vbLabelTab Then
                  .ReplaceLine i, Space(4) & strLine
                End If
              Else
              End If
            End If

          End If

        End If

NextLine:
      Next i

    ElseIf AddLineNumbersToEmptyLines And Scope = vbScopeThisProc Then
    
    End If

    .CodePane.Window.Visible = True
  End With

End Sub

Private Function IsProcEndLine(ByVal wbName As String, ByVal vbCompName As String, ByVal line As Long) As Boolean

  With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule
    If Trim(.Lines(line, 1)) Like "End Sub*" _
      Or Trim(.Lines(line, 1)) Like "End Function*" _
      Or Trim(.Lines(line, 1)) Like "End Property*" _
      Then IsProcEndLine = True
  End With

End Function

Private Sub IndentProcBodyLinesAsProcEndLine(ByVal wbName As String, ByVal vbCompName As String, ByVal LabelType As vbLineNumbers_LabelTypes, ByVal ProcEndLine As Long)
  Dim procName As String
  Dim startOfProcedure As Long
  Dim endOfProcedure As Long
  Dim bodyOfProcedure As Long, j As Long, strEnd As String, strLine As String

  With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule

    procName = .ProcOfLine(ProcEndLine, vbext_pk_Proc)
    bodyOfProcedure = .ProcBodyLine(procName, vbext_pk_Proc)
    endOfProcedure = ProcEndLine
    strEnd = .Lines(endOfProcedure, 1)

    j = bodyOfProcedure
    Do Until Not .Lines(j - 1, 1) Like "* _" And j <> bodyOfProcedure

      strLine = .Lines(j, 1)

      If LabelType = vbLabelColon Then
        If Mid(strEnd, Len(CStr(endOfProcedure)) + 1 + 1 + 1, 1) = " " Then
          .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 1) & strLine
        Else
          .ReplaceLine j, Space(Len(CStr(endOfProcedure)) + 2) & strLine
        End If
      ElseIf LabelType = vbLabelTab Then
        If endOfProcedure < 1000 Then
          .ReplaceLine j, Space(4) & strLine
        Else
          Debug.Print "This tool is limited to 999 lines of code to work properly."
        End If
      End If

      j = j + 1
    Loop

  End With
End Sub

Public Sub RemoveLineNumbers(ByVal wbName As String, ByVal vbCompName As String, ByVal LabelType As vbLineNumbers_LabelTypes)
  Dim i As Long
  Dim procName As String, InProcBodyLines As Boolean, LenghtBefore As Long, RemovedChars_previous_i As Long, LenghtAfter As Long
  Dim LengthBefore_previous_i As Long, LenghtAfter_previous_i As Long, LenOfRemovedLeadingCharacters As Long
  Dim bodyOfProcedure As String, j As Long, strLineBodyOfProc As String, LastLineBodyOfProc As Long, strLastLineBodyOfProc As String
  Dim strLineEndOfProc As String, k As Long
  With Workbooks(wbName).VBProject.VBComponents(vbCompName).CodeModule

    For i = 1 To .CountOfLines

      procName = .ProcOfLine(i, vbext_pk_Proc)

      If procName <> vbNullString Then

        If i = .ProcBodyLine(procName, vbext_pk_Proc) Then InProcBodyLines = True

        LenghtBefore = Len(.Lines(i, 1))
        If Not .Lines(i - 1, 1) Like "* _" Then
          InProcBodyLines = False
          .ReplaceLine i, RemoveOneLineNumber(.Lines(i, 1), LabelType)
        Else
          If InProcBodyLines Then ' original era IsInProcBodyLines
            ' do nothing
          Else
            .ReplaceLine i, Mid(.Lines(i, 1), RemovedChars_previous_i + 1)
          End If
        End If
        LenghtAfter = Len(.Lines(i, 1))

        LengthBefore_previous_i = LenghtBefore
        LenghtAfter_previous_i = LenghtAfter
        RemovedChars_previous_i = LengthBefore_previous_i - LenghtAfter_previous_i

        If Trim(.Lines(i, 1)) Like "End Sub*" Or Trim(.Lines(i, 1)) Like "End Function" Or Trim(.Lines(i, 1)) Like "End Property" Then

          LenOfRemovedLeadingCharacters = LenghtBefore - LenghtAfter

          procName = .ProcOfLine(i, vbext_pk_Proc)
          bodyOfProcedure = .ProcBodyLine(procName, vbext_pk_Proc)

          j = bodyOfProcedure
          strLineBodyOfProc = .Lines(bodyOfProcedure, 1)
          Do Until Not strLineBodyOfProc Like "* _"
            j = j + 1
            strLineBodyOfProc = .Lines(j, 1)
          Loop
          LastLineBodyOfProc = j
          strLastLineBodyOfProc = strLineBodyOfProc

          strLineEndOfProc = .Lines(i, 1)
          For k = bodyOfProcedure To j
            .ReplaceLine k, Mid(.Lines(k, 1), 1 + LenOfRemovedLeadingCharacters)
          Next k

          i = i + (j - bodyOfProcedure)
          GoTo NextLine

        End If
      Else
      ' GoTo NextLine
      End If
NextLine:
    Next i
  End With
End Sub

Private Function RemoveOneLineNumber(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes)
    RemoveOneLineNumber = aString
    If LabelType = vbLabelColon Then
        If aString Like "#:*" Or aString Like "##:*" Or aString Like "###:*" Then
            RemoveOneLineNumber = Mid(aString, 1 + InStr(1, aString, ":", vbTextCompare))
            If Left(RemoveOneLineNumber, 2) Like " [! ]*" Then RemoveOneLineNumber = Mid(RemoveOneLineNumber, 2)
        End If
    ElseIf LabelType = vbLabelTab Then
        If aString Like "#   *" Or aString Like "##  *" Or aString Like "### *" Then RemoveOneLineNumber = Mid(aString, 5)
        If aString Like "#" Or aString Like "##" Or aString Like "###" Then RemoveOneLineNumber = ""
    End If
End Function

Private Function HasLabel(ByVal aString As String, ByVal LabelType As vbLineNumbers_LabelTypes) As Boolean
    If LabelType = vbLabelColon Then HasLabel = InStr(1, aString & ":", ":") < InStr(1, aString & " ", " ")
    If LabelType = vbLabelTab Then
        HasLabel = Mid(aString, 1, 4) Like "#   " Or Mid(aString, 1, 4) Like "##  " Or Mid(aString, 1, 4) Like "### "
    End If
End Function

Private Function RemoveLeadingSpaces(ByVal aString As String) As String
    Do Until Left(aString, 1) <> " "
        aString = Mid(aString, 2)
    Loop
    RemoveLeadingSpaces = aString
End Function

Private Function WhatIsLineIndent(ByVal aString As String) As String
  Dim i As Long
    i = 1
    Do Until Mid(aString, i, 1) <> " "
        i = i + 1
    Loop
    WhatIsLineIndent = i
End Function

Private Function HowManyLeadingSpaces(ByVal aString As String) As String
  HowManyLeadingSpaces = WhatIsLineIndent(aString) - 1
End Function

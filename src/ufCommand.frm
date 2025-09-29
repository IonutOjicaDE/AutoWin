VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCommand 
   Caption         =   "Command"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7620
   OleObjectBlob   =   "ufCommand.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private commandList      As Object     ' Dictionary to store filtered commands
Private foundCommands    As Object
Private selectedCategory As String     ' Stores the selected category name
Private selectedCommand  As String     ' Stores the selected command name
Private selectedArgs     As Collection ' Stores references to dynamically created argument TextBoxes
Private cmdInfo          As Variant
Private lbl              As MSForms.Label
Private cmb              As MSForms.ComboBox
Private tmpV             As Variant
Private tmpL             As Long
Private tmpS             As String

Private Const allCommandsText   As String = "All Commands"
Private Const foundCommandsText As String = "Found Commands"

Private mSubIdent()   As String        ' Cached arrays (1D, zero-based) – may be Empty when none
Private mLabelIdent() As String
Private mForIdent()   As String
Private mLoopIdent()  As String
Private mIdentReady   As Boolean       ' Flag: is mSubNames filled



Private Sub UserForm_Activate()
  If Not commandList Is Nothing Then Set commandList = Nothing
  If Not selectedArgs Is Nothing Then Set selectedArgs = Nothing
  Set commandList = CreateObject("Scripting.Dictionary")
  Set foundCommands = CreateObject("Scripting.Dictionary")
  Set selectedArgs = New Collection

  Call InvalidateIdents

  Call PopulateCategories

  If ActiveSheet Is shAuto Then
    currentRow = ActiveCell.Row
    
    If currentRow >= startRow Then
      Call SaveCurrentRowValues

      selectedCommand = CleanString(CStr(currentRowArray(1, ColACommand)))


      If Len(selectedCommand) = 0 Then ' No command written: Select first available command
        lstCategories.Value = allCommandsText
        lstCommands.SetFocus
        lstCommands.ListIndex = 0


      ElseIf commandMap.Exists(selectedCommand) Then ' Exact match found
        cmdInfo = commandMap(selectedCommand)
        selectedCategory = cmdInfo(cmdCategory)

        lstCategories.Value = selectedCategory ' Select the corresponding category
        lstCommands.Value = cmdInfo(cmdDisplayName)

        Call LoadArgumentsFromSheet


      Else                             ' Partial match: Create "Found" category

        foundCommands.RemoveAll

        For Each tmpV In commandMap.Keys
          If InStr(1, tmpV, selectedCommand, vbTextCompare) > 0 Then
            Call foundCommands.Add(tmpV, commandMap(tmpV)(cmdDisplayName))
          End If
        Next
        
        If foundCommands.count > 0 Then ' Matches found: create "Found" category
          lstCategories.Value = foundCommandsText
          lstCommands.ListIndex = 0

        Else                           ' foundCommands.Count <= 0 : No matches => Default to "All Commands"
          lstCategories.Value = allCommandsText
          lstCommands.ListIndex = 0
          Call LoadArgumentsFromSheet

        End If                         ' foundCommands.Count > 0

      End If                           ' Len(selectedCommand) = 0

      Call LoadArgumentsFromSheet

    Else                               ' currentRow < startRow
      Hide
    End If

  Else                                 ' Not ActiveSheet Is shAuto
    Hide
  End If
  Call UpdateDebug
End Sub

Private Sub UpdateDebug()
  dbgSelectedCategory.Caption = selectedCategory
  dbgSelectedCommand.Caption = selectedCommand
End Sub

Private Sub PopulateCategories()
  Dim key As Variant, categories As Object
  Set categories = CreateObject("Scripting.Dictionary")

  For Each key In commandMap.Keys      ' Collect unique categories
    If Not categories.Exists(commandMap(key)(cmdCategory)) Then
      categories.Add commandMap(key)(cmdCategory), commandMap(key)(cmdCategory)
    End If
  Next key

  lstCategories.Clear                  ' Add categories to ListBox
  For Each key In categories.Keys
    Call lstCategories.AddItem(key)
  Next key

  Call lstCategories.AddItem(allCommandsText)
  Call lstCategories.AddItem(foundCommandsText)
End Sub

Private Sub PopulateCommands()
  lstCommands.Clear
  commandList.RemoveAll
  If selectedCategory = foundCommandsText Then
    For Each tmpV In foundCommands.Keys
      Call lstCommands.AddItem(foundCommands(tmpV))
      Call commandList.Add(foundCommands(tmpV), tmpV)
    Next
  Else
    For Each tmpV In commandMap.Keys
      If commandMap(tmpV)(cmdCategory) = selectedCategory Or selectedCategory = allCommandsText Then
        Call lstCommands.AddItem(commandMap(tmpV)(cmdDisplayName))   ' DisplayName
        Call commandList.Add(commandMap(tmpV)(cmdDisplayName), tmpV) ' Store CommandID
      End If
    Next
  End If
  If Len(selectedCommand) > 0 And Not IsEmpty(cmdInfo) Then
    If commandList.Exists(cmdInfo(cmdDisplayName)) Then lstCommands.Value = cmdInfo(cmdDisplayName)
  End If
  Call UpdateDebug
End Sub

Private Sub LoadArgumentsFromSheet()
  For tmpL = 10 To 1 Step -1
    tmpS = currentRowRange(1, ColAArg1 + tmpL - 1).Formula
    Set cmb = Me.Controls("cmbArg" & tmpL)
    cmb.Value = tmpS
  Next
End Sub



Private Sub lstCategories_Change()
  If IsNull(lstCategories.Value) Then Exit Sub
  selectedCategory = lstCategories.Value
  Call PopulateCommands
End Sub
'Public Const cmdFunctionName    As Long = 0& ' Name of VBA function
'Public Const cmdDisplayName     As Long = 1& ' Pretty name with spaces
'Public Const cmdCategory        As Long = 2& ' Category (General, Mouse, Window, etc.)
'Public Const cmdDescription     As Long = 3& ' Command description
'Public Const cmdArgName1        As Long = 4& ' Argument names
'Public Const cmdArgDescription1 As Long = 5& ' Argument descriptions

Private Sub lstCommands_Change()
  If lstCommands.ListIndex = -1 Then Exit Sub

                                      ' Retrieve command info
  selectedCommand = commandList(lstCommands.List(lstCommands.ListIndex))
  
  cmdInfo = commandMap(selectedCommand)

                                      ' Update description
  lblDescription.Caption = "{ " & cmdInfo(cmdCategory) & " }" & vbCrLf & "[ " & cmdInfo(cmdDisplayName) & " ]" & vbCrLf & cmdInfo(cmdDescription) ' Command Description

                                      ' Generate input fields for arguments
  For tmpL = cmdArgName1 To cmdArgName1 + 19 Step 2
    Set lbl = Me.Controls("lblArg" & ((tmpL - cmdArgName1) / 2 + 1))
    Set cmb = Me.Controls("cmbArg" & ((tmpL - cmdArgName1) / 2 + 1))

    If tmpL < UBound(cmdInfo) Then
      lbl.Enabled = True
      lbl.Caption = cmdInfo(tmpL)      ' Name of argument
      cmb.Tag = cmdInfo(tmpL + 1)      ' Description of argument

      tmpV = GetDoubleBraceItems(cmb.Tag)
      If IsArray(tmpV) Then
                                       ' Label contains {{...}} ? load options and show arrow drop button
        Call ConfigureCmbForList
      Else
        Call ConfigureCmbForFree
      End If
    Else
      lbl.Enabled = False
      lbl.Caption = "Argument" & ((tmpL - cmdArgName1) / 2 + 1)
      Call ConfigureCmbForFree
    End If
  Next

  If cmdArgDescription1 <= UBound(cmdInfo) Then lblArgDescription.Caption = cmdInfo(cmdArgDescription1) Else lblArgDescription.Caption = vbNullString
  Call UpdateDebug
End Sub



' Returns an array of items found inside the first {{...}} pair of s, split by "/"
' - Trims spaces around "/" (keeps inner-word spaces)
' - Returns Empty if no valid {{...}} content
Private Function GetDoubleBraceItems(ByVal s As String) As Variant
  Dim p1 As Long, p2 As Long, inner As String, raw As Variant
  Dim i As Long, buf() As String, n As Long, t As String
  
  p1 = InStr(1, s, "{{", vbBinaryCompare)
  If p1 = 0 Then Exit Function
  p2 = InStr(p1 + 2, s, "}}", vbBinaryCompare)
  If p2 = 0 Or p2 <= p1 + 2 Then Exit Function
  
  inner = Mid$(s, p1 + 2, p2 - p1 - 2)
  raw = Split(inner, "/")
  ReDim buf(0 To UBound(raw))          ' max needed
  
  For i = LBound(raw) To UBound(raw)
    t = Trim$(CStr(raw(i)))            ' remove spaces around "/"
    If LenB(t) > 0 Then
      buf(n) = t
      n = n + 1
    End If
  Next i
  
  If n = 0 Then Exit Function
  ReDim Preserve buf(0 To n - 1)
  GetDoubleBraceItems = buf
End Function


Private Sub ConfigureCmbForList()
  Dim prev As String: prev = cmb.Text          ' save current value
  Dim i As Long
  With cmb
    '.Style = fmStyleDropDownCombo             ' allow typing custom values too
    '.MatchRequired = False                    ' not restricted to list
    '.MatchEntry = fmMatchEntryComplete        ' autocomplete
    .DropButtonStyle = fmDropButtonStyleArrow  ' arrow button (1)
    .Clear                                     ' clear DropDown menu
    If IsArray(tmpV) Then
      For i = LBound(tmpV) To UBound(tmpV)
        If LenB(tmpV(i)) > 0 Then .AddItem CStr(tmpV(i))
      Next
    End If
    .Text = prev                               ' restore previous value
  End With
End Sub

Private Sub ConfigureCmbForFree()
  Dim prev As String: prev = cmb.Text          ' save current value
  With cmb
    '.Style = fmStyleDropDownCombo             ' allow typing custom values too
    '.MatchRequired = False                    ' not restricted to list
    '.MatchEntry = fmMatchEntryComplete        ' autocomplete
    .DropButtonStyle = IIf(ArgDescHasListTokens(.Tag), fmDropButtonStyleArrow, fmDropButtonStylePlain)
    .Clear                                     ' clear DropDown menu
    .Text = prev                               ' restore previous value
  End With
End Sub

' --- Detect if arg description requests dynamic lists via tokens ---
Private Function ArgDescHasListTokens(ByVal descText As String) As Boolean
  If InStr(descText, loopListSub) > 0 Then ArgDescHasListTokens = True: Exit Function
  If InStr(descText, loopListLabel) > 0 Then ArgDescHasListTokens = True: Exit Function
  If InStr(descText, loopListFor) > 0 Then ArgDescHasListTokens = True: Exit Function
  If InStr(descText, loopListLoop) > 0 Then ArgDescHasListTokens = True: Exit Function
End Function




' Handles formula evaluation
Private Sub cmbArg_Change(ByRef cmbArg As MSForms.ComboBox)
  If Left(cmbArg.Text, 1) = "=" Then
    tmpV = Application.Evaluate(cmbArg.Text)
    If (VBA.VarType(tmpV) = vbError) Then
      lblArgDescription.Caption = cmbArg.Tag
    Else
      lblArgDescription.Caption = cmbArg.Tag & vbCrLf & vbCrLf & "Value=" & tmpV
    End If
  Else
    lblArgDescription.Caption = cmbArg.Tag
  End If
End Sub

Private Sub cmbArg1_Enter():   Call cmbArg_Change(cmbArg1):  End Sub
Private Sub cmbArg2_Enter():   Call cmbArg_Change(cmbArg2):  End Sub
Private Sub cmbArg3_Enter():   Call cmbArg_Change(cmbArg3):  End Sub
Private Sub cmbArg4_Enter():   Call cmbArg_Change(cmbArg4):  End Sub
Private Sub cmbArg5_Enter():   Call cmbArg_Change(cmbArg5):  End Sub
Private Sub cmbArg6_Enter():   Call cmbArg_Change(cmbArg6):  End Sub
Private Sub cmbArg7_Enter():   Call cmbArg_Change(cmbArg7):  End Sub
Private Sub cmbArg8_Enter():   Call cmbArg_Change(cmbArg8):  End Sub
Private Sub cmbArg9_Enter():   Call cmbArg_Change(cmbArg9):  End Sub
Private Sub cmbArg10_Enter():  Call cmbArg_Change(cmbArg10): End Sub

Private Sub cmbArg1_Change():  Call cmbArg_Change(cmbArg1):  End Sub
Private Sub cmbArg2_Change():  Call cmbArg_Change(cmbArg2):  End Sub
Private Sub cmbArg3_Change():  Call cmbArg_Change(cmbArg3):  End Sub
Private Sub cmbArg4_Change():  Call cmbArg_Change(cmbArg4):  End Sub
Private Sub cmbArg5_Change():  Call cmbArg_Change(cmbArg5):  End Sub
Private Sub cmbArg6_Change():  Call cmbArg_Change(cmbArg6):  End Sub
Private Sub cmbArg7_Change():  Call cmbArg_Change(cmbArg7):  End Sub
Private Sub cmbArg8_Change():  Call cmbArg_Change(cmbArg8):  End Sub
Private Sub cmbArg9_Change():  Call cmbArg_Change(cmbArg9):  End Sub
Private Sub cmbArg10_Change(): Call cmbArg_Change(cmbArg10): End Sub


Private Sub cmbArg1_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg1):  End Sub
Private Sub cmbArg2_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg2):  End Sub
Private Sub cmbArg3_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg3):  End Sub
Private Sub cmbArg4_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg4):  End Sub
Private Sub cmbArg5_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg5):  End Sub
Private Sub cmbArg6_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg6):  End Sub
Private Sub cmbArg7_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg7):  End Sub
Private Sub cmbArg8_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg8):  End Sub
Private Sub cmbArg9_DropButtonClick():  Call cmbArg_DropButtonClick(cmbArg9):  End Sub
Private Sub cmbArg10_DropButtonClick(): Call cmbArg_DropButtonClick(cmbArg10): End Sub


Private Sub cmbArg_DropButtonClick(ByRef cmbArg As MSForms.ComboBox)
  If cmb.DropButtonStyle = fmDropButtonStyleArrow And cmb.ListCount = 0 Then
    ' has {list_*} tokens
    If Not mIdentReady Then Call BuildAllIdents
    Dim d As String: d = cmbArg.Tag

    If InStr(d, loopListSub) > 0 Then Call AddIdents(cmbArg, mSubIdent)
    If InStr(d, loopListLabel) > 0 Then Call AddIdents(cmbArg, mLabelIdent)
    If InStr(d, loopListFor) > 0 Then Call AddIdents(cmbArg, mForIdent)
    If InStr(d, loopListLoop) > 0 Then Call AddIdents(cmbArg, mLoopIdent)
  End If
End Sub

Private Sub AddIdents(ByRef cmbArg As MSForms.ComboBox, ByRef arr() As String)
  If (Not Not arr) = 0 Then            ' ArrOrEmpty = Empty
  Else                                 ' ArrOrEmpty = arr
    For tmpL = LBound(arr) To UBound(arr)
      If LenB(arr(tmpL)) > 0 Then cmbArg.AddItem arr(tmpL)
    Next
  End If
End Sub

' Invalidate cache explicitly when sheet changes
Private Sub InvalidateIdents(): mIdentReady = False: Erase mSubIdent: Erase mLabelIdent: Erase mForIdent: Erase mLoopIdent: End Sub
'
' Build all four lists in ONE pass (fast)
Private Sub BuildAllIdents()
  Dim lastRow As Long
  lastRow = maxLong(GetLastLineOnColumn(shAuto, ColACommand), GetLastLineOnColumn(shAuto, ColAArg1))
  If lastRow < 1 Then: Call InvalidateIdents: mIdentReady = True: Exit Sub


  Dim cmdTypeMap As Object: Set cmdTypeMap = CreateObject("Scripting.Dictionary")
  Dim cmdType As Long
  Dim desc As String, dispName As String
  For Each tmpV In commandMap.Keys
    desc = CStr(commandMap(tmpV)(cmdDescription))
    'dispName = CStr(commandMap(tmpV)(cmdDisplayName))

    cmdType = loopListFlagNone
    If InStr(desc, loopListSub) > 0 Then cmdType = cmdType Or loopListFlagSub
    If InStr(desc, loopListLabel) > 0 Then cmdType = cmdType Or loopListFlagLabel
    If InStr(desc, loopListFor) > 0 Then cmdType = cmdType Or loopListFlagFor
    If InStr(desc, loopListLoop) > 0 Then cmdType = cmdType Or loopListFlagLoop

    'If cmdType <> loopListFlagNone Then cmdTypeMap(dispName) = cmdType
    If cmdType <> loopListFlagNone Then cmdTypeMap(tmpV) = cmdType
  Next


  Dim vCmd As Variant, vArg1 As Variant
  vCmd = shAuto.Range(shAuto.Cells(1, ColACommand), shAuto.Cells(lastRow, ColACommand)).Value2
  vArg1 = shAuto.Range(shAuto.Cells(1, ColAArg1), shAuto.Cells(lastRow, ColAArg1)).Value2

  ReDim mSubIdent(0 To lastRow - 1)    ' preallocate to lastRow, trim later
  ReDim mLabelIdent(0 To lastRow - 1)
  ReDim mForIdent(0 To lastRow - 1)
  ReDim mLoopIdent(0 To lastRow - 1)

  Dim cSub As Long, cLabel As Long, cFor As Long, cLoop As Long
  Dim r As Long
  Dim cmdRaw As Variant, identRaw As Variant
  Dim cmdNorm As String, identNorm As String

  For r = 1 To lastRow
    cmdRaw = vCmd(r, 1)
    If VarType(cmdRaw) = vbError Then GoTo NextRow
    If LenB(cmdRaw) = 0 Then GoTo NextRow
    cmdNorm = LCase$(Replace$(CStr(cmdRaw), " ", vbNullString)) ' CleanString(cmdRaw)

    identRaw = vArg1(r, 1)
    If VarType(identRaw) = vbError Then GoTo NextRow
    If LenB(identRaw) = 0 Then GoTo NextRow
    identNorm = CStr(identRaw)

    If cmdTypeMap.Exists(cmdNorm) Then
      cmdType = cmdTypeMap(cmdNorm)
      If (cmdType And loopListFlagSub) <> 0 Then mSubIdent(cSub) = identNorm:       cSub = cSub + 1
      If (cmdType And loopListFlagLabel) <> 0 Then mLabelIdent(cLabel) = identNorm: cLabel = cLabel + 1
      If (cmdType And loopListFlagFor) <> 0 Then mForIdent(cFor) = identNorm:       cFor = cFor + 1
      If (cmdType And loopListFlagLoop) <> 0 Then mLoopIdent(cLoop) = identNorm:    cLoop = cLoop + 1
    End If


'    Select Case cmdNorm
'      Case loopTypeSub
'        mSubIdent(cSub) = identNorm: cSub = cSub + 1
'
'      Case loopTypeLabel
'        mLabelIdent(cLabel) = identNorm: cLabel = cLabel + 1
'
'      Case "for", "foreach"
'        mForIdent(cFor) = identNorm: cFor = cFor + 1
'
'      Case "do", "dowhile", "dountil"
'        mLoopIdent(cLoop) = identNorm: cLoop = cLoop + 1
'
'      'Case Else: ignore
'    End Select

NextRow:
  Next

  ' Trim arrays to actual counts or make them Empty if none
  If cSub <= 0 Then Erase mSubIdent Else ReDim Preserve mSubIdent(0 To cSub - 1)
  If cFor <= 0 Then Erase mForIdent Else ReDim Preserve mForIdent(0 To cFor - 1)
  If cLoop <= 0 Then Erase mLoopIdent Else ReDim Preserve mLoopIdent(0 To cLoop - 1)
  If cLabel <= 0 Then Erase mLabelIdent Else ReDim Preserve mLabelIdent(0 To cLabel - 1)

  mIdentReady = True
End Sub





Private Sub btnOK_Click()
  If Len(selectedCommand) > 0 And Not IsEmpty(cmdInfo) Then
    currentRowRange(1, ColACommand).Value = cmdInfo(cmdDisplayName)
    For tmpL = 1 To 10
      Set cmb = Me.Controls("cmbArg" & tmpL)

      If Left(cmb.Text, 1) = "=" Then
        tmpV = Application.Evaluate(cmb.Text)
        If (VBA.VarType(tmpV) = vbError) Then
          currentRowRange(1, ColAArg1 + tmpL - 1).Formula = "'" & cmb.Text
        Else
          currentRowRange(1, ColAArg1 + tmpL - 1).Formula = cmb.Text
        End If
      Else
        currentRowRange(1, ColAArg1 + tmpL - 1).Formula = cmb.Text
      End If
    Next
  End If
  Hide
End Sub

Private Sub btnCancel_Click(): Hide: End Sub




Private Sub QuickSortStrings(ByRef a As Variant, ByVal lo As Long, ByVal hi As Long)
  Dim i As Long, j As Long, p As String, t As Variant
  i = lo: j = hi: p = a((lo + hi) \ 2)
  Do While i <= j
    Do While StrComp(a(i), p, vbTextCompare) < 0: i = i + 1: Loop
    Do While StrComp(a(j), p, vbTextCompare) > 0: j = j - 1: Loop
    If i <= j Then t = a(i): a(i) = a(j): a(j) = t: i = i + 1: j = j - 1
  Loop
  If lo < j Then QuickSortStrings a, lo, j
  If i < hi Then QuickSortStrings a, i, hi
End Sub


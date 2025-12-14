Attribute VB_Name = "AutoWinStateCapture"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinStateCapture"

'############################################################
' Purpose: Save Control Path + Control Type + Value Before/After
'############################################################

' --- API declarations ---
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
#If Win64 Then
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
#End If

' --- Constants ---
Private Enum GW_Constants     ' for GetNextWindow, wFlag
  GW_HWNDFIRST = 0&            ' The retrieved handle identifies the window of the same type that is highest in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDLAST = 1&             ' The retrieved handle identifies the window of the same type that is lowest in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDNEXT = 2&             ' The retrieved handle identifies the window below the specified window in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDPREV = 3&             ' The retrieved handle identifies the window above the specified window in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_OWNER = 4&                ' The retrieved handle identifies the specified window's owner window, if any. For more information, see Owned Windows.
  GW_CHILD = 5&                ' The retrieved handle identifies the child window at the top of the Z order, if the specified window is a parent window; otherwise, the retrieved handle is NULL. The function examines only child windows of the specified window. It does not examine descendant windows.
  GW_ENABLEDPOPUP = 6&         ' The retrieved handle identifies the enabled popup window owned by the specified window (the search uses the first such window found using GW_HWNDNEXT); otherwise, if there are no enabled popup windows, the retrieved handle is that of the specified window.
End Enum

Const BM_GETCHECK = &HF0
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const BST_CHECKED = 1

Const GWL_STYLE = -16
Const BS_CHECKBOX = &H2
Const BS_RADIOBUTTON = &H4

Const Separator = " » "


'############################################################





Public Sub RecordFocusedControlState(ByVal eventType As String)
  Const FirstColumnToWrite As Long = 6
  Dim hwndTarget As LongPtr
  hwndTarget = GetFocusedControl()
  
  If hwndTarget = 0 Then Exit Sub

  Dim ctrlType As String
  ctrlType = DetectControlType(hwndTarget)
  If ctrlType = "Unknown" Then Exit Sub
  
  Dim valueBefore As Variant
  valueBefore = ReadControlValue(hwndTarget, ctrlType)
  
  Dim path As Collection
  Set path = BuildControlPath(hwndTarget)

  currentRowRange(1, ColAComment + FirstColumnToWrite + 0).Value = ctrlType
  currentRowRange(1, ColAComment + FirstColumnToWrite + 1).Value = FormatControlPathLinear(path)
  currentRowRange(1, ColAComment + FirstColumnToWrite + 2).Value = valueBefore
  currentRowRange(1, ColAComment + FirstColumnToWrite + 3).Value = "Captured Control on " & eventType
End Sub



Public Sub RecordControlStateUndeMouse(ByVal eventType As String)
  Const FirstColumnToWrite As Long = 6
  Dim hwndTarget As LongPtr
  hwndTarget = GetControlUndeMouse()
  
  If hwndTarget = 0 Then Exit Sub

  Dim ctrlType As String
  ctrlType = DetectControlType(hwndTarget)
  If ctrlType = "Unknown" Then Exit Sub
  
  Dim valueBefore As Variant
  valueBefore = ReadControlValue(hwndTarget, ctrlType)
  
  Dim path As Collection
  Set path = BuildControlPath(hwndTarget)

  currentRowRange(1, ColAComment + FirstColumnToWrite + 0).Value = ctrlType
  currentRowRange(1, ColAComment + FirstColumnToWrite + 1).Value = FormatControlPathLinear(path)
  currentRowRange(1, ColAComment + FirstColumnToWrite + 2).Value = valueBefore
  currentRowRange(1, ColAComment + FirstColumnToWrite + 3).Value = "Captured Control on " & eventType
End Sub



Private Function FormatControlPathLinear(ByVal path As Collection) As String
  Dim i As Long
  For i = 1 To path.count
    Dim ctrl As Object
    Set ctrl = path(i)
    
    FormatControlPathLinear = FormatControlPathLinear & Separator & ctrl("Class") & Separator & ctrl("Text") & Separator & ctrl("Index")
  Next i
  If Len(FormatControlPathLinear) >= Len(Separator) Then FormatControlPathLinear = Mid$(FormatControlPathLinear, Len(Separator) + 1)
End Function




'############################################################

' --- Helpers ---

' Get window text
Private Function GetWindowTextFromHwnd(ByVal hWnd As LongPtr) As String
  Dim sBuffer As String * 255
  Dim lLength As Long
  lLength = GetWindowText(hWnd, sBuffer, 255)
  If lLength > 0 Then GetWindowTextFromHwnd = Left$(sBuffer, lLength)
End Function

' Get class name
Private Function GetClassNameFromHwnd(ByVal hWnd As LongPtr) As String
  Dim sBuffer As String * 255
  Dim lLength As Long
  lLength = GetClassName(hWnd, sBuffer, 255)
  If lLength > 0 Then GetClassNameFromHwnd = Left$(sBuffer, lLength)
End Function

' Get control index among siblings with the same class
Public Function GetControlIndex(ByVal hWnd As LongPtr) As Long
  Dim parentHwnd As LongPtr
  parentHwnd = GetParent(hWnd)
  
  Dim sibling As LongPtr
  Dim index As Long: index = 1
  
  sibling = GetWindow(parentHwnd, GW_CHILD)
  Do While sibling <> 0
    If sibling = hWnd Then Exit Do
    If GetClassNameFromHwnd(sibling) = GetClassNameFromHwnd(hWnd) Then
      index = index + 1
    End If
    sibling = GetWindow(sibling, GW_HWNDNEXT)
  Loop
  
  GetControlIndex = index
End Function

'############################################################

' --- Detect control type ---
Public Function DetectControlType(ByVal hWnd As LongPtr) As String
  Dim cls As String
  cls = LCase(GetClassNameFromHwnd(hWnd))
  
  Select Case True
    Case InStr(cls, "button") > 0
      Dim style As LongPtr
      style = GetWindowLong(hWnd, GWL_STYLE) ' GWL_STYLE
      
      Select Case style
        Case BS_CHECKBOX: DetectControlType = "Checkbox"
        Case BS_RADIOBUTTON: DetectControlType = "RadioButton"
        Case Else: DetectControlType = "Button"
      End Select
      
    Case InStr(cls, "checkbox") > 0
      DetectControlType = "Checkbox"
        
    Case InStr(cls, "radiobutton") > 0
      DetectControlType = "RadioButton"
        
    Case InStr(cls, "edit") > 0
      DetectControlType = "Textbox"
        
    Case Else
      DetectControlType = "Unknown"
  End Select
End Function

'############################################################

' --- Read control value ---
Public Function ReadControlValue(ByVal hWnd As LongPtr, ByVal ctrlType As String) As Variant
  Select Case ctrlType
    Case "Checkbox"
      Dim state As Long
      state = SendMessage(hWnd, BM_GETCHECK, 0, 0)
      Select Case state
        Case 0: ReadControlValue = "Unchecked"
        Case 1: ReadControlValue = "Checked"
        Case 2: ReadControlValue = "Indeterminate"
        Case Else: ReadControlValue = "Unknown"
      End Select
    
    Case "RadioButton"
      Dim rState As Long
      rState = SendMessage(hWnd, BM_GETCHECK, 0, 0)
      If rState = BST_CHECKED Then
        ReadControlValue = "Checked"
      Else
        ReadControlValue = "Unchecked"
      End If
        
    Case "Textbox"
      Dim txtLen As Long
      txtLen = SendMessage(hWnd, WM_GETTEXTLENGTH, 0, 0)
      If txtLen > 0 Then
        Dim txtBuffer As String * 512
        Call SendMessageStr(hWnd, WM_GETTEXT, txtLen + 1, StrPtr(txtBuffer))
        ReadControlValue = Left$(txtBuffer, txtLen)
      Else
        ReadControlValue = ""
      End If
        
    Case Else
      ReadControlValue = Null
  End Select
End Function


'############################################################

' --- Build control path ---
Public Function BuildControlPath(ByVal hWnd As LongPtr) As Collection
  Dim path As New Collection
  Dim currentHwnd As LongPtr: currentHwnd = hWnd
  
  Do While currentHwnd <> 0
    Dim ctrlInfo As Object
    Set ctrlInfo = CreateObject("Scripting.Dictionary")
    
    ctrlInfo("Class") = GetClassNameFromHwnd(currentHwnd)
    ctrlInfo("Text") = GetWindowTextFromHwnd(currentHwnd)
    ctrlInfo("Index") = GetControlIndex(currentHwnd)
    
    path.Add ctrlInfo
    currentHwnd = GetParent(currentHwnd)
  Loop
  
  ' Reverse path
  Dim reversedPath As New Collection
  Dim i As Long
  For i = path.count To 1 Step -1
    reversedPath.Add path(i)
  Next i
  
  Set BuildControlPath = reversedPath
End Function

'############################################################

' --- Save full state ---
Public Sub SaveFullState(ByVal hWnd As LongPtr, ByVal currentRowRange As Range, ByVal ColAComment As Long)
  Dim ctrlType As String
  ctrlType = DetectControlType(hWnd)
  
  Dim valueBefore As Variant
  valueBefore = ReadControlValue(hWnd, ctrlType)
  
  Dim path As Collection
  Set path = BuildControlPath(hWnd)
  
  ' --- Save Path as formatted JSON ---
  Dim json As String
  json = "{\n"
  json = json & "  ""Type"": """ & ctrlType & """,\n"
  json = json & "  ""Path"": [\n"
  
  Dim i As Long
  For i = 1 To path.count
    Dim ctrl As Object
    Set ctrl = path(i)
    
    json = json & "    {""Class"":""" & ctrl("Class") & """,""Text"":""" & Replace(ctrl("Text"), """", "\""") & """,""Index"":" & ctrl("Index") & "}"
    If i < path.count Then json = json & ","
    json = json & vbNewLine
  Next i
  
  json = json & "  ]\n}"
  
  ' --- Save in worksheet ---
  currentRowRange(1, ColAComment + 1).Value = json
  currentRowRange(1, ColAComment + 2).Value = IIf(IsNull(valueBefore), "", valueBefore)
End Sub


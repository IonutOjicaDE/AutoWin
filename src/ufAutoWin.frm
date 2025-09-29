VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufAutoWin 
   Caption         =   "Choose macro to run"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "ufAutoWin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufAutoWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedLine As Long

Private Sub cbCancel_Click()
  Hide
End Sub
Private Sub cbRun_Click()
  SelectedLine = lbMacro.List(lbMacro.ListIndex, 0) + 1
  Hide
End Sub

Private Sub lbMacro_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Call cbRun_Click
End Sub

Private Sub lbMacro_Change()
  If lbMacro.ListCount = 0 Then Exit Sub
  If StrComp(lbMacro.List(lbMacro.ListIndex, 1), recordText, vbTextCompare) Then
    cbRun.Caption = "Run"
  Else
    cbRun.Caption = "Start recording"
  End If
End Sub


Private Sub UserForm_Activate()
  Dim s As String, c As Range, lastRow As Long
' ######### If ActiveCell is on a start from a macro, then run it without to show window
  If ActiveSheet Is shAuto Then
    Select Case LCase(shAuto.Cells(ActiveCell.Row, ColACommand).Text)
      Case "sub"
        SelectedLine = ActiveCell.Row + 1
        Hide
        Exit Sub
      Case "record"
        SelectedLine = ActiveCell.Row + 1
        Hide
        Exit Sub
    End Select
  End If
' ######### Fill the list box (first clear it)
  lbMacro.Clear
  Set c = shAuto.Columns(ColACommand).Find("sub", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
  If Not c Is Nothing Then
    Do
      lbMacro.AddItem c.Row
      lbMacro.List(lbMacro.ListCount - 1, 1) = c.Offset(0, 1).Text
      lastRow = c.Row
      Set c = shAuto.Columns(ColACommand).FindNext(c)
    Loop Until c.Row <= lastRow
  End If
  Set c = shAuto.Columns(ColACommand).Find("record", LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
  If Not c Is Nothing Then
    Do
      lbMacro.AddItem c.Row
      lbMacro.List(lbMacro.ListCount - 1, 1) = "Record"
      lastRow = c.Row
      Set c = shAuto.Columns(ColACommand).FindNext(c)
    Loop Until c.Row <= lastRow
  End If
' ######### Check how many AutoWin macros were found
  SelectedLine = -1
  Select Case lbMacro.ListCount
    Case 0 ' no AutoWin macro found
      Hide
      If MsgBox("There is no AutoWin macro found." & vbCrLf & vbCrLf & _
        "Should I run beginning with line " & startRow & "?", _
        vbYesNoCancel, "No AutoWin macro found") = vbYes Then
        SelectedLine = startRow
      End If
    Case 1 ' only one AutoWin macro found => run it, without the window to show up
      Hide
      SelectedLine = lbMacro.List(0, 0)
    Case Else '>1 more AutoWin macros found
      lbMacro.ListIndex = 0
      lbMacro.SetFocus
  End Select
End Sub

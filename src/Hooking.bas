Attribute VB_Name = "Hooking"
Option Explicit
'https://www.vbforums.com/showthread.php?417465-RESOLVED-Form-Move-Event

Private PrevProc_ufCommand As LongPtr


'##########+##########+##########+##########+##########
' Form ufCommand
'##########+##########+##########+##########+##########
'You would need to subclass your form and look for the WM_MOVING message.

'Public Sub HookForm_ufCommand()
'  On Error GoTo eh
'  PrevProc_ufCommand = SetWindowLong(ufCommand.hWnd, GWL_WNDPROC, AddressOf WindowProc_ufCommand)
'done:
'  Exit Sub
'eh:
'  RaiseError "Hooking.HookForm_ufCommand", Err.Number, Err.Source, Err.description, Erl
'End Sub
'Public Sub UnHookForm_ufCommand()
'  On Error GoTo eh
'  SetWindowLong ufCommand.hWnd, GWL_WNDPROC, PrevProc_ufCommand
'done:
'  Exit Sub
'eh:
'  RaiseError "Hooking.UnHookForm_ufCommand", Err.Number, Err.Source, Err.description, Erl
'End Sub
'Public Function WindowProc_ufCommand(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'  On Error GoTo eh
'  WindowProc_ufCommand = CallWindowProc(PrevProc_ufCommand, hWnd, uMsg, wParam, lParam)
'
'  If uMsg = WM_MOVING Then 'WM_WINDOWPOSCHANGING
'      'Debug.Print "Moving " & Form1.Left / 15 & ", " & Form1.Top / 15
'    ufListButtons.UpdatePosition
'  End If
'done:
'  Exit Function
'eh:
'  RaiseError "Hooking.WindowProc_ufCommand", Err.Number, Err.Source, Err.description, Erl
'  UnHookForm_ufCommand
'End Function



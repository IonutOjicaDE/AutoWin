VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufRecorder 
   Caption         =   "Recording window"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ufRecorder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' https://stackoverflow.com/questions/49901675/why-windows-mouse-hook-in-excel-vba-slows-down-excel-if-there-isnt-userform


'* System hook for mouse that we'll use
Const WH_MOUSE = 7 ' does not work
Const WH_MOUSE_LL = 14
Const WH_KEYBOARD_LL = 13
Const WH_KEYBOARD = 2

Public Function RemoveHook() As Boolean
  On Error Resume Next
  RemoveHook = UnhookWindowsHookEx(MouseHookHandle)
  ' Debug.Print "RemoveHook from Key: " & RemoveHook
  If RemoveHook Then MouseHookHandle = NULL_

  RemoveHook = UnhookWindowsHookEx(KeyHookHandle)
  ' Debug.Print "RemoveHook from Key: " & RemoveHook
  If RemoveHook Then KeyHookHandle = NULL_
End Function

'WH_CALLWNDPROC

Public Sub SetHook()
  If MouseHookHandle = NULL_ Then
    MouseHookHandle = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseHookProc, 0, 0)
    'Debug.Print "MouseHookHandle = " & MouseHookHandle & " , Err.LastDllError = " & Err.LastDllError, " , InstallHook = " & (MouseHookHandle <> NULL_)
  Else
    MsgBox "Already hooked!" & vbCrLf & "MouseHookHandle = " & MouseHookHandle, vbCritical
  End If
  If KeyHookHandle = NULL_ Then
    KeyHookHandle = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyHookProc, 0, 0)
    'Debug.Print "KeyHookHandle = " & KeyHookHandle & " , Err.LastDllError = " & Err.LastDllError, " , InstallHook = " & (KeyHookHandle <> NULL_)
  Else
    MsgBox "Already hooked!" & vbCrLf & "KeyHookHandle = " & KeyHookHandle, vbCritical
  End If

' don't focus too much on the term, it doesn't clarify anything.
' There's a huge difference between the two.
' WH_KEYBOARD_LL installs a hook that requires the callback to be implemented in your own program. And you must pump a message loop so that Windows can make the callback whenever it is about to dispatch a keyboard message. Which makes it really easy to get going.
' WH_KEYBOARD works very differently, it requires a DLL that can be safely injected into hooked processes. Which makes it notoriously difficult to get going, injecting DLLs without affecting a process isn't easy. Particularly on a 64-bit operating system. Nor is taking care of the inter-process communication you might need if some other process needs to know about the keystroke. Like a key logger.
' The advantage of WH_KEYBOARD is that it has access to the keyboard state. Which is a per-process property in Windows. State like the active keyboard layout and the state of the modifier and dead keys matter a great deal when you want to use the hook to translate virtual keys to typing keys yourself. You can't reliably call ToUnicodeEx() from an external process.

' https://www.reddit.com/r/vba/comments/4q5xac/keyboard_hook_works_everywhere_but_the_vbe/
' https://stackoverflow.com/questions/17502485/how-do-i-use-setwindowshookex-to-filter-low-level-key-events <= c++

End Sub

Private Sub PlayPauseButton_Click()
  Select Case PlayPauseButton.Caption
    Case "Continue":
      PlayPauseButton.Caption = "Pause"
      StatusLabel.Caption = "Recording..."
      Sleep 500
      SetHook
    Case "Pause":
      PlayPauseButton.Caption = "Continue"
      StatusLabel.Caption = "Paused."
      RemoveHook
  End Select
End Sub

Private Sub StopButton_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  SetHook
End Sub

Private Sub UserForm_Terminate()
  RemoveHook
End Sub


'https://docs.microsoft.com/en-us/windows/win32/winmsg/about-hooks

'https://www.vbforums.com/showthread.php?642587-RESOLVED-Hook-App-HELP
'You may want to play with SetWindowsHookEx API, specifically one of these
'1) WH_CBT and looking for the HCBT_CREATEWND message. Window is not yet displayed when this is called. Not limited to top-level windows
'https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644977(v=vs.85)

'2) WH_SHELL and looking for the HSHELL_WINDOWCREATED message. Top level windows only.
'Both would need to be global hooks vs. thread. I don't recall off hand whether the callbacks for these specific global types can exist in a module vs. a non-ActiveX DLL. So playing & some research may be needed.





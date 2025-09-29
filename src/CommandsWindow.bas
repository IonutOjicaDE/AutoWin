Attribute VB_Name = "CommandsWindow"
Option Explicit
Private Const MODULE_NAME As String = "CommandsWindow"
Option Base 0
Option Compare Text 'for Like: (A=a) < (À=à) < (B=b) < (E=e) < (Ê=ê) < (Z=z) < (Ø=ø)
'Option Compare Binary 'for Like: A < B < E < Z < a < b < e < z < À < Ê < Ø < à < ê < ø

'Like pattern: ?=any single char, *=zero or more chars, #=any single digit (0-9)
'[charlist]=any single char in charlist, [!charlist] ; [A-Z]
'MyCheck = "aBBBa" Like "a*a"          ' Returns True
'MyCheck = "F" Like "[A-Z]"            ' Returns True
'MyCheck = "F" Like "[!A-Z]"           ' Returns False
'MyCheck = "a2a" Like "a#a"            ' Returns True
'MyCheck = "aM5b" Like "a[L-P]#[!c-e]" ' Returns True
'MyCheck = "BAT123khg" Like "B?T*"     ' Returns True
'MyCheck = "CAT123khg" Like "B?T*"     ' Returns False
'MyCheck = "ab" Like "a*b"             ' Returns True
'MyCheck = "a*b" Like "a [*]b"         ' Returns False
'MyCheck = "axxxxxb" Like "a [*]b"     ' Returns False
'MyCheck = "a [xyz" Like "a [[]*"      ' Returns True
'MyCheck = "a [xyz" Like "a [*"        ' Throws Error 93 (invalid pattern string)
'https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/like-operator


Private Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type


Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
' FindWindowEx:
' hWnd1 – handle of the parrent (or 0 for desktop).
' hWnd2 – handle of the last child found (to continue the search).
' lpsz1 – the class of the child window. Must match exactly, without wildcards.
' lpsz2 – title of the child window. Must match exactly, without wildcards.
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As LongPtr, ByVal wFlag As Long) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hWnd As LongPtr, ByVal gaFlags As Long) As LongPtr ' https://www.vbarchiv.net/api/api_getancestor.html
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
        
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

#If Win64 Then
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
#Else
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
#End If

Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long ' https://www.vbarchiv.net/api/api_showwindow.html
Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hWnd As LongPtr) As Long 'Determines whether a window is maximized.

Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
   ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function GetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare PtrSafe Function SetWindowPlacement Lib "user32" (ByVal hWnd As LongPtr, lpwndpl As WINDOWPLACEMENT) As Long
    
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private Declare PtrSafe Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As LongPtr, ByVal hModule As LongPtr, ByVal lpFilename As String, ByVal nSize As Long) As Long


Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Enum Process_Constants ' for OpenProcess, dwDesiredAccess
  PROCESS_TERMINATE = &H1                    ' Required to terminate a process using TerminateProcess.
  PROCESS_CREATE_THREAD = &H2                ' Required to create a thread in the process.
  PROCESS_VM_OPERATION = &H8                 ' Required to perform an operation on the address space of a process (see VirtualProtectEx and WriteProcessMemory).
  PROCESS_VM_READ = &H10                     ' Required to read memory in a process using ReadProcessMemory.
  PROCESS_VM_WRITE = &H20                    ' Required to write to memory in a process using WriteProcessMemory.
  PROCESS_DUP_HANDLE = &H40                  ' Required to duplicate a handle using DuplicateHandle.
  PROCESS_CREATE_PROCESS = &H80              ' Required to use this process as the parent process with PROC_THREAD_ATTRIBUTE_PARENT_PROCESS.
  PROCESS_SET_QUOTA = &H100                  ' Required to set memory limits using SetProcessWorkingSetSize.
  PROCESS_SET_INFORMATION = &H200            ' Required to set certain information about a process, such as its priority class (see SetPriorityClass).
  PROCESS_QUERY_INFORMATION = &H400          ' Required to retrieve certain information about a process, such as its token, exit code, and priority class (see OpenProcessToken).
  PROCESS_SUSPEND_RESUME = &H800             ' Required to suspend or resume a process.
  PROCESS_QUERY_LIMITED_INFORMATION = &H1000 ' Required to retrieve certain information about a process (see GetExitCodeProcess, GetPriorityClass, IsProcessInJob, QueryFullProcessImageName). A handle that has the PROCESS_QUERY_INFORMATION access right is automatically granted PROCESS_QUERY_LIMITED_INFORMATION
  PROCESS_ALL_ACCESS = &HFFFF                ' All possible access rights for a process object.
End Enum

Private Enum GW_Constants     ' for GetNextWindow, wFlag
  GW_HWNDFIRST = 0&            ' The retrieved handle identifies the window of the same type that is highest in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDLAST = 1&             ' The retrieved handle identifies the window of the same type that is lowest in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDNEXT = 2&             ' The retrieved handle identifies the window below the specified window in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_HWNDPREV = 3&             ' The retrieved handle identifies the window above the specified window in the Z order. If the specified window is a topmost window, the handle identifies a topmost window. If the specified window is a top-level window, the handle identifies a top-level window. If the specified window is a child window, the handle identifies a sibling window.
  GW_OWNER = 4&                ' The retrieved handle identifies the specified window's owner window, if any. For more information, see Owned Windows.
  GW_CHILD = 5&                ' The retrieved handle identifies the child window at the top of the Z order, if the specified window is a parent window; otherwise, the retrieved handle is NULL. The function examines only child windows of the specified window. It does not examine descendant windows.
  GW_ENABLEDPOPUP = 6&         ' The retrieved handle identifies the enabled popup window owned by the specified window (the search uses the first such window found using GW_HWNDNEXT); otherwise, if there are no enabled popup windows, the retrieved handle is that of the specified window.
End Enum

Private Enum GWL_Constants    ' Window field offsets for GetWindowLong() and GetWindowWord()
  GWL_WNDPROC = -4&
  GWL_HINSTANCE = -6&
  GWL_HWNDPARENT = -8&
  GWL_STYLE = -16&
  GWL_EXSTYLE = -20&
  GWL_USERDATA = -21&
  GWL_ID = -12&
End Enum

Private Enum WS_Constants     ' Window Styles - result of GetWindowLong() or GetWindowLongPtr()
  WS_OVERLAPPED = &H0&
  WS_POPUP = &H80000000
  WS_CHILD = &H40000000
  WS_MINIMIZE = &H20000000
  WS_VISIBLE = &H10000000
  WS_DISABLED = &H8000000
  WS_CLIPSIBLINGS = &H4000000
  WS_CLIPCHILDREN = &H2000000
  WS_MAXIMIZE = &H1000000
  WS_CAPTION = &HC00000       ' WS_BORDER Or WS_DLGFRAME
  WS_BORDER = &H800000
  WS_DLGFRAME = &H400000
  WS_VSCROLL = &H200000
  WS_HSCROLL = &H100000
  WS_SYSMENU = &H80000
  WS_THICKFRAME = &H40000
  WS_GROUP = &H20000
  WS_TABSTOP = &H10000
  
  WS_MINIMIZEBOX = &H20000
  WS_MAXIMIZEBOX = &H10000
  
  WS_TILED = WS_OVERLAPPED
  WS_ICONIC = WS_MINIMIZE
  WS_SIZEBOX = WS_THICKFRAME
  WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
  WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
  '   Common Window Styles
  WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
  WS_CHILDWINDOW = WS_CHILD
End Enum

Private Enum HWND_Constants   ' for SetWindowPos, hWndInsertAfter; https://www.activevb.de/cgi-bin/apiwiki/SetWindowPos
  HWND_BOTTOM = 1&             ' Places the window at the bottom of the Z order. If the hWnd parameter identifies a topmost window, the window loses its topmost status and is placed at the bottom of all other windows
  HWND_NOTOPMOST = -2&         ' Places the window above all non-topmost windows (that is, behind all topmost windows). This flag has no effect if the window is already a non-topmost window
  HWND_TOP = 0&                ' Places the window at the top of the Z order
  HWND_TOPMOST = -1&           ' Places the window above all non-topmost windows. The window maintains its topmost position even when it is deactivated
End Enum

Private Enum SWP_Constants    ' for SetWindowPos, wFlags
  SWP_NOSIZE = &H1&            ' Retains/doesn't change the current size (ignores the cx and cy parameters; cx and cy can be set to 0).
  SWP_NOMOVE = &H2&            ' Retains the current position (ignores X and Y parameters).
  SWP_NOZORDER = &H4&          ' Retains the current Z order (ignores the hWndInsertAfter parameter).
  SWP_NOREDRAW = &H8&          ' Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent window uncovered as a result of the window being moved. When this flag is set, the application must explicitly invalidate or redraw any parts of the window and parent window that need redrawing.
  SWP_NOACTIVATE = &H10&       ' Does not activate the window. If this flag is not set, the window is activated (focused) and moved to the top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter parameter).
  SWP_FRAMECHANGED = &H20&     ' Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE is sent only when the window's size is being changed.
  SWP_SHOWWINDOW = &H40&       ' Displays the window. Same as .Show(), respectivelly change of the Visible-Parameter.
  SWP_HIDEWINDOW = &H80&       ' Hides the window. Same as .Hide(), respectivelly change of the Visible-Parameter.
  SWP_NOCOPYBITS = &H100&      ' Discards the entire contents of the client area, requiring a redraw. If this flag is not specified, the valid contents of the client area are saved and copied back into the client area after the window is sized or repositioned.
  SWP_NOOWNERZORDER = &H200&   ' Does not change the owner window's position in the Z order.
  SWP_NOSENDCHANGING = &H400&  ' Prevents the window from receiving the WM_WINDOWPOSCHANGING message.
  SWP_DEFERERASE = &H2000&     ' Prevents generation of the WM_SYNCPAINT message.
  SWP_ASYNCWINDOWPOS = &H4000& ' If the calling thread and the thread that owns the window are attached to different input queues, the system posts the request to the thread that owns the window. This prevents the calling thread from blocking its execution while other threads process the request.
End Enum

Private Enum SW_Constants     ' for ShowWindow, nCmdShow; https://www.vbarchiv.net/api/api_showwindow.html
  SW_HIDE = 0&                 ' Hides the window and activates another window
  SW_SHOWNORMAL = 1&           ' Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position
  SW_SHOWMINIMIZED = 2&        ' Activates the window and displays it as a minimized window
  SW_SHOWMAXIMIZED = 3&        ' Activates the window and displays it as a maximized window
  SW_SHOWNOACTIVATE = 4&       ' Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated
  SW_SHOW = 5&                 ' Activates the window and displays it in its current size and position
  SW_MINIMIZE = 6&             ' Minimizes the specified window and activates the next top-level window in the Z order
  SW_SHOWMINNOACTIVE = 7&      ' Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated
  SW_SHOWNA = 8&               ' Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated
  SW_RESTORE = 9&              ' Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window
  SW_SHOWDEFAULT = 10&         ' Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application
  SW_FORCEMINIMIZE = 11&       ' Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread
  SW_NORMAL = 1&               ' SW_SHOWNORMAL
  SW_MAXIMIZE = 3&             ' SW_SHOWMAXIMIZED
End Enum

Private Enum GA_Constants     ' for GetAncestor, gaFlags
  GA_PARENT = 1&               ' Retrieves the parent window. This does not include the owner, as it does with the GetParent function.
  GA_ROOT = 2&                 ' Retrieves the root window by walking the chain of parent windows.
  GA_ROOTOWNER = 3&            ' Retrieves the owned root window by walking the chain of parent and owner windows returned by GetParent.
End Enum


Private Const MAXLEN = 255
Private rc              As RECT
Private Title           As String * MAXLEN
Private remoteThreadId  As Long
Private currentThreadId As Long
Private MousePos        As POINTAPI
Private WndPlcmt        As WINDOWPLACEMENT
Private tmpL            As Long
Private tmpS            As String

Private foundHwnd       As LongPtr ' for EnumChildProc
Private searchTitle     As String  ' for EnumChildProc
Private searchClass     As String  ' for EnumChildProc

Private hDesktopWnd     As LongPtr, hTmpWnd As LongPtr


Public Sub RegisterCommandsWindow()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "getactivewindowname", Array("GetActiveWindowName", "Get Active Window Name", _
    MODULE_NAME, "Write in first argument the title of the active window", _
    "Title", "Retrieves the title of the current window")

  commandMap.Add "getactivewindowclass", Array("GetActiveWindowClass", "Get Active Window Class", _
    MODULE_NAME, "Write in first argument the class of the active window", _
    "Class", "Retrieves the class of the current window")

  commandMap.Add "getwindownamefromclass", Array("GetWindowNameFromClass", "Get Window Name from Class", _
    MODULE_NAME, "Write in the second argument the title of the window that has the class mentioned in the first argument", _
    "Class", "Enter the class of the window to read the title", _
    "Title", "Retrieves the title of the window with the mentioned class")


  commandMap.Add "setactivewindowbyname", Array("SetActiveWindowByName", "Set Active Window By Name", _
    MODULE_NAME, "Activate the window that is found to contain the text from the first argument", _
    "Title", "Text to be contained by the window title to activate")

  commandMap.Add "activatewindowbyname", Array("SetActiveWindowByName", "Activate Window By Name", _
    MODULE_NAME, "Activate the window that is found to contain the text from the first argument", _
    "Title", "Text to be contained by the window title to activate")


  commandMap.Add "getwindowposition", Array("GetWindowPosition", "Get Window Position", _
    MODULE_NAME, "Retrieves position of the window that has the title specified in first argument", _
    "Title", "Text to be contained by the window title; if no title provided, the active window will be chosen", _
    "Left", "x1 value, smallest x value (Top)", _
    "Top", "y1 value, smallest y value (Left)", _
    "Right", "x2 value, biggest x value (Right)", _
    "Unten", "y2 value, biggest y value (Bottom)")

  commandMap.Add "setwindowposition", Array("SetWindowPosition", "Set Window Position", _
    MODULE_NAME, "Set the position of the window that has the title specified in first argument", _
    "Title", "Text to be contained by the window title; if no title provided, the active window will be chosen", _
    "Left", "x1 value, smallest x value (Top); if x1 or y1 are not provided, then the window will be resized, by keeping current x1 and y1 of the window", _
    "Top", "y1 value, smallest y value (Left); if x1 or y1 are not provided, then the window will be resized, by keeping current x1 and y1 of the window", _
    "Right", "x2 value, biggest x value (Right); if x2 or y2 are not provided, then the window will be moved, by keeping current size of the window", _
    "Unten", "y2 value, biggest y value (Bottom); if x2 or y2 are not provided, then the window will be moved, by keeping current size of the window")


  commandMap.Add "windowrestore", Array("WindowRestore", "Window Restore", _
    MODULE_NAME, "Restore the window", _
    "Title", "Text to be contained by the window title")

  commandMap.Add "windowminimize", Array("WindowMinimize", "Window Minimize", _
    MODULE_NAME, "Minimize the window", _
    "Title", "Text to be contained by the window title")

  commandMap.Add "windowmaximize", Array("WindowMaximize", "Window Maximize", _
    MODULE_NAME, "Maximize the window", _
    "Title", "Text to be contained by the window title")

  commandMap.Add "getwindowstate", Array("GetWindowState", "Get Window State", _
    MODULE_NAME, "Retrieve the state of the window", _
    "Title", "Text to be contained by the window title", _
    "State", "Retrieve the state of the window: minimized, maximized, hidden, normal")

  hDesktopWnd = GetDesktopWindow()

  remoteThreadId = 0&
  currentThreadId = 0&

  WndPlcmt.Length = Len(WndPlcmt)
done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsWindow", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsWindow()
  On Error GoTo eh

  Call RemoveAttachedThread
done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsWindow", Err.Number, Err.Source, Err.description, Erl
End Sub



'#############################################
'#######                               #######
'#######      WaitWindowToActivate     #######
'#######                               #######
'#############################################

Public Function WaitWindowToActivate(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Static WindowName As String: WindowName = currentRowArray(1, ColAWindow)
  If Len(WindowName) Then
    Static ActiveWindowHwnd   As LongPtr
    Static timeToCheckAndWait As Long
    Static activeWindowTitle  As String
    ' Check foreground window
    timeToCheckAndWait = windowCheckMax
    Do Until timeToCheckAndWait <= 0
      ActiveWindowHwnd = GetForegroundWindow()
      activeWindowTitle = GetTitleFromHwnd(ActiveWindowHwnd)
      If activeWindowTitle Like WindowName Then
        WaitWindowToActivate = True
        Exit Function
      End If
      Sleep minLong(timeToCheckAndWait, windowCheckSplit)
      timeToCheckAndWait = timeToCheckAndWait - windowCheckSplit
      DoEvents
    Loop

    ' Check all other windows
    Do
      timeToCheckAndWait = windowCheckMax
      Do Until timeToCheckAndWait <= 0
        If ActivateWindowByHandle(GetHwndWithTitle(WindowName)) <> NULL_ Then
          DoEvents
          Sleep windowCheckSplit
          DoEvents
          WaitWindowToActivate = True
          Exit Function
        End If
      Loop

      ' Window not found, check again?
      Select Case AskNextStep("No window with title " & WindowName & " found. You can open the window manually then retry the checking", vbAbortRetryIgnore, "Window not found")
        Case vbIgnore:
          WaitWindowToActivate = True
          Exit Function
        Case vbAbort:
          WaitWindowToActivate = False
          Exit Function
        Case vbRetry:
          ' check again
      End Select
    Loop While True
    WaitWindowToActivate = False
  Else
    WaitWindowToActivate = True
  End If

done:
  Exit Function
eh:
  WaitWindowToActivate = False
  RaiseError MODULE_NAME & ".WaitWindowToActivate", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'#############################################
'#######                               #######
'#######          AskNextStep          #######
'#######                               #######
'#############################################

Public Function AskNextStep(Prompt As String, Buttons As VbMsgBoxStyle, Title As String) As VbMsgBoxResult
  On Error GoTo eh
  Static ActiveWindowHwnd As LongPtr: ActiveWindowHwnd = GetForegroundWindow()
  AskNextStep = MsgBox(Prompt, Buttons, Title)
  Call Sleep(windowCheckSplit)
  DoEvents
  Call ActivateWindowByHandle(ActiveWindowHwnd)
  Call Sleep(windowCheckSplit)
  DoEvents

done:
  Exit Function
eh:
  AskNextStep = vbAbort
  RaiseError MODULE_NAME & ".AskNextStep", Err.Number, Err.Source, Err.description, Erl
End Function

'#############################################
'#######                               #######
'#######      GetActiveWindowName      #######
'#######                               #######
'#############################################
Public Function GetActiveWindowName(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName
  On Error GoTo eh
  currentRowRange(1, ColAArg1).Value = GetTitleFromHwnd(GetForegroundWindow())

done:
  GetActiveWindowName = True
  Exit Function
eh:
  GetActiveWindowName = False
  RaiseError MODULE_NAME & ".GetActiveWindowName", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function GetActiveWindowClass(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=ClassName
  On Error GoTo eh
  currentRowRange(1, ColAArg1).Value = GetClassFromHwnd(GetForegroundWindow())

done:
  GetActiveWindowClass = True
  Exit Function
eh:
  GetActiveWindowClass = False
  RaiseError MODULE_NAME & ".GetActiveWindowClass", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function GetWindowNameFromClass(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=ClassName
  On Error GoTo eh
  Dim className As String: className = currentRowArray(1, ColAArg1)
  If Len(className) = 0 Then
    GetWindowNameFromClass = False
    RaiseError MODULE_NAME & ".GetWindowNameFromClass", Err.Number, Err.Source, "Classname not provided", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If
  

  hTmpWnd = GetNextWindow(hDesktopWnd, GW_CHILD)
  Do While hTmpWnd <> NULL_
    tmpS = GetClassFromHwnd(hTmpWnd)
    If GetWindowLong(hTmpWnd, GWL_STYLE) And WS_VISIBLE Then
      If tmpS Like className Then Exit Do
    End If
    hTmpWnd = GetNextWindow(hTmpWnd, GW_HWNDNEXT)
  Loop
  If hTmpWnd = Null Then
    GetWindowNameFromClass = False
    RaiseError MODULE_NAME & ".GetWindowNameFromClass", Err.Number, Err.Source, "No window found with Classname provided: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If
  currentRowRange(1, ColAArg1 + 1).Value = GetTitleFromHwnd(hTmpWnd)

done:
  GetWindowNameFromClass = True
  Exit Function
eh:
  GetWindowNameFromClass = False
  RaiseError MODULE_NAME & ".GetWindowNameFromClass", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'#############################################
'#######                               #######
'#######     SetActiveWindowByName     #######
'#######                               #######
'#############################################
Public Function SetActiveWindowByName(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName
  On Error GoTo eh
  Dim WindowName As String: WindowName = currentRowArray(1, ColAArg1)
  If Len(WindowName) = 0 Then
    SetActiveWindowByName = False
    RaiseError MODULE_NAME & ".SetActiveWindowByName", Err.Number, Err.Source, "Window name not provided", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

  hTmpWnd = GetNextWindow(hDesktopWnd, GW_CHILD)
  Do While hTmpWnd <> NULL_
    tmpS = GetTitleFromHwnd(hTmpWnd)
    If GetWindowLong(hTmpWnd, GWL_STYLE) And WS_VISIBLE Then
      If tmpS Like WindowName Then Exit Do
    End If
    hTmpWnd = GetNextWindow(hTmpWnd, GW_HWNDNEXT)
  Loop

  If ActivateWindowByHandle(hTmpWnd) = NULL_ Then
    SetActiveWindowByName = False
    RaiseError MODULE_NAME & ".SetActiveWindowByName", Err.Number, Err.Source, "No window found with name provided: Arg1=[" & CStr(currentRowArray(1, ColAArg1)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

done:
  SetActiveWindowByName = True
  Exit Function
eh:
  SetActiveWindowByName = False
  RaiseError MODULE_NAME & ".SetActiveWindowByName", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


'#############################################
'#######                               #######
'#######         WindowPosition        #######
'#######                               #######
'#############################################
Public Function GetWindowPosition(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName, when not provided, the active window will be chosen
' Arg2=Left, Arg3=Top, Arg4=Right, Arg5=Bottom
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  GetWindowRect hTmpWnd, rc
  currentRowRange(1, ColAArg1 + 1).Value = rc.Left
  currentRowRange(1, ColAArg1 + 2).Value = rc.Top
  currentRowRange(1, ColAArg1 + 3).Value = rc.Right
  currentRowRange(1, ColAArg1 + 4).Value = rc.Bottom

done:
  GetWindowPosition = True
  Exit Function
eh:
  GetWindowPosition = False
  RaiseError MODULE_NAME & ".GetWindowPosition", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

Public Function SetWindowPosition(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName, Arg2=Left, Arg3=Top, Arg4=Right, Arg5=Bottom
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  If IsIconic(hTmpWnd) Or IsZoomed(hTmpWnd) Then ShowWindow hTmpWnd, SW_RESTORE

  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 1))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 2))) Then
    If IsNumber(CStr(currentRowArray(1, ColAArg1 + 3))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 4))) Then
      SetWindowPos hTmpWnd, 0&, CLng(currentRowArray(1, ColAArg1 + 1)), CLng(currentRowArray(1, ColAArg1 + 2)), CLng(currentRowArray(1, ColAArg1 + 3)), CLng(currentRowArray(1, ColAArg1 + 4)), SWP_NOZORDER

    Else ' Window will be moved to new location, keeping the width and height
      GetWindowRect hTmpWnd, rc
      SetWindowPos hTmpWnd, 0&, CLng(currentRowArray(1, ColAArg1 + 1)), CLng(currentRowArray(1, ColAArg1 + 2)), rc.Right - rc.Left, rc.Bottom - rc.Top, SWP_NOZORDER
    End If

  ElseIf IsNumber(CStr(currentRowArray(1, ColAArg1 + 3))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 4))) Then
    ' Window size will be changed
    GetWindowRect hTmpWnd, rc
    SetWindowPos hTmpWnd, 0&, rc.Left, rc.Top, CLng(currentRowArray(1, ColAArg1 + 3)), CLng(currentRowArray(1, ColAArg1 + 4)), SWP_NOZORDER
  
  Else ' no dimensions mentioned
    SetWindowPosition = False
    RaiseError MODULE_NAME & ".SetWindowPosition", Err.Number, Err.Source, "Please enter the new location as numbers: x1=Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "] y1=Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "] x2=Arg4=[" & CStr(currentRowArray(1, ColAArg1 + 3)) & "] y2=Arg5=[" & CStr(currentRowArray(1, ColAArg1 + 4)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If

done:
  SetWindowPosition = True
  Exit Function
eh:
  SetWindowPosition = False
  RaiseError MODULE_NAME & ".SetWindowPosition", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'#############################################
'#######                               #######
'#######         Window Restore        #######
'#######                               #######
'#############################################
Public Function WindowRestore(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName
'https://activevb.de/tipps/vb6tipps/tipp0214.html
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  With WndPlcmt
    .Length = Len(WndPlcmt)
    If GetWindowPlacement(hTmpWnd, WndPlcmt) Then
      If .showCmd = SW_SHOWMINIMIZED Then
        .flags = 0&
        .showCmd = SW_SHOWNORMAL
        Call SetWindowPlacement(hTmpWnd, WndPlcmt)
      Else
        Call SetForegroundWindow(hTmpWnd)
        Call BringWindowToTop(hTmpWnd)
      End If
    End If
  End With

done:
  WindowRestore = True
  Exit Function
eh:
  WindowRestore = False
  RaiseError MODULE_NAME & ".WindowRestore", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
'#############################################
Public Function WindowMinimize(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  ShowWindow hTmpWnd, SW_MINIMIZE 'Minimizes the specified window and activates the next top-level window in the Z order

done:
  WindowMinimize = True
  Exit Function
eh:
  WindowMinimize = False
  RaiseError MODULE_NAME & ".WindowMinimize", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
'#############################################
Public Function WindowMaximize(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  ShowWindow hTmpWnd, SW_SHOWMAXIMIZED 'Activates the window and displays it as a maximized window

done:
  WindowMaximize = True
  Exit Function
eh:
  WindowMaximize = False
  RaiseError MODULE_NAME & ".WindowMaximize", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
'#############################################
Public Function GetWindowState(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' Arg1=WindowName, Arg2=State
  On Error GoTo eh
  hTmpWnd = GetHwndWithTitleOrActiveWindow(CStr(currentRowArray(1, ColAArg1)))

  If IsIconic(hTmpWnd) Then
    currentRowRange(1, ColAArg1 + 1).Value = "minimized"
  ElseIf IsZoomed(hTmpWnd) Then
    currentRowRange(1, ColAArg1 + 1).Value = "maximized"
  ElseIf IsWindowVisible(hTmpWnd) Then
    currentRowRange(1, ColAArg1 + 1).Value = "hidden"
  Else
    currentRowRange(1, ColAArg1 + 1).Value = "normal"
  End If

done:
  GetWindowState = True
  Exit Function
eh:
  GetWindowState = False
  RaiseError MODULE_NAME & ".GetWindowState", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


'#############################################
'#######                               #######
'#######          T H R E A D          #######
'#######                               #######
'#############################################


Private Sub AttachThreadToWindow(ByRef hWnd As LongPtr)
'AttachThreadInput is needed so we can get the handle of a focused window in another app
  Dim tmpL As Long
  tmpL = GetWindowThreadProcessId(hWnd, 0&)
  If tmpL = remoteThreadId Then Exit Sub
  If remoteThreadId <> 0& Then RemoveAttachedThread
  remoteThreadId = tmpL
  If currentThreadId = 0& Then currentThreadId = GetCurrentThreadId
  If currentThreadId = remoteThreadId Then
    remoteThreadId = 0&
    Exit Sub
  End If
  AttachThreadInput remoteThreadId, currentThreadId, True
End Sub
'#############################################
Public Sub AttachThreadToWindowFromPoint(ByRef x As Long, ByRef y As Long)
'AttachThreadInput is needed so we can get the handle of a focused window in another app
  AttachThreadToWindow WindowFromPoint(x, y)
End Sub
'#############################################
Public Sub AttachThreadToActiveWindow()
'AttachThreadInput is needed so we can get the handle of a focused window in another app
  AttachThreadToWindow GetForegroundWindow()
End Sub
'#############################################
Public Sub RemoveAttachedThread()
  If currentThreadId = 0& Or remoteThreadId = 0& Then Exit Sub
  AttachThreadInput remoteThreadId, currentThreadId, False
  remoteThreadId = 0&
End Sub




'#############################################
' Callback function for EnumChildWindows
Private Function EnumChildProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
  Dim winText As String
  Dim classText As String
  Dim titleMatch As Boolean, classMatch As Boolean

  winText = GetTitleFromHwnd(hWnd)
  classText = GetClassFromHwnd(hWnd)

  If searchTitle <> "" Then
    titleMatch = (InStr(1, winText, searchTitle, vbTextCompare) > 0)
  Else
    titleMatch = True
  End If

  If searchClass <> "" Then
    classMatch = (InStr(1, classText, searchClass, vbTextCompare) > 0)
  Else
    classMatch = True
  End If

  If titleMatch And classMatch Then
    foundHwnd = hWnd
    EnumChildProc = 0 ' stop enumeration
    Exit Function
  End If

  EnumChildProc = 1 ' continue enumeration
End Function

'#############################################
' NON-recursive function: search first child with partial match
Public Function FindChildWindowPartial(ByVal hWndParent As LongPtr, _
                                       Optional ByVal partialClass As String = "", _
                                       Optional ByVal partialTitle As String = "") As LongPtr
  If partialClass = "" And partialTitle = "" Then Exit Function

  foundHwnd = 0
  searchTitle = partialTitle
  searchClass = partialClass

  EnumChildWindows hWndParent, AddressOf EnumChildProc, 0
  FindChildWindowPartial = foundHwnd
End Function



'#############################################
Private Function FindChildWindowRecursiveInternal(hWndParent As LongPtr) As Boolean
  Dim childHwnd As LongPtr
  childHwnd = FindChildWindowPartial(hWndParent, searchClass, searchTitle)

  If childHwnd <> 0 Then
    foundHwnd = childHwnd
    FindChildWindowRecursiveInternal = True
    Exit Function
  End If

  Dim hChild As LongPtr
  hChild = 0
  Do
    'hChild = FindChildWindowPartial(hWndParent,, ) ' get next child (non-filtered)
    If hChild = 0 Then Exit Do

    If FindChildWindowRecursiveInternal(hChild) Then
      FindChildWindowRecursiveInternal = True
      Exit Function
    End If
  Loop
End Function
'#############################################
Public Function FindChildWindowRecursive(ByVal hWndParent As LongPtr, _
                                         Optional ByVal partialClass As String = "", _
                                         Optional ByVal partialTitle As String = "") As LongPtr
  If partialClass = "" And partialTitle = "" Then Exit Function

  foundHwnd = 0
  searchTitle = partialTitle
  searchClass = partialClass

  FindChildWindowRecursiveInternal hWndParent
  FindChildWindowRecursive = foundHwnd
End Function






'#############################################
Public Function GetWindowTitleFromPoint(ByRef x As Long, ByRef y As Long) As String
  GetWindowTitleFromPoint = GetTitleFromHwnd(WindowFromPoint(x, y))
End Function
Public Function GetWindowTitleUnderCursor() As String
  Dim MousePos As POINTAPI
  GetCursorPos MousePos
  GetWindowTitleUnderCursor = GetTitleFromHwnd(WindowFromPoint(MousePos.x, MousePos.y))
End Function
Public Function GetActiveWindowTitle() As String
  GetActiveWindowTitle = GetTitleFromHwnd(GetForegroundWindow())
End Function
Public Function GetExeFilenameFromWindowTitle(ByRef sTitle As String) As String
  GetExeFilenameFromWindowTitle = GetExeFilenameFromHwnd(GetHwndWithTitle(sTitle))
End Function
'#############################################
Private Function GetHwndWithTitleOrActiveWindow(ByRef sTitle As String) As LongPtr
  On Error GoTo eh
  If Len(sTitle) = 0 Then
    GetHwndWithTitleOrActiveWindow = GetForegroundWindow()
  Else
    GetHwndWithTitleOrActiveWindow = GetHwndWithTitle(sTitle)

    If GetHwndWithTitleOrActiveWindow = NULL_ Then
      RaiseError MODULE_NAME & ".GetHwndWithTitleOrActiveWindow", Err.Number, Err.Source, "No window found with name provided: [" & sTitle & "]", Erl, 1
      Exit Function
    End If
  End If

done:
  Exit Function
eh:
  GetHwndWithTitleOrActiveWindow = NULL_
  RaiseError MODULE_NAME & ".GetHwndWithTitleOrActiveWindow", Err.Number, Err.Source, Err.description, Erl
End Function
'#############################################
Private Function GetHwndWithTitle(ByRef sTitle As String) As LongPtr
  If Len(sTitle) = 0 Then
    GetHwndWithTitle = NULL_
  Else
    Dim CurrentWindowTitel As String
    hTmpWnd = GetNextWindow(hDesktopWnd, GW_CHILD)
    Do While hTmpWnd <> NULL_
      CurrentWindowTitel = GetTitleFromHwnd(hTmpWnd)
      If GetWindowLong(hTmpWnd, GWL_STYLE) And WS_VISIBLE Then
        If CurrentWindowTitel Like sTitle Then Exit Do
      End If
      hTmpWnd = GetNextWindow(hTmpWnd, GW_HWNDNEXT)
    Loop
    GetHwndWithTitle = hTmpWnd
  End If
End Function
'#############################################
Private Function ActivateWindowByHandle(ByRef hWnd As LongPtr) As LongPtr
  If hWnd = NULL_ Then GoTo ErrorTrigger
  On Error GoTo ErrorTrigger
  SetWindowPos hWnd, NULL_, 0&, 0&, 0&, 0&, SWP_NOSIZE + SWP_NOMOVE + SWP_SHOWWINDOW
  SetForegroundWindow hWnd
  ActivateWindowByHandle = hWnd
  Exit Function
ErrorTrigger:
  ActivateWindowByHandle = NULL_
  Err.Clear
End Function
'#############################################
Private Function GetTitleFromHwnd(ByRef hWnd As LongPtr) As String
  If hWnd = NULL_ Then
    GetTitleFromHwnd = vbNullString
  Else
    Static LengthOfText As Long
    LengthOfText = GetWindowText(hWnd, Title, Len(Title))
    GetTitleFromHwnd = Left$(Title, LengthOfText)
  End If
End Function
'#############################################
Private Function GetClassFromHwnd(ByRef hWnd As LongPtr) As String
  If hWnd = NULL_ Then
    GetClassFromHwnd = vbNullString
  Else
    Static LengthOfText As Long
    LengthOfText = GetClassName(hWnd, Title, Len(Title))
    GetClassFromHwnd = Left$(Title, LengthOfText)
  End If
End Function
'#############################################
Private Function GetExeFilenameFromHwnd(ByRef hWnd As LongPtr) As String
  If hWnd = NULL_ Then
    GetExeFilenameFromHwnd = vbNullString

  Else
    Dim pid      As Long
    Dim hProcess As LongPtr
    Dim exePath  As String * 512
  
    GetWindowThreadProcessId hWnd, pid
  
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pid)
    If hProcess <> 0 Then
      If GetModuleFileNameExA(hProcess, 0, exePath, 512) > 0 Then GetExeFilenameFromHwnd = Left(exePath, InStr(exePath, vbNullChar) - 1)
      CloseHandle hProcess
    End If

  End If
End Function
'#############################################


Public Function GetFocusedControl() As LongPtr
  Call AttachThreadToActiveWindow
  GetFocusedControl = GetFocus()
  Call RemoveAttachedThread
End Function

Public Function GetControlUndeMouse() As LongPtr
  Dim MousePos As POINTAPI
  GetCursorPos MousePos
  Call AttachThreadToWindowFromPoint(MousePos.x, MousePos.y)
  
  GetControlUndeMouse = WindowFromPoint(MousePos.x, MousePos.y)
  Call RemoveAttachedThread
End Function

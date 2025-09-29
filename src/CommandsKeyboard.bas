Attribute VB_Name = "CommandsKeyboard"
Option Explicit
Private Const MODULE_NAME As String = "CommandsKeyboard"

Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
'http://allapi.mentalis.org/apilist/keyb_event.shtml , http://www.vbforums.com/showthread.php?553261-RESOLVED-Send-SHIFT-mouse-click-by-code

Private Declare PtrSafe Function GetKeyboardLayoutNameA Lib "user32" (ByVal pwszKLID As String) As Long
Private Declare PtrSafe Function LoadKeyboardLayoutA Lib "user32" (ByVal pwszKLID As String, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As LongPtr, ByVal flags As Long) As LongPtr

Private Declare PtrSafe Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT_KEYBDINPUT, ByVal cbSize As Long) As Long
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long


Private Enum KEYEVENTF_Constants       'for keybd_event
  KEYEVENTF_EXTENDEDKEY = &H1&
  KEYEVENTF_KEYUP = &H2&
  KEYEVENTF_UNICODE = &H4&
  KEYEVENTF_SCANCODE = &H8&
End Enum

Private Enum WM_Constants              'for keybd_event
  WM_KEYFIRST = &H100&
  WM_KEYDOWN = &H100&
  WM_KEYUP = &H101&
  WM_CHAR = &H102& 'https://docs.microsoft.com/en-us/windows/win32/inputdev/using-keyboard-input
  WM_DEADCHAR = &H103&
  WM_SYSKEYDOWN = &H104&
  WM_SYSKEYUP = &H105&
  WM_SYSCHAR = &H106&
  WM_SYSDEADCHAR = &H107&
  WM_KEYLAST = &H108&
End Enum

Private Enum LOCALE_Constants 'for GetLocaleInfo
  LOCALE_ILANGUAGE = &H1&         'language id                   1009
  LOCALE_SLANGUAGE = &H2&         'localized name of language    Canada
  LOCALE_SENGLANGUAGE = &H1001&   'English name of language      Canada
  LOCALE_SABBREVLANGNAME = &H3&   'abbreviated language name     English
  LOCALE_SCOUNTRY = &H6&          'localized name of country     CA
  LOCALE_SENGCOUNTRY = &H1002&    'English name of country       en
  LOCALE_SABBREVCTRYNAME = &H7&   'abbreviated country name      ENC
  '#if(WINVER >=  &H0400)
  LOCALE_SISO639LANGNAME = &H59&  'ISO abbreviated language name
  LOCALE_SISO3166CTRYNAME = &H5A& 'ISO abbreviated country name
End Enum

'for GetKeyboardLayoutName, LoadKeyboardLayout
Private Const KL_NAMELENGTH As Long = 9&
Private KLID     As String * KL_NAMELENGTH
#If Win64 Then
  Private NewHKL As LongPtr, ThisHKL As LongPtr
#Else
  Private NewHKL As Long, ThisHKL As Long
#End If


Private Const INPUT_KEYBOARD = 1

Private Enum VK_Constants
  VK_BACK = &H8&          'Backspace
  VK_TAB = &H9&
  VK_RETURN = &HD&        'Enter
  VK_ENTER = &HD&         'Enter
  
  VK_SHIFT = &H10&        'Shift
  VK_CONTROL = &H11&      'Ctrl
  VK_MENU = &H12&         'Alt
  VK_CAPITAL = &H14&      'CapsLock
  VK_CAPSLOCK = &H14&     'CapsLock
  VK_ESC = &H1B&          'Escape
  
  vk_SPACE = &H20&
  VK_PRIOR = &H21&        'Page Up
  VK_NEXT = &H22&         'Page Down
  VK_END = &H23&
  VK_HOME = &H24&
  VK_LEFT = &H25&
  VK_UP = &H26&
  VK_RIGHT = &H27&
  VK_DOWN = &H28&
  VK_SELECT = &H29&
  VK_PRINT = &H2A&
  VK_EXECUTE = &H2B&
  VK_SNAPSHOT = &H2C&     'Print Screen
  VK_INSERT = &H2D&
  VK_DELETE = &H2E&
  VK_HELP = &H2F&
  
  vk_0 = &H30&            '0
  VK_9 = &H39&            '9
  VK_A = &H41&            'A
  vk_Z = &H5A&            'Z
  
  VK_LWIN = &H5B&
  VK_RWIN = &H5C&
  VK_APPS = &H5D&
  VK_SLEEP = &H5F&
  
  VK_STARTKEY = &H5B&
  VK_WINKEY = &H5B&
  VK_CONTEXTKEY = &H5D&
  
  vk_NUMPAD0 = &H60&      '0
  VK_NUMPAD9 = &H69&      '9
  VK_MULTIPLY = &H6A&     '*
  VK_ADD = &H6B&          '+
  VK_SEPARATOR = &H6C&    ',
  VK_SUBSTRACT = &H6D&    '-
  VK_DECIMAL = &H6E&      '.
  vk_DIVIDE = &H6F&       '/
  
  VK_NUMLOCK = &H90&      'info: CapsLock = &H14&
  VK_SCROLL = &H91&
  VK_SCROLLLOCK = &H91&
  
  VK_LSHIFT = &HA0&       'Shift
  VK_RSHIFT = &HA1&       'Shift
  VK_LCONTROL = &HA2&     'Ctrl
  VK_RCONTROL = &HA3&     'Ctrl
  VK_LMENU = &HA4&        'Alt
  VK_RMENU = &HA5&        'Alt
End Enum

#If Win64 Then
  Private Type GENERALINPUT_KEYBDINPUT
    dwType As Long
    dummy As Long
    'KEYBDINPUT:
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    dummy3 As Long
    time As Long
    dummy4 As Long
    dwExtraInfo As Long
    dummy1 As Long
    dummy2 As Long
  End Type
#ElseIf VBA7 Then
  Private Type GENERALINPUT_KEYBDINPUT
    dwType As Long
    'KEYBDINPUT:
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
    dummy1 As Long
    dummy2 As Long
  End Type
#Else
  Private Type GENERALINPUT_KEYBDINPUT
    dwType As Long
    'KEYBDINPUT:
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
  End Type
#End If


Private Enum MAPVK_Constants 'for MapVirtualKey
  MAPVK_VK_TO_VSC = 0&    'The uCode parameter is a virtual-key code and is translated into a scan code. If it is a virtual-key code that does not distinguish between left- and right-hand keys, the left-hand scan code is returned. If there is no translation, the function returns 0.
  MAPVK_VSC_TO_VK = 1&    'The uCode parameter is a scan code and is translated into a virtual-key code that does not distinguish between left- and right-hand keys. If there is no translation, the function returns 0.
  MAPVK_VK_TO_CHAR = 2&   'The uCode parameter is a virtual-key code and is translated into an unshifted character value in the low order word of the return value. Dead keys (diacritics) are indicated by setting the top bit of the return value. If there is no translation, the function returns 0.
  MAPVK_VSC_TO_VK_EX = 3& 'The uCode parameter is a scan code and is translated into a virtual-key code that distinguishes between left- and right-hand keys. If there is no translation, the function returns 0.
End Enum

Private GInput(0 To 3) As GENERALINPUT_KEYBDINPUT
'Private tmpR As Range, tmpB As Byte, tmpI As Integer
Private lastCharCode As Byte


Private WshShell      As Object
Private tmpL          As Long
Private tmpS          As String
Private tmpR          As Range

Private Codes(0 To 9) As Byte


Public Sub RegisterCommandsKeyboard()
  On Error GoTo eh
' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "sendkeystoactivewindow", Array("SendKeysToActiveWindow", "Send Keys To Active Window", _
    MODULE_NAME, "Simulate the writing of the text from the arguments to the active window", _
    "Text", "Text to write to the active window", _
    "Text2", "Text to write to the active window", _
    "Text3", "Text to write to the active window (and so on)")
  
  commandMap.Add "write", Array("SendKeysToActiveWindow", "Write", _
    MODULE_NAME, "Simulate the writing of the text from the arguments to the active window", _
    "Text", "Text to write to the active window", _
    "Text2", "Text to write to the active window", _
    "Text3", "Text to write to the active window (and so on)")


  commandMap.Add "keypress", Array("KeyPress", "Key Press", _
    MODULE_NAME, "Simulate the pressing of the key specified in the arguments. After all keys were pressed, simulate the eliberation of all the keys pressed.", _
    "Key", "Key to be pressed", _
    "Key2", "Key to be pressed", _
    "Key3", "Key to be pressed (and so on)")
  
  commandMap.Add "keydown", Array("KeyDown", "Key Down", _
    MODULE_NAME, "Simulate the pressing of the key specified in the arguments.", _
    "Key", "Key to be pressed", _
    "Key2", "Key to be pressed", _
    "Key3", "Key to be pressed (and so on)")
  
  commandMap.Add "keyup", Array("KeyUp", "Key Up", _
    MODULE_NAME, "Simulate the eliberation of the key specified in the arguments.", _
    "Key", "Key to be depressed", _
    "Key2", "Key to be depressed", _
    "Key3", "Key to be depressed (and so on)")



  commandMap.Add "getkeyblayout", Array("GetKeybLayout", "Get Keyb Layout", _
    MODULE_NAME, "Retrieves the name of the active input locale identifier (formerly called the keyboard layout) for the system. The input locale identifier is a broader concept than a keyboard layout, since it can also encompass a speech-to-text converter, an Input Method Editor (IME), or any other form of input.", _
    "Identifier", "Here will be written the keyboard layout identifier")

  commandMap.Add "getkeyboardlayout", Array("GetKeybLayout", "Get Keyboard Layout", _
    MODULE_NAME, "Retrieves the name of the active input locale identifier (formerly called the keyboard layout) for the system. The input locale identifier is a broader concept than a keyboard layout, since it can also encompass a speech-to-text converter, an Input Method Editor (IME), or any other form of input.", _
    "Identifier", "Here will be written the keyboard layout identifier")

  commandMap.Add "setkeyblayout", Array("SetKeybLayout", "Set Keyb Layout", _
    MODULE_NAME, "Sets the input locale identifier (formerly called the keyboard layout handle) for the calling thread or the current process. The input locale identifier specifies a *locale* as well as the *physical* layout of the keyboard.", _
    "Identifier", "Specify the keyboard layout identifier")

  commandMap.Add "setkeyboardlayout", Array("SetKeybLayout", "Set Keyboard Layout", _
    MODULE_NAME, "Sets the input locale identifier (formerly called the keyboard layout handle) for the calling thread or the current process. The input locale identifier specifies a *locale* as well as the *physical* layout of the keyboard.", _
    "Identifier", "Specify the keyboard layout identifier")



  commandMap.Add "chardown", Array("CharDown", "Char Down", _
    MODULE_NAME, "Simulate the press of the key specified in the arguments. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "charup", Array("CharUp", "Char Up", _
    MODULE_NAME, "Simulate the release of the key specified in the arguments. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "charpress", Array("CharPress", "Char Press", _
    MODULE_NAME, "Simulate the press of the key specified in the arguments, after that the key will be released. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")



  commandMap.Add "altchardown", Array("AltCharDown", "Alt Char Down", _
    MODULE_NAME, "Simulate the press of the Alt key and right after the key specified in the arguments. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "altcharup", Array("AltCharUp", "Alt Char Up", _
    MODULE_NAME, "Simulate the release of the key specified in the arguments and right after the release of the Alt key. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "altcharpress", Array("AltCharPress", "Alt Char Press", _
    MODULE_NAME, "Simulate the press of the Alt key and right after the key specified in the arguments will be pressed and released, after that the Alt key will be released. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")



  commandMap.Add "ctrlchardown", Array("CtrlCharDown", "Ctrl Char Down", _
    MODULE_NAME, "Simulate the press of the Ctrl key and right after the key specified in the arguments. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "ctrlcharup", Array("CtrlCharUp", "Ctrl Char Up", _
    MODULE_NAME, "Simulate the release of the key specified in the arguments and right after the release of the Ctrl key. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")

  commandMap.Add "ctrlcharpress", Array("CtrlCharPress", "Ctrl Char Press", _
    MODULE_NAME, "Simulate the press of the Ctrl key and right after the key specified in the arguments will be pressed and released, after that the Ctrl key will be released. Only first letter of each argument is considered, only 0...9 and a...z and Space are considered.", _
    "Char", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char2", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "Char3", "Only the first letter is considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")


  commandMap.Add "stringpress", Array("StringPress", "String Press", _
    MODULE_NAME, "Simulate the press of the keys specified in the arguments. All characters of each argument are considered, only 0...9 and a...z and Space are considered.", _
    "String", "All characters are considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "String2", "All characters are considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key.", _
    "String3", "All characters are considered, only 0...9 and a...z and Space is considered. Else is considered as code name of the key (and so on)")


  commandMap.Add "altdown", Array("AltDown", "Alt Down", _
    MODULE_NAME, "Simulate the press of the Alt key.")

  commandMap.Add "altup", Array("AltUp", "Alt Up", _
    MODULE_NAME, "Simulate the release of the Alt key.")

  commandMap.Add "altpress", Array("AltPress", "Alt Press", _
    MODULE_NAME, "Simulate the press of the Alt key.")


  commandMap.Add "ctrldown", Array("CtrlDown", "Ctrl Down", _
    MODULE_NAME, "Simulate the press of the Ctrl key.")

  commandMap.Add "ctrlup", Array("CtrlUp", "Ctrl Up", _
    MODULE_NAME, "Simulate the release of the Ctrl key.")

  commandMap.Add "ctrlpress", Array("CtrlPress", "Ctrl Press", _
    MODULE_NAME, "Simulate the press of the Ctrl key.")


  commandMap.Add "shiftdown", Array("ShiftDown", "Shift Down", _
    MODULE_NAME, "Simulate the press of the Shift key.")

  commandMap.Add "shiftup", Array("ShiftUp", "Shift Up", _
    MODULE_NAME, "Simulate the release of the Shift key.")

  commandMap.Add "shiftpress", Array("ShiftPress", "Shift Press", _
    MODULE_NAME, "Simulate the press of the Shift key.")

  Set WshShell = CreateObject("WScript.Shell")
  
  For tmpL = LBound(GInput) To UBound(GInput)
    GInput(tmpL).dwType = INPUT_KEYBOARD
  Next

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsKeyboard", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsKeyboard()
  On Error GoTo eh

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsKeyboard", Err.Number, Err.Source, Err.description, Erl
End Sub


Public Function SendKeysToActiveWindow(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  For tmpL = 0 To 9
    If Len(currentRowArray(1, ColAArg1 + tmpL)) <> 0 Then
      Call WshShell.SendKeys(currentRowArray(1, ColAArg1 + tmpL), True)
      Call Sleep(waitAfterKeyPress)
    End If
  Next

done:
  SendKeysToActiveWindow = True
  Exit Function
eh:
  SendKeysToActiveWindow = False
  RaiseError MODULE_NAME & ".SendKeysToActiveWindow", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


Public Function KeyPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadCodes
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then keybd_event Codes(tmpL), 0, 0, 0
  Next
  Sleep waitAfterKeyPress
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then keybd_event Codes(tmpL), 0, KEYEVENTF_KEYUP, 0
  Next

done:
  KeyPress = True
  Exit Function
eh:
  KeyPress = False
  RaiseError MODULE_NAME & ".KeyPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function KeyDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadCodes
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then keybd_event Codes(tmpL), 0, 0, 0
  Next

done:
  KeyDown = True
  Exit Function
eh:
  KeyDown = False
  RaiseError MODULE_NAME & ".KeyDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function KeyUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadCodes
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then keybd_event Codes(tmpL), 0, KEYEVENTF_KEYUP, 0
  Next

done:
  KeyUp = True
  Exit Function
eh:
  KeyUp = False
  RaiseError MODULE_NAME & ".KeyUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function




'############### GetCode ###############
Private Function GetCode(s As String) As Byte
  On Error GoTo eh
  If Len(s) = 0 Then GoTo Invalid
  s = Replace(s, "*", "MULTIPLY")
  s = Replace(s, ".", "DECIMAL")
  s = Replace(s, ",", "SEPARATOR")
  Set tmpR = shKey.Columns(ColKeyName).Find(UCase(s), LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
  If tmpR Is Nothing Then GoTo Invalid
  GetCode = CByte(tmpR.Offset(0, ColKeyCodeDec - ColKeyName).Value)
  Exit Function
Invalid:
  GetCode = 0&

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".GetCode", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function ReadCodes() As Boolean
  On Error GoTo eh
  For tmpL = 0 To 9
    Codes(tmpL) = GetCode(CStr(currentRowArray(1, ColAArg1 + tmpL)))
  Next

done:
  ReadCodes = True
  Exit Function
eh:
  ReadCodes = False
  RaiseError MODULE_NAME & ".ReadCodes", Err.Number, Err.Source, Err.description, Erl
End Function



Public Function GetKeybLayoutAsString() As String
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeyboardlayoutnamea
' Retrieves the name of the active input locale identifier (formerly called the keyboard layout) for the system.
' The input locale identifier is a broader concept than a keyboard layout,
' since it can also encompass a speech-to-text converter, an Input Method Editor (IME),
' or any other form of input.
  On Error GoTo eh
  Call GetKeyboardLayoutNameA(KLID)
  GetKeybLayoutAsString = KLID

done:
  Exit Function
eh:
  GetKeybLayoutAsString = vbNullString
  RaiseError MODULE_NAME & ".GetKeybLayoutAsString", Err.Number, Err.Source, Err.description, Erl
End Function
Public Function GetKeybLayout(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getkeyboardlayoutnamea
' Retrieves the name of the active input locale identifier (formerly called the keyboard layout) for the system.
' The input locale identifier is a broader concept than a keyboard layout,
' since it can also encompass a speech-to-text converter, an Input Method Editor (IME),
' or any other form of input.
' Arg1=Identifier: Here will be written the keyboard layout identifier
  On Error GoTo eh
  Call GetKeyboardLayoutNameA(KLID)
  currentRowRange(1, ColAArg1).Value = KLID

done:
  GetKeybLayout = True
  Exit Function
eh:
  GetKeybLayout = False
  RaiseError MODULE_NAME & ".PrepareExitCommandsFile", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function SetKeybLayout(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
' https://docs.microsoft.com/de-de/windows/win32/api/winuser/nf-winuser-activatekeyboardlayout
' Sets the input locale identifier (formerly called the keyboard layout handle)
' for the calling thread or the current process.
' The input locale identifier specifies a *locale* as well as the *physical* layout of the keyboard.
' Arg1=Identifier: provide the keyboard layout identifier
  On Error GoTo eh
  NewHKL = LoadKeyboardLayout(currentRowArray(1, ColAArg1))
  If NewHKL = 0 Then
    SetKeybLayout = False
    RaiseError MODULE_NAME & ".SetKeybLayout", Err.Number, Err.Source, "Keyboard layout not supported: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If
  ThisHKL = ActivateKeyboardLayout(NewHKL, 0)
  If NewHKL = ThisHKL Then
    SetKeybLayout = False
    RaiseError MODULE_NAME & ".SetKeybLayout", Err.Number, Err.Source, "Unable to activate keyboard layout: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "]", Erl, 2, ExecutingTroughApplicationRun
    Exit Function
  End If

done:
  SetKeybLayout = True
  Exit Function
eh:
  SetKeybLayout = False
  RaiseError MODULE_NAME & ".SetKeybLayout", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

#If Win64 Then
Private Function LoadKeyboardLayout(ByVal LCID As String) As LongPtr
#Else
Private Function LoadKeyboardLayout(ByVal LCID As String) As Long
#End If
' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-loadkeyboardlayouta
' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/0c8f80b2-5084-4dba-b272-6797620f84a3/automatically-change-keyboard-input-language-when-set-focus-to-textbox-in-excel-userform?forum=exceldev
' Loads a new input locale identifier (formerly called the keyboard layout) into the system.
' The layout is not activated, use ActivateKeyboardLayout

  On Error GoTo eh
  KLID = Right(String(KL_NAMELENGTH - 1, "0") & LCID, KL_NAMELENGTH - 1) & vbNullChar
  LoadKeyboardLayout = LoadKeyboardLayoutA(KLID, 0)

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".LoadKeyboardLayout", Err.Number, Err.Source, Err.description, Erl
End Function


' interesting implemented Numlock, ScrollLock, Capslock: http://access.mvps.org/access/api/api0046.htm

' https://www.codeproject.com/Articles/7305/Keyboard-Events-Simulation-using-keybd-event-funct
' bVk //Virtual Keycode of keys. E.g., VK_RETURN, VK_TAB…
' bScan //Scan Code value of keys. E.g., 0xb8 for “Left Alt” key.
' dwFlags //Flag that is set for key state. E.g., KEYEVENTF_KEYUP.
' dwExtraInfo //32-bit extra information about keystroke.

' use of keybd_event: <keybd_event VK_LSHIFT, 0, 0, 0> for Lshift pressed; <keybd_event VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0> for Lshift depressed
' http://pinvoke.net/default.aspx/user32.keybd_event





'############### Char ###############
Public Function CharDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_KeyDown Codes(tmpL)
  Next

done:
  CharDown = True
  Exit Function
eh:
  CharDown = False
  RaiseError MODULE_NAME & ".CharDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CharUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_KeyUp Codes(tmpL)
  Next

done:
  CharUp = True
  Exit Function
eh:
  CharUp = False
  RaiseError MODULE_NAME & ".CharUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CharPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_KeyPress Codes(tmpL)
  Next

done:
  CharPress = True
  Exit Function
eh:
  CharPress = False
  RaiseError MODULE_NAME & ".CharPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'############### Alt+Char ###############
Public Function AltCharDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_AltKeyDown Codes(tmpL)
  Next

done:
  AltCharDown = True
  Exit Function
eh:
  AltCharDown = False
  RaiseError MODULE_NAME & ".AltCharDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function AltCharUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_AltKeyUp Codes(tmpL)
  Next

done:
  AltCharUp = True
  Exit Function
eh:
  AltCharUp = False
  RaiseError MODULE_NAME & ".AltCharUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function AltCharPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_AltKeyPress Codes(tmpL)
  Next

done:
  AltCharPress = True
  Exit Function
eh:
  AltCharPress = False
  RaiseError MODULE_NAME & ".AltCharPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'############### Control+Char ###############
Public Function CtrlCharDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_CtrlKeyDown Codes(tmpL)
  Next

done:
  CtrlCharDown = True
  Exit Function
eh:
  CtrlCharDown = False
  RaiseError MODULE_NAME & ".CtrlCharDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CtrlCharUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_CtrlKeyUp Codes(tmpL)
  Next

done:
  CtrlCharUp = True
  Exit Function
eh:
  CtrlCharUp = False
  RaiseError MODULE_NAME & ".CtrlCharUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CtrlCharPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Call ReadChars
  For tmpL = 0 To 9
    If Codes(tmpL) > 0 Then SendInput_CtrlKeyPress Codes(tmpL)
  Next

done:
  CtrlCharPress = True
  Exit Function
eh:
  CtrlCharPress = False
  RaiseError MODULE_NAME & ".CtrlCharPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


'############### GetChar ###############
Private Function GetChar(ch As String) As Byte
  On Error GoTo eh
  If Len(ch) = 0 Then GoTo Invalid
  GetChar = Asc(UCase(Left(ch, 1)))
  If (GetChar > vk_0 And GetChar < vk_Z) Or GetChar = vk_SPACE Then 'Or (GetChar > vk_NUMPAD0 And GetChar < vk_DIVIDE)
    GetChar = GetCode(Left(ch, 1))
  End If
  Exit Function
Invalid:
  GetChar = 0&

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsFile", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function ReadChars() As Boolean
  On Error GoTo eh
  For tmpL = 0 To 9
    Codes(tmpL) = GetChar(CStr(currentRowArray(1, ColAArg1 + tmpL)))
  Next

done:
  ReadChars = True
  Exit Function
eh:
  ReadChars = False
  RaiseError MODULE_NAME & ".ReadChars", Err.Number, Err.Source, Err.description, Erl
End Function
'
'Public Function IsChar(ch As String) As Boolean
'  If Len(ch) = 0 Then GoTo Invalid
'  lastCharCode = Asc(UCase(Left(ch, 1)))
'  IsChar = (lastCharCode > vk_0 And lastCharCode < vk_Z) Or lastCharCode = vk_SPACE 'Or (lastCharCode > vk_NUMPAD0 And lastCharCode < vk_DIVIDE)
'  If Not IsChar Then
'    lastCharCode = GetCode(Left(ch, 1))
'    If lastCharCode > 0 Then IsChar = True
'  End If
'  Exit Function
'Invalid:
'  IsChar = False
'End Function

'############### String ###############
Public Function StringPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Dim i As Long
  For i = 0 To 9
    tmpS = currentRowArray(1, ColAArg1 + tmpL)
    If Len(tmpS) > 0 Then
      For tmpL = 1 To Len(tmpS)
        SendInput_KeyPress GetChar(Mid(tmpS, tmpL, 1))
      Next
    End If
  Next

done:
  StringPress = True
  Exit Function
eh:
  StringPress = False
  RaiseError MODULE_NAME & ".StringPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function





'############### Key ###############
Private Function SendInput_KeyDown(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 1, GInput(0), Len(GInput(0))

done:
  SendInput_KeyDown = True
  Exit Function
eh:
  SendInput_KeyDown = False
  RaiseError MODULE_NAME & ".SendInput_KeyDown", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_KeyUp(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = bKey
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 1, GInput(0), Len(GInput(0))

done:
  SendInput_KeyUp = True
  Exit Function
eh:
  SendInput_KeyUp = False
  RaiseError MODULE_NAME & ".SendInput_KeyUp", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_KeyPress(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 2, GInput(0), Len(GInput(0))

done:
  SendInput_KeyPress = True
  Exit Function
eh:
  SendInput_KeyPress = False
  RaiseError MODULE_NAME & ".SendInput_KeyPress", Err.Number, Err.Source, Err.description, Erl
End Function

'############### Alt+Key ###############
Private Function SendInput_AltKeyDown(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_MENU
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 2, GInput(0), Len(GInput(0))

done:
  SendInput_AltKeyDown = True
  Exit Function
eh:
  SendInput_AltKeyDown = False
  RaiseError MODULE_NAME & ".SendInput_AltKeyDown", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_AltKeyUp(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_MENU
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
With GInput(1)
    .wVk = bKey
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 2, GInput(0), Len(GInput(0))

done:
  SendInput_AltKeyUp = True
  Exit Function
eh:
  SendInput_AltKeyUp = False
  RaiseError MODULE_NAME & ".SendInput_AltKeyUp", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_AltKeyPress(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_MENU
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(2)
    .wVk = GInput(1).wVk
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(3)
    .wVk = GInput(0).wVk
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 4, GInput(0), Len(GInput(0))

done:
  SendInput_AltKeyPress = True
  Exit Function
eh:
  SendInput_AltKeyPress = False
  RaiseError MODULE_NAME & ".SendInput_AltKeyPress", Err.Number, Err.Source, Err.description, Erl
End Function

'############### Control+Key ###############
Private Function SendInput_CtrlKeyDown(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_CONTROL
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 2, GInput(0), Len(GInput(0))

done:
  SendInput_CtrlKeyDown = True
  Exit Function
eh:
  SendInput_CtrlKeyDown = False
  RaiseError MODULE_NAME & ".SendInput_CtrlKeyDown", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_CtrlKeyUp(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_CONTROL
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 2, GInput(0), Len(GInput(0))

done:
  SendInput_CtrlKeyUp = True
  Exit Function
eh:
  SendInput_CtrlKeyUp = False
  RaiseError MODULE_NAME & ".SendInput_CtrlKeyUp", Err.Number, Err.Source, Err.description, Erl
End Function
Private Function SendInput_CtrlKeyPress(bKey As Byte) As Boolean
  On Error GoTo eh
  With GInput(0)
    .wVk = VK_CONTROL
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(1)
    .wVk = bKey
    .dwFlags = 0
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(2)
    .wVk = GInput(1).wVk
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  With GInput(3)
    .wVk = GInput(0).wVk
    .dwFlags = KEYEVENTF_KEYUP
    .wScan = MapVirtualKey(.wVk, MAPVK_VK_TO_VSC)
  End With
  SendInput 4, GInput(0), Len(GInput(0))

done:
  SendInput_CtrlKeyPress = True
  Exit Function
eh:
  SendInput_CtrlKeyPress = False
  RaiseError MODULE_NAME & ".SendInput_CtrlKeyPress", Err.Number, Err.Source, Err.description, Erl
End Function



'############### Alt ###############
Public Function AltDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  AltDown = SendInput_KeyDown(VK_MENU)

done:
  Exit Function
eh:
  AltDown = False
  RaiseError MODULE_NAME & ".AltDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
  End Function
Public Function AltUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  AltUp = SendInput_KeyUp(VK_MENU)

done:
  Exit Function
eh:
  AltUp = False
  RaiseError MODULE_NAME & ".AltUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
  End Function
Public Function AltPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  AltPress = SendInput_KeyPress(VK_MENU)

done:
  Exit Function
eh:
  AltPress = False
  RaiseError MODULE_NAME & ".AltPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'############### Control ###############
Public Function CtrlDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  CtrlDown = SendInput_KeyDown(VK_CONTROL)

done:
  Exit Function
eh:
  CtrlDown = False
  RaiseError MODULE_NAME & ".CtrlDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CtrlUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  CtrlUp = SendInput_KeyUp(VK_CONTROL)

done:
  Exit Function
eh:
  CtrlUp = False
  RaiseError MODULE_NAME & ".CtrlUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function CtrlPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  CtrlPress = SendInput_KeyPress(VK_CONTROL)

done:
  Exit Function
eh:
  CtrlPress = False
  RaiseError MODULE_NAME & ".CtrlPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'############### Shift ###############
Public Function ShiftDown(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  ShiftDown = SendInput_KeyDown(VK_SHIFT)

done:
  Exit Function
eh:
  ShiftDown = False
  RaiseError MODULE_NAME & ".ShiftDown", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function ShiftUp(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  ShiftUp = SendInput_KeyUp(VK_SHIFT)

done:
  Exit Function
eh:
  ShiftUp = False
  RaiseError MODULE_NAME & ".ShiftUp", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function
Public Function ShiftPress(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  ShiftPress = SendInput_KeyPress(VK_SHIFT)

done:
  Exit Function
eh:
  ShiftPress = False
  RaiseError MODULE_NAME & ".ShiftPress", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function


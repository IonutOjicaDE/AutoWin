Attribute VB_Name = "CommandsDisplay"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsDisplay"
'https://bytes.com/topic/visual-basic/answers/952015-changing-screen-resolution-excel-vba
'https://www.vbarchiv.net/api/api_changedisplaysettings.html

Private tmpHwnd As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As DEVMODE, ByVal DwFlag As Long) As Long
Private Declare PtrSafe Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As LongPtr
'https://www.vbarchiv.net/api/api_enumdisplaydevices.html
Private Declare PtrSafe Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (DeviceName As Any, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Long



Private Const CCFORMNAME = 32&
Private Const CCDEVICENAME = 32&

Private Type DEVMODE
  dmDeviceName      As String * CCDEVICENAME
  dmSpecVersion     As Integer
  dmDriverVersion   As Integer
  dmSize            As Integer
  dmDriverExtra     As Integer
  
  dmFields          As Long 'see DM_Constants
  dmOrientation     As Integer
  dmPaperSize       As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long 'properties of the graphics card - see DM_Constants
  dmDisplayFrequency As Long
End Type

'======================================================
'This code shows how to change the screen resolution.
'Call the function like this:
' ChangeResolution 640, 480
'This would change the screen resolution to 640 pixels x 480 pixels. Note that
'you can only change the resolution to values supported by the display.

Private Enum DM_Constants 'for DEVMODE dmFields
  DM_BITSPERPEL = &H40000      'The structure should be filled with the color depth
  DM_PELSWIDTH = &H80000       'the structure should contain the width of the screen in pixels
  DM_PELSHEIGHT = &H100000     'the structure should contain the height of the screen in pixels
  DM_DISPLAYFREQUENCY = &H400000 'the structure should contain the RefreshFrequency in Herz
  DM_DISPLAYFLAGS = &H200000   'The structure should contain the properties of the graphics card
  
  'for DEVMODE dmDisplayFlags
  DM_GRAYSCALE = 1&            'Device does not support colors, gray tones are supported
  DM_INTERLACED = 2&           'Device supports colors
End Enum

Private Enum ENUM_Constants 'for EnumDisplaySettings iModeNum
  ENUM_CURRENT_SETTINGS = -1&  'The function should fill the structure with the current settings
  ENUM_REGISTRY_SETTINGS = -2& 'The function should fill the structure with the registry settings
End Enum

Private Enum CDS_Constants 'for ChangeDisplaySettings dwFlags
  CDS_UPDATEREGISTRY = &H1     'The settings are saved in the Registry
  CDS_TEST = &H2               'Tests the resolution without changing the resolution, the return is one of the return constants (see Enum DISP_CHANGE_Constants)
  CDS_FULLSCREEN = &H4         'The graphics mode should be displayed in full screen, this setting cannot be saved
  CDS_GLOBAL = &H8             'Saves the settings for all Users (in connection with CDS_UPDATEREGISTRY)
  CDS_SET_PRIMARY = &H10       'The specified graphics card will be the standard graphics card
  CDS_RESET = &H40000000       'Changes the resolution, even if it is the same as that currently displayed
  CDS_NORESET = &H10000000     'Saves the settings in the Registry, the changes will only take effect after a restart (in Connection with CDS_UPDATEREGISTRY)
End Enum

Private Enum DISP_CHANGE_Constants 'ChangeDisplaySettings return constants
  DISP_CHANGE_SUCCESSFUL = 0&  'Changing or testing the the resolution was successful
  DISP_CHANGE_RESTART = 1&     'Changing the resolution requires a restart
  DISP_CHANGE_FAILED = -1&     'Changing or testing the resolution failed
  DISP_CHANGE_BADMODE = -2&    'The specified resolution is not supported
  DISP_CHANGE_NOTUPDATED = -3& '(Win NT / 2000) The settings were not saved
  DISP_CHANGE_BADFLAGS = -4&   'Wrong flags were specified
  DISP_CHANGE_BADPARAM = -5&   'Wrong parameters were specified
End Enum

Private Type DISPLAY_DEVICE
  cb           As Long 'Erwartet die Gr��e der Struktur in Bytes
  DeviceName   As String * 32  'Erh�lt den String, der die Grafikkarte beschreibt in dem Format "\\.\DISPLAYX", wobei das X durch den Index des Ger�tes ersetzt wird. Der String ist mit einem "VBNullChar"-Zeichenterminiert
  DeviceString As String * 128 'Erh�lt den Namen der Grafikkarte mit abschlie�endem "VBNullChar"-Zeichen
  StateFlags   As Long 'Erh�lt den Typ der Grafikkarte, der mittels einer oder mehreren StateFlags-Konstanten ausgewertet werden kann
  DeviceID     As String * 128 '(Win 98/ME) Erh�lt einen Plug & PlayIdentifier, der die Grafikkarte oder den Monitor in Form eines Strings beschreibt
  DeviceKey    As String * 128 'Reserviert, diese Option wird nicht genutzt
End Type

Private Enum DISPLAY_DEVICE_Constants 'EnumDisplayDevices StateFlags constants
  DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = &H1 ' Das Ger�t ist Teil des Desktops
  DISPLAY_DEVICE_MIRRORING_DRIVER = &H8    ' Dieses Ger�t ist ein unsichtbarer Pseudo-Monitor
  DISPLAY_DEVICE_MODESPRUNED = &H8000000   ' Dieses Ger�t hat mehr Grafikmodes, als das Ausgabeger�t unterst�tzt
  DISPLAY_DEVICE_PRIMARY_DEVICE = &H4      ' Das Ger�t ist die Standardgrafikkarte
  DISPLAY_DEVICE_VGA_COMPATIBLE = &H10     ' Das Ger�t ist VGA-kompatibel
End Enum

Private rc               As RECT
Private tmpL             As Long
Private tmpS             As String


Public Sub RegisterCommandsDisplay()
  On Error GoTo eh
  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "changeresolution", Array("ChangeResolution", "Change Resolution", _
    MODULE_NAME, "Change the resolution to a supported display resolution (640x480 , 800x600 , 1024x768 ...).", _
    "Width", "Number of pixels horizontally.", _
    "Height", "Number of pixels vertically.")
  commandMap.Add "getresolution", Array("GetResolution", "Get Resolution", _
    MODULE_NAME, "Retreive the current resolution of the main display.", _
    "Width", "Number of pixels horizontally.", _
    "Height", "Number of pixels vertically.")

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsDisplay", Err.Number, Err.Source, Err.description, Erl
End Sub
Public Sub PrepareExitCommandsDisplay()
  On Error GoTo eh

done:
  Exit Sub
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsDisplay", Err.Number, Err.Source, Err.description, Erl
End Sub


Public Function ChangeResolution(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If Not (IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 1)))) Then
    ChangeResolution = False
    RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
      "Arguments need to be valid numbers: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If
  
  Dim DevM As DEVMODE
  Dim a As Boolean
  
  tmpL = 0
  
  Do 'Enumerate settings
    a = EnumDisplaySettings(lpszDeviceName:=0&, iModeNum:=tmpL&, lpDevMode:=DevM)
    tmpL = tmpL + 1
  Loop Until (a = False)
  
  DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Change settings
  
  DevM.dmPelsWidth = currentRowArray(1, ColAArg1 + 0)
  DevM.dmPelsHeight = currentRowArray(1, ColAArg1 + 1)
  
  If ChangeDisplaySettings(DevM, CDS_TEST) <= 0 Then ChangeDisplaySettings DevM, CDS_UPDATEREGISTRY

done:
  ChangeResolution = True
  Exit Function
eh:
  ChangeResolution = False
  RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

'*****************************************************************
' RETURN:
' The current screen resolution. Typically one of the following:
' 640 x 480
' 800 x 600
' 1024 x 768
'*****************************************************************
Public Function GetResolution(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  Dim DevM As DEVMODE
  tmpL = 0
  'while enumdisplaydevices(

  GetWindowRect GetDesktopWindow(), rc
  With rc
    currentRowRange(1, ColAArg1 + 0).Value = .Right - .Left
    currentRowRange(1, ColAArg1 + 1).Value = .Bottom - .Top
  End With
done:
  GetResolution = True
  Exit Function
eh:
  GetResolution = False
  RaiseError MODULE_NAME & "GetResolution", Err.Number, Err.Source, Err.description, Erl, , ExecutingTroughApplicationRun
End Function

' Alle m�glichen Aufl�sungen aller Grafikkarten ermitteln (Windows 98, ME, NT, 2000)
Private Sub Command2_Click()
  Dim retval As Long, Dev As DEVMODE, i As Long, j As Long, Disp As DISPLAY_DEVICE
  Dim TmpDevName As String ' Dient dazu den DeviceNamen ohne
  ' VBNullChar-Zeichen zwischenzuspeichern
  Dim TmpDevString As String ' Dient dazu den DeviceString ohne
  ' VBNullChar-Zeichen zwischenzuspeichern
 
  ' Falls Windows 95 l�uft
  On Error GoTo ErrWin95
 
  ' DEVMODE-Struktur vorinitialisieren
  Dev.dmSize = Len(Dev)
  Dev.dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT Or _
  DM_DISPLAYFREQUENCY Or DM_DISPLAYFLAGS
 
  ' DISPLAY_DEVICE-Struktur vorinitialisieren
  Disp.cb = Len(Disp)
 
  ' Alle Grafikkarten ermitteln
  Do While EnumDisplayDevices(ByVal 0&, i, Disp, 0&) <> 0
 
    ' VBNullChar-Zeichen abtrennen
    TmpDevName = Left$(Disp.DeviceName, InStr(1, Disp.DeviceName, _
    vbNullChar) - 1)
    TmpDevString = Left$(Disp.DeviceString, InStr(1, _
    Disp.DeviceString, vbNullChar) - 1)
 
    ' Informationen zu der gefundenen Grafikkarte ausgeben
    If CBool(DISPLAY_DEVICE_PRIMARY_DEVICE And Disp.StateFlags) Then
       Debug.Print "Grafikkarte (Standard): " & TmpDevString
    ElseIf CBool(DISPLAY_DEVICE_ATTACHED_TO_DESKTOP And _
    Disp.StateFlags) Then
       Debug.Print "Grafikkarte (Desktop): " & TmpDevString
    Else
      Debug.Print "Grafikkarte: " & TmpDevString
    End If
    i = i + 1
 
    ' F�r jede gefundene Grafikkarte die Bildschirmaufl�sungen ermitteln
    Do While EnumDisplaySettings(TmpDevName, j, Dev) <> 0
      Debug.Print "Bildschrimaufl�sung (" & CStr(j) & "): " & _
      Dev.dmPelsWidth & "x" & Dev.dmPelsHeight & "x" & _
      Dev.dmBitsPerPel & " Freq: " & Dev.dmDisplayFrequency
      DoEvents
      j = j + 1
    Loop
    Debug.Print vbCrLf & "--------------------------------------"
  Loop
 
  Exit Sub
ErrWin95:
  If Err.Number = 453 Then
    MsgBox "Dieses Beispiel l�uft nur ab Windows 98"
  End If
End Sub




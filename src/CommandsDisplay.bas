Attribute VB_Name = "CommandsDisplay"
Option Explicit
Private Const MODULE_NAME   As String = "CommandsDisplay"
'https://bytes.com/topic/visual-basic/answers/952015-changing-screen-resolution-excel-vba
'https://www.vbarchiv.net/api/api_changedisplaysettings.html

Private tmpHwnd As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr

Private Declare PtrSafe Function ChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" (ByVal lpszDeviceName As String, lpDevMode As DEVMODE, ByVal hwnd As LongPtr, ByVal dwflags As Long, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As DEVMODE, ByVal DwFlag As Long) As Long
Private Declare PtrSafe Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As LongPtr
'https://www.vbarchiv.net/api/api_enumdisplaydevices.html
Private Declare PtrSafe Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (DeviceName As Any, ByVal iDevNum As Long, lpDisplayDevice As DISPLAY_DEVICE, ByVal dwflags As Long) As Long


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
  cb           As Long 'Erwartet die Größe der Struktur in Bytes
  DeviceName   As String * 32  'Erhält den String, der die Grafikkarte beschreibt in dem Format "\\.\DISPLAYX", wobei das X durch den Index des Gerätes ersetzt wird. Der String ist mit einem "VBNullChar"-Zeichenterminiert
  DeviceString As String * 128 'Erhält den Namen der Grafikkarte mit abschließendem "VBNullChar"-Zeichen
  StateFlags   As Long 'Erhält den Typ der Grafikkarte, der mittels einer oder mehreren StateFlags-Konstanten ausgewertet werden kann
  DeviceID     As String * 128 '(Win 98/ME) Erhält einen Plug & PlayIdentifier, der die Grafikkarte oder den Monitor in Form eines Strings beschreibt
  DeviceKey    As String * 128 'Reserviert, diese Option wird nicht genutzt
End Type

Private Enum DISPLAY_DEVICE_Constants 'EnumDisplayDevices StateFlags constants
  DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = &H1 ' Das Gerät ist Teil des Desktops
  DISPLAY_DEVICE_MIRRORING_DRIVER = &H8    ' Dieses Gerät ist ein unsichtbarer Pseudo-Monitor
  DISPLAY_DEVICE_MODESPRUNED = &H8000000   ' Dieses Gerät hat mehr Grafikmodes, als das Ausgabegerät unterstützt
  DISPLAY_DEVICE_PRIMARY_DEVICE = &H4      ' Das Gerät ist die Standardgrafikkarte
  DISPLAY_DEVICE_VGA_COMPATIBLE = &H10     ' Das Gerät ist VGA-kompatibel
End Enum

Private rc               As RECT
Private tmpL             As Long
Private tmpS             As String


Public Function RegisterCommandsDisplay()
  On Error GoTo eh
  ' Array(FunctionName, DisplayName, Category, Description, ArgName, ArgDescription...)
  commandMap.Add "changeresolution", Array("ChangeResolution", "Change Resolution", _
    MODULE_NAME, "Change the resolution to a supported display resolution (640x480 , 800x600 , 1024x768 ...).", _
    "Width", "Number of pixels horizontally.", _
    "Height", "Number of pixels vertically.", _
    "Monitor", "Index of the monitor. If nothing specified, the main monitor will be used.")
  commandMap.Add "getresolution", Array("GetResolution", "Get Resolution", _
    MODULE_NAME, "Retreive the current resolution of the main display.", _
    "Width", "Number of pixels horizontally.", _
    "Height", "Number of pixels vertically.", _
    "Monitor", "Index of the monitor. If nothing specified, the main monitor will be used.")

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".RegisterCommandsDisplay", Err.Number, Err.Source, Err.Description, Erl
End Function
Public Function PrepareExitCommandsDisplay()
  On Error GoTo eh

done:
  Exit Function
eh:
  RaiseError MODULE_NAME & ".PrepareExitCommandsDisplay", Err.Number, Err.Source, Err.Description, Erl
End Function


Public Function ChangeResolution(Optional ExecutingTroughApplicationRun As Boolean = False) As Boolean
  On Error GoTo eh
  If Not (IsNumber(CStr(currentRowArray(1, ColAArg1 + 0))) And IsNumber(CStr(currentRowArray(1, ColAArg1 + 1)))) Then
    ChangeResolution = False
    RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
      "Arguments need to be valid numbers: Arg1=[" & CStr(currentRowArray(1, ColAArg1 + 0)) & "] Arg2=[" & CStr(currentRowArray(1, ColAArg1 + 1)) & "]", Erl, 1, ExecutingTroughApplicationRun
    Exit Function
  End If
  
  Dim DevM As DEVMODE
  Dim Disp As DISPLAY_DEVICE
  Dim MonitorIndex As Long
  Dim TmpDevName As String
  Dim a As Boolean
  
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 2))) Then

    ' Read monitor index from array
    MonitorIndex = CLng(currentRowArray(1, ColAArg1 + 2))
    
    ' Initialize DISPLAY_DEVICE structure
    Disp.cb = Len(Disp)
    
    ' Get the device name for the specified monitor index
    If EnumDisplayDevices(ByVal 0&, MonitorIndex, Disp, 0&) <> 0 Then
      TmpDevName = Left$(Disp.DeviceName, InStr(1, Disp.DeviceName, vbNullChar) - 1)
      
      ' Initialize DEVMODE structure
      DevM.dmSize = Len(DevM)
      
      ' Get current settings for the selected monitor
      If EnumDisplaySettings(TmpDevName, ENUM_CURRENT_SETTINGS, DevM) <> 0 Then
          
          ' Update DEVMODE fields for new resolution
          DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
          DevM.dmPelsWidth = CLng(currentRowArray(1, ColAArg1 + 0))
          DevM.dmPelsHeight = CLng(currentRowArray(1, ColAArg1 + 1))
          
          ' Test the new resolution before applying; no user intervention required
          If ChangeDisplaySettingsEx(TmpDevName, DevM, 0&, CDS_TEST, 0&) = DISP_CHANGE_SUCCESSFUL Then
            ' Apply and save the new resolution
            ChangeDisplaySettingsEx TmpDevName, DevM, 0&, CDS_UPDATEREGISTRY, 0&
          Else
            ChangeResolution = False
            RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
              "Resolution change test failed for monitor : Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "].", Erl, 1, ExecutingTroughApplicationRun
            Exit Function
          End If
        Else
          ChangeResolution = False
          RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
            "Cannot retrieve current settings for monitor : Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "].", Erl, 2, ExecutingTroughApplicationRun
          Exit Function
        End If
    Else
        ChangeResolution = False
        RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
          "Monitor with index : Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "] does not exist.", Erl, 3, ExecutingTroughApplicationRun
        Exit Function
    End If
  
  Else
    tmpL = 0
    
    Do 'Enumerate settings
      a = EnumDisplaySettings(lpszDeviceName:=0&, iModeNum:=tmpL&, lpDevMode:=DevM)
      tmpL = tmpL + 1
    Loop Until (a = False)
    
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Change settings
    
    DevM.dmPelsWidth = currentRowArray(1, ColAArg1 + 0)
    DevM.dmPelsHeight = currentRowArray(1, ColAArg1 + 1)
    
    If ChangeDisplaySettings(DevM, CDS_TEST) <= 0 Then
      ChangeDisplaySettings DevM, CDS_UPDATEREGISTRY
    Else
      ChangeResolution = False
      RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, _
        "Resolution change test failed.", Erl, 4, ExecutingTroughApplicationRun
      Exit Function
    End If
  End If
  
done:
  ChangeResolution = True
  Exit Function
eh:
  ChangeResolution = False
  RaiseError MODULE_NAME & "ChangeResolution", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
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
  
  If IsNumber(CStr(currentRowArray(1, ColAArg1 + 2))) Then
  
    Dim DevM As DEVMODE
    Dim Disp As DISPLAY_DEVICE
    Dim MonitorIndex As Long
    Dim TmpDevName As String
    MonitorIndex = CLng(currentRowArray(1, ColAArg1 + 2))

   ' Initialize structures
    Disp.cb = Len(Disp)
    DevM.dmSize = Len(DevM)
    
    ' Take device for given index
    If EnumDisplayDevices(ByVal 0&, MonitorIndex, Disp, 0&) <> 0 Then
        TmpDevName = Left$(Disp.DeviceName, InStr(1, Disp.DeviceName, vbNullChar) - 1)
        
        ' Take current resolution
        If EnumDisplaySettings(TmpDevName, ENUM_CURRENT_SETTINGS, DevM) <> 0 Then
          currentRowRange(1, ColAArg1 + 0).Value = DevM.dmPelsWidth
          currentRowRange(1, ColAArg1 + 1).Value = DevM.dmPelsHeight
        Else
          GetResolution = False
          RaiseError MODULE_NAME & "GetResolution", Err.Number, Err.Source, _
            "Cannot retrieve current resolution for monitor : Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "].", Erl, 1, ExecutingTroughApplicationRun
          Exit Function
        End If
    Else
      GetResolution = False
      RaiseError MODULE_NAME & "GetResolution", Err.Number, Err.Source, _
        "Monitor with index : Arg3=[" & CStr(currentRowArray(1, ColAArg1 + 2)) & "] does not exist.", Erl, 2, ExecutingTroughApplicationRun
      Exit Function
    End If

  Else
    GetWindowRect GetDesktopWindow(), rc
    With rc
      currentRowRange(1, ColAArg1 + 0).Value = .Right - .Left
      currentRowRange(1, ColAArg1 + 1).Value = .Bottom - .Top
    End With
  End If
done:
  GetResolution = True
  Exit Function
eh:
  GetResolution = False
  RaiseError MODULE_NAME & "GetResolution", Err.Number, Err.Source, Err.Description, Erl, , ExecutingTroughApplicationRun
End Function

' Alle möglichen Auflösungen aller Grafikkarten ermitteln (Windows 98, ME, NT, 2000)
Private Sub Command2_Click()
  Dim retval As Long, Dev As DEVMODE, i As Long, j As Long, Disp As DISPLAY_DEVICE
  Dim TmpDevName As String ' Dient dazu den DeviceNamen ohne
  ' VBNullChar-Zeichen zwischenzuspeichern
  Dim TmpDevString As String ' Dient dazu den DeviceString ohne
  ' VBNullChar-Zeichen zwischenzuspeichern
 
  ' Falls Windows 95 läuft
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
 
    ' Für jede gefundene Grafikkarte die Bildschirmauflösungen ermitteln
    Do While EnumDisplaySettings(TmpDevName, j, Dev) <> 0
      Debug.Print "Bildschrimauflösung (" & CStr(j) & "): " & _
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
    MsgBox "Dieses Beispiel läuft nur ab Windows 98"
  End If
End Sub




Attribute VB_Name = "Watch"
Option Explicit

'    #If Win64 Then
'        Private Start As LongLong, Current As LongLong
'        Private Const MaxTimer As LongLong = 3600000 ' one hour, in ms
'        Private Declare PtrSafe Function GetTickCount Lib "kernel32" Alias "GetTickCount64" () As LongLong
'    #Else
Private Start As Long, Current As Long
Private Const MaxTimer As Long = 3600000 ' one hour, in ms
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
'    #End If


Public Function StartTimer()
  Start = GetTickCount
End Function

'#If Win64 Then
'Public Function EndTimer() As LongLong
'#Else
Public Function EndTimer() As Long
'#End If
  Current = GetTickCount
  EndTimer = Current - Start
  Start = Current
  If EndTimer > MaxTimer Then EndTimer = 0
End Function



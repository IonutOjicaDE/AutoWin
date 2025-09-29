Attribute VB_Name = "AutoWinDefaults"
Option Explicit
Private Const MODULE_NAME As String = "AutoWinDefaults"

' Integer   % Dim i%
' Long      & Dim i&
' LongLong  ^ Dim i^
' Single    ! Dim i!
' Double    # Dim i#
' Currency  @ Dim i@
' String    $ Dim i$

Public Const MinLongNumber As Long = -2147483648#
Public Const MaxLongNumber As Long = 2147483647#

Public Const NULL_         As LongPtr = 0&

' === Sheet Automation, CodeName ShAuto      ===
Public Const ColAStatus      As Long = 1&  ' A
Public Const ColACommand     As Long = 2&  ' B
Public Const ColAArg1        As Long = 3&  ' 3+0=C ... 3+9=12=L
Public Const ColAWindow      As Long = 13& ' M
Public Const ColAColor       As Long = 14& ' N
Public Const ColAPause       As Long = 15& ' O
Public Const ColAKeybd       As Long = 16& ' P
Public Const ColAonError     As Long = 17& ' Q
Public Const ColAComment     As Long = 18& ' R

' === Sheet KeyPress, CodeName shKey         ===
Public Const ColKeyName          As Long = 1& ' A
Public Const ColKeyCodeHex       As Long = 2& ' B
Public Const ColKeyCodeDec       As Long = 3& ' C
Public Const ColKeyDescription   As Long = 4& ' D
Public Const ColKeySendKeys      As Long = 5& ' E
Public Const ColKeyChar          As Long = 6& ' F
Public Const ColKeyPressed       As Long = 7& ' G

' === Characters to be used in Status column ===
Public Const StatusNOW   As String = ">"
Public Const StatusOK    As String = "+" ' alternative ChrW(9786) 'HEX 263A = DEC 9786
Public Const StatusNOK   As String = "!"
Public Const StatusSKIP  As String = "-"
Public Const StatusPause As String = "p"

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

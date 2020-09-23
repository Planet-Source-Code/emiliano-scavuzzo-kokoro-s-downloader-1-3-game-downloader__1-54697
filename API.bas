Attribute VB_Name = "API"
'**************************************************************************************
'This code is part of Kokoro's Downloader
'by Emiliano Scavuzzo <anshoku@yahoo.com>
'**************************************************************************************

Option Explicit

'API to create MenssageBox that doesn't stop the program
Public Declare Function MessageBox& Lib "user32" Alias "MessageBoxA" (ByVal hwnd&, ByVal lpText$, ByVal lpCaption$, ByVal wType&)
 'constants for MessageBox
  
  'Public Const MB_OK                As Long = &H0&
  'Public Const MB_OKCANCEL          As Long = &H1&
  'Public Const MB_ABORTRETRYIGNORE  As Long = &H2&
  'Public Const MB_YESNOCANCEL       As Long = &H3&
  Public Const MB_YESNO             As Long = &H4&
  'Public Const MB_RETRYCANCEL       As Long = &H5&
  
  Public Const MB_ICONERROR         As Long = &H10&
  Public Const MB_ICONQUESTION      As Long = &H20&
  Public Const MB_ICONEXCLAMATION   As Long = &H30&
  Public Const MB_ICONINFORMATION   As Long = &H40&
  
  'Public Const MB_DEFBUTTON1        As Long = &H0&
  Public Const MB_DEFBUTTON2        As Long = &H100&
  'Public Const MB_DEFBUTTON3        As Long = &H200&
  'Public Const MB_DEFBUTTON4        As Long = &H300&
  
  'Public Const MB_APPLMODAL         As Long = &H0&
  'Public Const MB_SYSTEMMODAL       As Long = &H1000&
  'Public Const MB_TASKMODAL         As Long = &H2000&
  
  'Public Const IDOK                 As Long = 1
  'Public Const IDCANCEL             As Long = 2
  'Public Const IDABORT              As Long = 3
  'Public Const IDRETRY              As Long = 4
  'Public Const IDIGNORE             As Long = 5
  Public Const IDYES                As Long = 6
  Public Const IDNO                 As Long = 7

'API to open a file with its default application
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'APIs to move the form
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'APIs to the label link
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'APIs to show the icon on the taskbar button for a
'borderless form and to draw the form

Public Const SC_MOVE = &HF010&
Public Const SC_SIZE = &HF000&
Public Const SC_MAXIMIZE = &HF030&
Public Const MF_BYCOMMAND = 0&

'Rectangle
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Point
Public Type POINTAPI
        x As Long
        y As Long
End Type

'Constants to minimize the window clicking on the little
'button and to play the minimization sound
Public Const SC_MINIMIZE = &HF020&
Public Const WM_SYSCOMMAND = &H112

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'API to change focus
Public Declare Function SetThisFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'API to see if and URL is valid
Public Declare Function IsValidURL Lib "URLMON.DLL" (ByVal pbc As Long, ByVal szURL As String, ByVal dwReserved As Long) As Long
Public Const S_OK = &H0

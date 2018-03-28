Attribute VB_Name = "mdlGlobe"
  Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Const GW_HWNDNEXT = 2
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 6

Public Const SW_SHOW = 5
Public Const VK_CONTROL = &H11
Public Const VK_ADD = &H6B
Public Const VK_F2 = &H71


  Public Const HWND_TOPMOST = -1
  Public Const SWP_SHOWWINDOW = &H40

Attribute VB_Name = "mdlGlobe"
  Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
  Public Const HWND_TOPMOST = -1
  Public Const SWP_SHOWWINDOW = &H40

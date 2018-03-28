Attribute VB_Name = "mdlMain"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'��Ϣ
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const WM_KEYDOWN = &H100

'�����С
Public Const WIN_NORMAL = 0
Public Const WIN_MIN = 1
Public Const WIN_MAX = 2

'�ַ���
Public Const STR_NULL = "Null"

'������
Public Const SB_LINEUP = 0
Public Const SB_LINELEFT = 0
Public Const SB_LINEDOWN = 1
Public Const SB_LINERIGHT = 1
Public Const SB_PAGEUP = 2
Public Const SB_PAGELEFT = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_PAGERIGHT = 3
Public Const SB_THUMBPOSITION = 4
Public Const SB_TOP = 6
Public Const SB_LEFT = 6
Public Const SB_BOTTOM = 7
Public Const SB_RIGHT = 7
Public Const SB_ENDSCROLL = 8

'�Զ���İ�������
Public Const KEY_ENTER = 13
Public Const KEY_SPACE = 32

Public Const KEY_LEFT = 37
Public Const KEY_UP = 38
Public Const KEY_RIGHT = 39
Public Const KEY_DOWN = 40

Public Const KEY_INSERT = 45
Public Const KEY_DEL = 46
Public Const KEY_HOME = 36
Public Const KEY_END = 35
Public Const KEY_UPPAGE = 33
Public Const KEY_DOWNPAGE = 34
Public Const KEY_PAUSE = 19

Public Const KEY_C = 67
Public Const KEY_D = 68
Public Const KEY_P = 80
Public Const KEY_Q = 81
Public Const KEY_V = 86

Public Const SHIFT_CTRL = 2
Public Const SHIFT_SHIFT = 1
Public Const SHIFT_ALT = 4
Public Const SHIFT_NONE = 0

'�Զ�����갴������
Public Const MOUSE_LEFT = 1
Public Const MOUSE_RIGHT = 2

'API�и��Ƶİ�������
Public Const VK_LEFT = 37
Public Const VK_UP = 38
Public Const VK_RIGHT = 39
Public Const VK_DOWN = 40

'�Զ��崰���˳�����
Public Const QUIT_TRUE = 0
Public Const QUIT_FALSE = 1

'�Զ��尴������
Public Const ASC_KEY_ENTER = 13
Public Const ASC_KEY_CTRL_ENTER = 10

'�Զ���״̬����
Public Enum STA_BUTTON
    STA_CANCEL = 0
    STA_OK = 1
End Enum


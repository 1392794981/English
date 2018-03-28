VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmCarefullListen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "细听"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3000
      Top             =   360
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2778
      _cy             =   1085
   End
   Begin VB.Label Label1 
      Caption         =   "×1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmCarefullListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileName As String
Public PositionFrom As Double
Public PositionTo As Double
Public rate As Double
Private Declare Function SetWindowPos Lib "user32" (ByVal HWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private WithEvents hk As clsRegHotKeys

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case KEY_UP
'            rate = rate + 0.1
'            Me.EnglishPlay
'        Case KEY_DOWN
'            rate = rate - 0.1
'            Me.EnglishPlay
'        Case KEY_LEFT
'            Me.EnglishPlay
'        Case KEY_RIGHT
'            Unload Me
'        Case 107, 229, 219, 221  '大键盘]键 '小键盘+键,大键盘[键 '小键盘+键
'            Me.EnglishPlay
'        Case KEY_DEL
'            Unload Me
'        Case KEY_ENTER
'            If WindowsMediaPlayer1.playState = wmppsPlaying Then
'                WindowsMediaPlayer1.Controls.pause
'            Else
'                WindowsMediaPlayer1.Controls.play
'            End If
'    End Select
'End Sub
'
'
'
'Public Sub EnglishPlay()
'    frmMain.mciRead.Notify = False
'    frmMain.mciRead.Wait = True
'    frmMain.mciRead.Command = "stop"
'    'frmMain.mciRead.Command = "close"
'    WindowsMediaPlayer1.URL = Me.FileName
'    WindowsMediaPlayer1.Controls.currentPosition = Me.PositionFrom / 1000
'    WindowsMediaPlayer1.settings.rate = Me.rate
'    Label1.Caption = "×" + Trim(Format(Me.rate, "0.0"))
'    WindowsMediaPlayer1.Controls.play
'End Sub

Private Sub Form_Load()
    SetWindowPos Me.HWnd, -1, 0, 0, 0, 0, 2 Or 1
'    rate = 0.5
'    ProgressBar1.Value = ProgressBar1.Min
'    DoEvents
'    EnglishPlay
'
'
'    '''''''''''''''''''''''''''''''''''''''''''''''
'
'    Set hk = New clsRegHotKeys
'
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyDown, "Alt_Down"
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyUp, "Alt_Up"
'
'
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyAdd, "Alt_Add"
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeySubtract, "Alt_Subtract"
'
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyDivide, "Alt_Divide"
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyMultiply, "Alt_Multiply"
'
'    hk.RegHotKeys Me.hwnd, ShiftKeys.altKey, vbKeyReturn, "Alt_Enter"
'
'
'    Me.Show   '这个不能省略，否则窗体无法显示出来！
'
'    hk.WaitMsg
'    ''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

'Private Sub hk_HotKeysDown(Key As String)
'    Select Case Key
'        Case "Alt_Down"
'            rate = rate + 0.1
'            Me.EnglishPlay
'        Case "Alt_Up"
'            rate = rate - 0.1
'            Me.EnglishPlay
'        Case "Alt_Divide"
'           ' Call JumpPlay(mciRead.Position - 5000)
'        Case "Alt_Multiply"
'            'Call JumpPlay(mciRead.Position + 5000)
'        Case "Alt_Subtract"
'            Me.EnglishPlay
'        Case "Alt_Enter"
'            Unload Me
'        Case "Alt_Add"
'            If WindowsMediaPlayer1.playState = wmppsPlaying Then
'                WindowsMediaPlayer1.Controls.pause
'            Else
'                WindowsMediaPlayer1.Controls.play
'            End If
'    End Select
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    'frmMain.blShowCarefullListenForm = False
End Sub

'Private Sub Timer1_Timer()
'    On Error Resume Next
'    If WindowsMediaPlayer1.Controls.currentPosition * 1000 >= (Me.PositionTo) Then
'        WindowsMediaPlayer1.Controls.pause
'    End If
'    ProgressBar1.Min = 0
'    ProgressBar1.Max = Int(PositionTo) - Int(PositionFrom)
'    ProgressBar1.Value = Int(WindowsMediaPlayer1.Controls.currentPosition * 1000) - Int(PositionFrom)
'End Sub

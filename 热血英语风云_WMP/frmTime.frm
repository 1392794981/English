VERSION 5.00
Begin VB.Form frmTime 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrMoveForm 
      Interval        =   10
      Left            =   1560
      Top             =   0
   End
   Begin VB.Label lblBeginTime 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   80
      Width           =   1455
   End
   Begin VB.Label lblTimeShow 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   40
      Width           =   855
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SWP_SHOWWINDOW = &H40
Private Const HWND_TOPMOST = -1

Dim typMousePosition As POINTAPI
Dim tmrBeginTime

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'frmMain.txtTRY.SetFocus
End Sub

Private Sub Form_Load()
    tmrBeginTime = Time
    lblBeginTime.Caption = Format(tmrBeginTime, "hh:mm:ss AMPM")
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, (Me.Width) \ 15, (Me.Height) \ 15, SWP_SHOWWINDOW
End Sub

Private Sub tmrMoveForm_Timer()
    '''''''''''''''''计时功能'''''''''''''''''''''
    lblTimeShow.Caption = CStr((Hour(Time) - Hour(tmrBeginTime)) * 60 + (Minute(Time) - Minute(tmrBeginTime))) + "分"
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''窗体动画'''''''''''''''''''''
    Static lngFormPositionY As Long
    Static lngStayTime As Long
    Static blNotMoved As Boolean

    Dim lngValidY As Long
    Dim lngMoveSpeed As Long
    Dim lngStaySecond As Long
    lngValidY = 20
    lngMoveSpeed = 2
    lngStaySecond = 0.1 '鼠标离开后，窗体停留时间

    DoEvents
    GetCursorPos typMousePosition
    If typMousePosition.x < ((Me.Width) \ 15) And typMousePosition.y < lngValidY Then
        If lngFormPositionY < 0 Then
            lngFormPositionY = lngFormPositionY + lngMoveSpeed
            SetWindowPos Me.hwnd, HWND_TOPMOST, 0, lngFormPositionY, (Me.Width) \ 15, (Me.Height) \ 15, SWP_SHOWWINDOW
        Else
            blNotMoved = False '已经处于显示位置
        End If
        lngStayTime = 0
    Else
        ''''''停留时间'''''''
        lngStayTime = lngStayTime + 1
        '''''''''''''''''''''
        If ((lngFormPositionY > -((Me.Height) \ 15) And lngFormPositionY < 0)) Or _
            (blNotMoved = False And lngStayTime > ((1000 / Me.tmrMoveForm.Interval) * lngStaySecond)) Then
            If lngFormPositionY > -((Me.Height) \ 15) Then
                lngFormPositionY = lngFormPositionY - lngMoveSpeed
                SetWindowPos Me.hwnd, HWND_TOPMOST, 0, lngFormPositionY, (Me.Width) \ 15, (Me.Height) \ 15, SWP_SHOWWINDOW
            End If
            blNotMoved = True
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''
End Sub

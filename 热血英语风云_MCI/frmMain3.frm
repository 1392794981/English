VERSION 5.00
Begin VB.Form frmMain3 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   510
   ClientLeft      =   7020
   ClientTop       =   465
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "隐"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   70
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   240
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "计时"
      Height          =   370
      Left            =   5040
      TabIndex        =   1
      Top             =   10
      Width           =   1065
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   370
      Left            =   6240
      TabIndex        =   0
      Top             =   10
      Width           =   1065
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   70
      Width           =   2655
   End
   Begin VB.Label lblTime 
      Caption         =   "欢迎使用！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   70
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private blST   As Boolean
Private StartTime As Date
Private i As Long

Private Sub Check1_Click()
   If Check1.Value = Checked Then
    lblTime.Visible = False
    Label1.Visible = False
   Else
    lblTime.Visible = True
    Label1.Visible = True
   End If
End Sub

Private Sub cmdQuit_Click()
    If MsgBox("确定？", vbOKCancel) = vbOK Then
        Unload Me
    End If
End Sub

Private Sub cmdTime_Click()
    If MsgBox("确定？", vbOKCancel) = vbOK Then
        StartTime = Now
        i = 0
        blST = True
        cmdTime.Caption = "进行中..."
    End If
End Sub

Private Sub Form_Load()
    Check1.Value = Checked
    blST = False
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 488, 26, SWP_SHOWWINDOW
End Sub

Private Sub Timer1_Timer()
    
    
    If blST = False Then Exit Sub
    
    i = i + 1
    Label1.Caption = StartTime
    lblTime.Caption = Str(i \ 60) + "分"
End Sub

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAllResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "&比较结果"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer tmrQuickCompare 
      Interval        =   5
      Left            =   4800
      Top             =   6960
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   6990
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "返回"
      Height          =   375
      Left            =   8250
      TabIndex        =   0
      Top             =   6990
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtA 
      Height          =   3435
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6059
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAllResult.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtB 
      Height          =   3435
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3525
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6059
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAllResult.frx":009A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAllResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdQuit_Click()
    Unload Me
    
    Unload frmMain2
    
    frmMain.txtTRY.SetFocus
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Unload Me
End Sub

Private Sub Form_Load()
    txtA.Text = gstrA
    txtB.Text = gstrB
    
    Dim i As Long
    
    For i = 1 To UBound(StartLenA)
        txtA.SelStart = Max(StartLenA(i).tStart - 1, 1)
        txtA.SelLength = StartLenA(i).tLen
        txtA.SelBold = True
        txtA.SelColor = &HFF0000
    Next
    
    txtB.SelStart = 0
    txtB.SelLength = Len(txtB.Text)
    txtB.SelColor = &HFF
    For i = 1 To UBound(StartLenB)
        txtB.SelStart = Max(StartLenB(i).tStart - 1, 1)
        txtB.SelLength = StartLenB(i).tLen
        txtB.SelBold = True
        txtB.SelColor = &HFF0000
    Next
    
    SetWindowPos Me.hwnd, HWND_TOPMOST, 174, 45, (Me.Width) \ 15, (Me.Height) \ 15, 0
End Sub

Private Sub tmrQuickCompare_Timer()
            DoEvents
        tmrQuickCompare.Enabled = False
        DoEvents

        DoEvents
'        SendMessage txtA.hwnd, WM_HSCROLL, SB_BOTTOM, 0
'        SendMessage txtB.hwnd, WM_HSCROLL, SB_BOTTOM, 0
        If IsQuickCompare = True Then
        
            cmdQuit.SetFocus
        Else
            Command1.SetFocus
        End If
        
        DoEvents
End Sub

Private Sub txtA_KeyDown(KeyCode As Integer, Shift As Integer)
'    Unload Me
End Sub

Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
'    Unload Me
End Sub

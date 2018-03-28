VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtNewWord 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000315BA&
      Height          =   7470
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 45, (Me.Width) \ 15, (Me.Height) \ 15, 0
End Sub


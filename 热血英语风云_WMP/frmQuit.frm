VERSION 5.00
Begin VB.Form frmQuit 
   BackColor       =   &H80000003&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3285
   ClientLeft      =   2970
   ClientTop       =   2355
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOpenRecordFile 
      Caption         =   "�򿪼�¼�ļ�"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "���¼��ϰʱ�䣡"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1320
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "frmQuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public blQuitCancel As Boolean
Private Sub cmdCancel_Click()
    blQuitCancel = True
    Unload Me
End Sub

Private Sub cmdOpenRecordFile_Click()
    On Error Resume Next

    Dim strEnTxtFileName As String
    strEnTxtFileName = "F:\����\�ĺ��2008�ꡰӢ��֮�Ρ�ͳ�Ʊ�.doc"
    
    Dim fso As New FileSystemObject
    
    If fso.FileExists(strEnTxtFileName) = False Then
        MsgBox "û���ҵ���¼!"
        Exit Sub
    End If
    
    
'    Dim wrdApp As Word.Application
'    Dim wrdDoc As Word.Document
'
'    Set wrdApp = CreateObject("Word.Application")
'    Set wrdDoc = wrdApp.Documents.Open(strEnTxtFileName)
'
'    wrdApp.Visible = True
End Sub

Private Sub cmdQuit_Click()

    Unload Me
End Sub


Private Sub Form_Load()
    blQuitCancel = False
End Sub

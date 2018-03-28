VERSION 5.00
Begin VB.Form frmDic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "字典"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtResult 
      Height          =   2535
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   480
      Width           =   2420
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查询"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   50
      Width           =   735
   End
   Begin VB.TextBox txtWord 
      Height          =   375
      Left            =   50
      TabIndex        =   0
      Top             =   50
      Width           =   1575
   End
End
Attribute VB_Name = "frmDic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type typWord
    strWord As String
    strMean As String
End Type
Private myWord() As typWord

Private Sub cmdSearch_Click()
    txtResult.Text = ""
    Dim i, lngFound As Long
    lngFound = 0
    For i = LBound(myWord) To UBound(myWord)
        If myWord(i).strWord Like Trim(txtWord.Text) Then
            txtResult.Text = txtResult.Text + myWord(i).strWord + vbCrLf
            txtResult.Text = txtResult.Text + myWord(i).strMean + vbCrLf + vbCrLf
            lngFound = lngFound + 1
            If lngFound >= 100 Then
                txtResult.Text = "找到太多单词，仅列出前100个……" + vbCrLf + vbCrLf + txtResult.Text
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
    'On Error Resume Next
  
    Dim fso As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    ReDim myWord(0)
    
'    For Each fil In fso.GetFolder(App.Path + "\TXT版牛津字典\").Files
'        Set ts = fil.OpenAsTextStream
'        Do Until ts.AtEndOfStream
'            ReDim Preserve myWord(UBound(myWord) + 1)
'            myWord(UBound(myWord)).strWord = Trim(ts.ReadLine)
'            myWord(UBound(myWord)).strMean = Trim(ts.ReadLine)
'        Loop
'        ts.Close
'    Next
    
    Set ts = fso.GetFile(App.Path + "\字典\考研字典.txt").OpenAsTextStream
    Do Until ts.AtEndOfStream
        ReDim Preserve myWord(UBound(myWord) + 1)
        myWord(UBound(myWord)).strWord = Trim(ts.ReadLine)
        myWord(UBound(myWord)).strMean = Trim(ts.ReadLine)
    Loop
    ts.Close
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 45, (Me.Width) \ 15, (Me.Height) \ 15, 0
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 480, 0, (Me.Width) \ 15, (Me.Height) \ 15, 0
    
    IsfrmDicExist = True
End Sub

Private Sub Form_Resize()
    txtResult.Height = Me.ScaleHeight - txtResult.Top - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    IsfrmDicExist = False
End Sub

Private Sub txtWord_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KEY_ENTER Then
        Call cmdSearch_Click
    End If
End Sub

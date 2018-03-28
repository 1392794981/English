VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "英文对比"
   ClientHeight    =   6675
   ClientLeft      =   2835
   ClientTop       =   1170
   ClientWidth     =   9270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrQuickCompare 
      Interval        =   400
      Left            =   4200
      Top             =   6120
   End
   Begin VB.CommandButton cmdMySelf 
      Caption         =   "自编"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "退出"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "对比"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox txtB 
      Height          =   6135
      Left            =   4560
      TabIndex        =   4
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10821
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain2.frx":0000
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
   Begin RichTextLib.RichTextBox txtA 
      Height          =   6135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   10821
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain2.frx":009A
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
   Begin VB.Label lblResult 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   6240
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strInformation As String

Public strInitA, strInitB As String

Private Sub cmdCompare_Click()
    
    strInformation = ""
    
    Dim DifCount As Long
    
    DifCount = myCompare
    
    lblResult.Caption = "经过比较，有" + CStr(DifCount) + "处不同！" + Chr(13) + strInformation
    
'    Dim i As Long
'    For i = 1 To UBound(StartLenA)
'        MsgBox CStr(i) + "  " + CStr(StartLenA(i).tStart)
'    Next
    
    frmAllResult.Caption = "有" + CStr(DifCount) + "处不同！" + strInformation
    frmAllResult.Show 1
End Sub

Private Function myCompare() As Long
    'If Len(Trim(txtA.Text)) = 0 Or Len(Trim(txtB.Text)) = 0 Then Exit Function
    
    Dim strA, strB As String
    Dim intA, intB, strLen As Long
    Dim i, j As Long
    Dim intDifCount As Long
    
    intDifCount = 0
    strA = txtA.Text
    strB = txtB.Text
    
    strA = Replace(strA, Chr(8), " ")
    strA = Replace(strA, Chr(9), " ")
    strA = Replace(strA, Chr(10), " ")
    strA = Replace(strA, Chr(13), " ")
    strA = Replace(strA, vbCrLf, " ")
    
    strA = Replace(strA, " '", "'")
    strA = Replace(strA, " .", ".")
    strA = Replace(strA, " ?", "?")
    strA = Replace(strA, " !", "!")
    strA = Replace(strA, " ,", ",")
    
    For i = 1 To 100
        strA = Replace(strA, "  ", " ")
    Next
    
    strB = Replace(strB, Chr(8), " ")
    strB = Replace(strB, Chr(9), " ")
    strB = Replace(strB, Chr(10), " ")
    strB = Replace(strB, Chr(13), " ")
    strB = Replace(strB, vbCrLf, " ")
    
    
    strB = Replace(strB, " '", "'")
    strB = Replace(strB, " .", ".")
    strB = Replace(strB, " ?", "?")
    strB = Replace(strB, " !", "!")
    strB = Replace(strB, " ,", ",")
    
    For i = 1 To 100
        strB = Replace(strB, "  ", " ")
    Next
    
    txtA.Text = strA
    txtB.Text = strB
    'MsgBox strA
    strLen = Min(Len(strA), Len(strB))
    intA = 1
    intB = 1
    
    ReDim StartLenA(0)
    ReDim StartLenB(0)
    gstrA = strA
    gstrB = strB
    
    Do While intA < strLen And intB < strLen
        If Mid(strA, intA, 1) <> Mid(strB, intB, 1) Then
            If Mid(strA, intA, 1) = """" Then
'                Or Mid(strA, intA, 1) = "." Or Mid(strA, intA, 1) = "'" _
'                Or Mid(strA, intA, 1) = "-" Or Mid(strA, intA, 1) = "," _

                intA = intA + 1
            ElseIf Mid(strB, intB, 1) = """" Then
'                Or Mid(strB, intB, 1) = "." Or Mid(strB, intB, 1) = "'" _
'                Or Mid(strB, intB, 1) = "-" Or Mid(strB, intB, 1) = "," _

                intB = intB + 1
                
'''''''''''''不区分大小写'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            ElseIf Asc(Mid(strA, intA, 1)) >= 65 And _
'                   Asc(Mid(strA, intA, 1)) <= 90 And _
'                   Asc(Mid(strA, intA, 1)) + 32 = Asc(Mid(strB, intB, 1)) Then
'                intA = intA + 1
'                intB = intB + 1
'            ElseIf Asc(Mid(strB, intB, 1)) >= 65 And _
'                   Asc(Mid(strB, intB, 1)) <= 90 And _
'                   Asc(Mid(strB, intB, 1)) + 32 = Asc(Mid(strA, intA, 1)) Then
'                intA = intA + 1
'                intB = intB + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Else
                intDifCount = intDifCount + 1
                For i = 1 To 40
                    For j = 1 To 40
                        If Mid(strA, intA + i, 11) = Mid(strB, intB + j, 11) Then
                            
                            frmResult.SetPara strA, strB, intA, intB, i, j, strLen
                            frmResult.Show 1
                            
                            intA = intA + i
                            intB = intB + j
                            
                            'MsgBox CStr(UBound(StartLenA)) + "   " + CStr(StartLenA(UBound(StartLenA)).tStart)
                            GoTo OutHandle
                        End If
                    Next
                Next
                frmResult.SetPara strA, strB, intA, intB, 1, 1, strLen
                frmResult.Show 1
                myCompare = intDifCount
                strInformation = "可能有太多不同了！所以没法比较下去！"
                MsgBox strInformation
                Exit Function
            End If
        Else
            intA = intA + 1
            intB = intB + 1
        End If
OutHandle:
    Loop
    
    myCompare = intDifCount
End Function

Private Sub cmdMySelf_Click()
    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    If fso.FileExists(App.Path + "\第四册（自编）\" + Trim(strCurrentMediaFileName) + ".txt") Then
        Set ts = fso.OpenTextFile(App.Path + "\第四册（自编）\" + Trim(strCurrentMediaFileName) + ".txt", ForReading)
        txtA.Text = ts.ReadAll
        ts.Close
    Else
        MsgBox "没有自编文件！"
    End If
    
End Sub

Private Sub cmdQuit_Click()
    Unload Me
    frmMain.txtTRY.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 35 Then
        Call cmdCompare_Click
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    DoEvents
    'Call cmdCompare_Click
    
    SetWindowPos Me.hwnd, 0, 174, 45, (Me.Width) \ 15, (Me.Height) \ 15, 0
End Sub


Private Sub tmrQuickCompare_Timer()
    On Error Resume Next

    If IsQuickCompare = True Then
        DoEvents
        tmrQuickCompare.Enabled = False
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        Call cmdCompare_Click
    End If
End Sub

Private Sub txtA_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 35 Then
        Call cmdCompare_Click
        KeyCode = 0
    End If
End Sub

Private Sub txtB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 35 Then
        Call cmdCompare_Click
        KeyCode = 0
    End If
End Sub

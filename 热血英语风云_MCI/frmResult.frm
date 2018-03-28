VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmResult 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "结果"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrQuickCompare 
      Interval        =   50
      Left            =   4800
      Top             =   4440
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "继续"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox txtA 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3836
      _Version        =   393217
      TextRTF         =   $"frmResult.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtB 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3836
      _Version        =   393217
      TextRTF         =   $"frmResult.frx":009A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strA, strB As String
Public intA, intB, intLenA, intLenB, strLen As Long

Private Sub cmdContinue_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strShowA, strShowB As String
    strShowA = Mid(strA, Max(1, intA - 20), Min(intA + 20, strLen))
    strShowB = Mid(strB, Max(1, intB - 20), Min(intB + 20, strLen))
    txtA.Text = strShowA
    txtB.Text = strShowB
    
    Dim intShowA, intShowB, intTemp As Long
    intShowA = intA - Max(1, intA - 20) + 1
    intShowB = intB - Max(1, intB - 20) + 1
    
    intTemp = intShowA + intLenA
    Do While intShowA > 1
        If Mid(strShowA, intShowA, 1) = " " Then Exit Do
        intShowA = intShowA - 1
    Loop
    Do While intTemp < Len(strShowA)
        If Mid(strShowA, intTemp, 1) = " " Then Exit Do
        intTemp = intTemp + 1
    Loop
    txtA.SelStart = intShowA
    txtA.SelLength = intTemp - intShowA
    txtA.SelBold = True
    txtA.SelColor = &HFF0000
    
    ReDim Preserve StartLenA(UBound(StartLenA) + 1)
    StartLenA(UBound(StartLenA)).tStart = intShowA + Max(1, intA - 20)
    StartLenA(UBound(StartLenA)).tLen = intTemp - intShowA
    
    '=================================================
    txtB.SelStart = 0
    txtB.SelLength = Len(txtB.Text)
    txtB.SelColor = &HFF
    
    intTemp = intShowB + intLenB
    Do While intShowB > 1
        If Mid(strShowB, intShowB, 1) = " " Then Exit Do
        intShowB = intShowB - 1
    Loop
    Do While intTemp < Len(strShowB)
        If Mid(strShowB, intTemp, 1) = " " Then Exit Do
        intTemp = intTemp + 1
    Loop
    txtB.SelStart = intShowB
    txtB.SelLength = intTemp - intShowB
    txtB.SelBold = True
    txtB.SelColor = &HFF0000
    
    ReDim Preserve StartLenB(UBound(StartLenB) + 1)
    StartLenB(UBound(StartLenB)).tStart = intShowB + Max(1, intB - 20)
    StartLenB(UBound(StartLenB)).tLen = intTemp - intShowB
    
End Sub

Public Sub SetPara(vstrA, vstrB As String, vintA, vintB, vLenA, vLenB, vstrLen As Long)
    strA = vstrA
    strB = vstrB
    intA = vintA
    intB = vintB
    intLenA = vLenA
    intLenB = vLenB
    strLen = vstrLen
End Sub

Private Sub tmrQuickCompare_Timer()
    If IsQuickCompare = True Then
        DoEvents
        tmrQuickCompare.Enabled = False
        Unload Me
    End If
End Sub
